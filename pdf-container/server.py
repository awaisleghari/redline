"""
PDF converter: POST a .docx to /convert, get the PDF back.

Privacy rules (PLAN.md, "PDF step and privacy"):
  - each request works in its own temporary folder, which is deleted before the
    reply is sent, whether the conversion worked or not;
  - nothing is logged: no file names, no contents, no sizes, no request lines;
  - the file is never given a name derived from the request.
"""
import os
import shutil
import signal
import socket
import subprocess
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

MAX_BYTES = 30 * 1024 * 1024      # a little above the page's 25 MB upload limit
TIMEOUT_SECONDS = float(os.environ.get("CONVERT_TIMEOUT_SECONDS", "120"))

# LibreOffice settings written into each request's throwaway profile: never
# update links, and never load a linked picture or file from anywhere (a
# document must not be able to pull a file from this machine, or from the
# internet, into its PDF).
PROFILE_SETTINGS = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Field" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Chart" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
"""
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
# LibreOffice is heavy; convert at most two documents at a time.
SLOTS = threading.BoundedSemaphore(2)


def convert(docx: bytes) -> bytes:
    """Convert in a private temporary folder and delete it afterwards, always."""
    folder = tempfile.mkdtemp(prefix="c-")
    process = None
    try:
        source = os.path.join(folder, "in.docx")
        with open(source, "wb") as handle:
            handle.write(docx)
        os.makedirs(os.path.join(folder, "profile", "user"))
        with open(os.path.join(folder, "profile", "user", "registrymodifications.xcu"), "w", encoding="utf-8") as handle:
            handle.write(PROFILE_SETTINGS)
        process = subprocess.Popen(
            [
                "soffice",
                "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nofirststartwizard",
                # A profile of its own inside the temporary folder: nothing is remembered between requests.
                f"-env:UserInstallation=file://{folder}/profile",
                # Read the file strictly as a Word document, whatever it claims to be.
                "--infilter=MS Word 2007 XML",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", folder,
                source,
            ],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            cwd=folder,
            # Everything LibreOffice writes, including its own scratch files, goes inside the folder.
            env={"HOME": folder, "TMPDIR": folder, "TMP": folder, "TEMP": folder, "XDG_RUNTIME_DIR": folder,
                 "PATH": os.environ.get("PATH", "/usr/bin:/bin"), "LANG": "C.UTF-8"},
            # Its own process group, so that every process it starts can be stopped with it.
            start_new_session=True,
        )
        if process.wait(timeout=TIMEOUT_SECONDS) != 0:
            raise RuntimeError("conversion failed")
        with open(os.path.join(folder, "in.pdf"), "rb") as handle:
            pdf = handle.read()
        if not pdf.startswith(b"%PDF-"):
            raise RuntimeError("no PDF produced")
        return pdf
    finally:
        # Stop LibreOffice and everything it started before deleting, so nothing can write the folder back.
        if process is not None:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                pass
            try:
                process.wait(timeout=10)
            except Exception:
                pass
            # Wait until every process of that group has really gone (they hold files and a socket open until then).
            deadline = time.monotonic() + 10
            while time.monotonic() < deadline:
                try:
                    os.killpg(process.pid, 0)
                except (ProcessLookupError, PermissionError):
                    break
                time.sleep(0.05)
        shutil.rmtree(folder, ignore_errors=True)
        remove_dead_sockets()


def remove_dead_sockets() -> None:
    """
    LibreOffice keeps a small socket file in /tmp while it runs and removes it
    when it finishes normally. One that was stopped leaves its socket behind.
    Delete every such socket that nothing is listening on any more (a socket
    still in use belongs to a conversion that is running, and is left alone).
    """
    try:
        names = [name for name in os.listdir(tempfile.gettempdir()) if name.startswith("OSL_PIPE_")]
    except OSError:
        return
    for name in names:
        path = os.path.join(tempfile.gettempdir(), name)
        probe = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        probe.settimeout(1)
        try:
            probe.connect(path)
        except (ConnectionRefusedError, FileNotFoundError):
            try:
                os.unlink(path)
            except OSError:
                pass
        except OSError:
            pass
        finally:
            probe.close()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "converter"
    sys_version = ""

    def log_message(self, *_args):  # nothing is logged, ever
        pass

    def reply(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        try:
            self.end_headers()
            self.wfile.write(body)
        except OSError:  # the caller went away; there is nobody to tell
            self.close_connection = True

    def refuse(self, status: int, message: str) -> None:
        self.close_connection = True
        self.reply(status, message.encode(), "text/plain; charset=utf-8")

    def do_GET(self):  # noqa: N802
        if self.path == "/health":
            self.reply(200, b"ok", "text/plain; charset=utf-8")
        else:
            self.refuse(404, "Not found.")

    def do_POST(self):  # noqa: N802
        if self.path != "/convert":
            return self.refuse(404, "Not found.")
        try:
            length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            return self.refuse(411, "The size of the file must be stated.")
        if length <= 0:
            return self.refuse(400, "No file was sent.")
        if length > MAX_BYTES:
            # Read past a moderately oversized upload (without keeping it) so that the
            # refusal reaches the sender; anything far larger is simply cut off.
            if length <= 4 * MAX_BYTES:
                remaining = length
                while remaining > 0:
                    chunk = self.rfile.read(min(remaining, 1024 * 1024))
                    if not chunk:
                        break
                    remaining -= len(chunk)
            return self.refuse(413, "The file is too large to convert.")
        if self.headers.get("Content-Type", "").split(";")[0].strip() != DOCX:
            return self.refuse(415, "Only Word documents (.docx) can be converted.")
        try:
            docx = self.rfile.read(length)
        except OSError:
            self.close_connection = True
            return
        if len(docx) != length or not docx.startswith(b"PK"):
            return self.refuse(400, "The file is not a Word document.")
        if not SLOTS.acquire(timeout=60):
            return self.refuse(503, "The converter is busy. Please try again in a moment.")
        try:
            pdf = convert(docx)
        except Exception:  # the reason is deliberately not reported or logged: it could describe the document
            return self.refuse(422, "The document could not be converted to PDF.")
        finally:
            SLOTS.release()
        self.reply(200, pdf, "application/pdf")


class QuietServer(ThreadingHTTPServer):
    daemon_threads = True

    def handle_error(self, request, client_address):  # noqa: ARG002
        # The default prints a traceback (with the caller's address) when a connection
        # drops. Nothing is logged here, ever.
        pass


if __name__ == "__main__":
    QuietServer(("0.0.0.0", int(os.environ.get("PORT", "8080"))), Handler).serve_forever()
