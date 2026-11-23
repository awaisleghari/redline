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
import subprocess
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

MAX_BYTES = 30 * 1024 * 1024      # a little above the page's 25 MB upload limit
TIMEOUT_SECONDS = 120
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
# LibreOffice is heavy; convert at most two documents at a time.
SLOTS = threading.BoundedSemaphore(2)


def convert(docx: bytes) -> bytes:
    """Convert in a private temporary folder and delete it afterwards, always."""
    folder = tempfile.mkdtemp(prefix="c-")
    try:
        source = os.path.join(folder, "in.docx")
        with open(source, "wb") as handle:
            handle.write(docx)
        subprocess.run(
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
            timeout=TIMEOUT_SECONDS, check=True, cwd=folder,
            env={"HOME": folder, "PATH": os.environ.get("PATH", "/usr/bin:/bin"), "LANG": "C.UTF-8"},
        )
        with open(os.path.join(folder, "in.pdf"), "rb") as handle:
            pdf = handle.read()
        if not pdf.startswith(b"%PDF-"):
            raise RuntimeError("no PDF produced")
        return pdf
    finally:
        shutil.rmtree(folder, ignore_errors=True)


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
        self.end_headers()
        self.wfile.write(body)

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
        docx = self.rfile.read(length)
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


if __name__ == "__main__":
    ThreadingHTTPServer(("0.0.0.0", int(os.environ.get("PORT", "8080"))), Handler).serve_forever()
