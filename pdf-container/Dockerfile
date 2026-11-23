# The PDF converter: LibreOffice (headless) behind a tiny HTTP server.
# It receives a static redline (.docx), returns it as a PDF, and keeps nothing.
#
# Ubuntu base on purpose: the build environment can reach Ubuntu's package
# mirrors but not Debian's (see PLAN.md, "Container base image").
FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive
# LibreOffice Writer without its desktop interface, Python for the server, and
# fonts with the same metrics as the ones Word documents usually ask for, so
# that lines and pages break where they do in Word:
#   Liberation  -> Times New Roman, Arial, Courier New
#   Carlito     -> Calibri
#   Caladea     -> Cambria
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-writer-nogui \
      python3 tini \
      fonts-liberation fonts-liberation2 fonts-crosextra-carlito fonts-crosextra-caladea fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/* \
 && fc-cache -f

# An unprivileged account: the converter never runs as root.
RUN useradd --system --create-home --home-dir /home/converter --shell /usr/sbin/nologin converter
COPY --chown=root:root server.py /opt/converter/server.py
USER converter
WORKDIR /home/converter
ENV HOME=/home/converter PORT=8080
EXPOSE 8080
# tini is a minimal init: it collects the remains of any LibreOffice process that had to be stopped.
ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["python3", "-u", "/opt/converter/server.py"]
