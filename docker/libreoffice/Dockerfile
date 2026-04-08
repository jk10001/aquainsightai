# Smaller base (still Debian/bookworm repos): ~130MB
FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive

# Install minimal LibreOffice bits + UNO bridge + fonts
# Note: no Java here. If LO complains but still runs, you win ~200–300MB.
RUN apt-get update -o Acquire::Retries=3 && \
    apt-get install -y --no-install-recommends \
      libreoffice-writer \
      python3-uno \
      fonts-liberation fonts-crosextra-caladea fonts-crosextra-carlito \
      ca-certificates procps curl bash coreutils && \
    # remove help files, manpages, docs, and unused locales to shrink image
    rm -rf /usr/share/doc/* /usr/share/man/* /usr/share/info/* && \
    find /usr/share/locale -mindepth 1 -maxdepth 1 ! -name 'en*' -exec rm -rf {} + && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# (Optional) in-container Python deps
COPY requirements-docker.txt /tmp/requirements-docker.txt
RUN pip install --no-cache-dir -r /tmp/requirements-docker.txt

# Scripts
COPY update_and_export.py /app/update_and_export.py
COPY start-lo.sh /usr/local/bin/start-lo.sh
RUN chmod +x /usr/local/bin/start-lo.sh

HEALTHCHECK --interval=5s --timeout=3s --retries=30 \
  CMD bash -lc "exec 3<>/dev/tcp/127.0.0.1/2002 || exit 1"

EXPOSE 2002
ENTRYPOINT ["/usr/local/bin/start-lo.sh"]
