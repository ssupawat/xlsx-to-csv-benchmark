FROM python:3.12-slim

# Install LibreOffice
RUN apt-get update && apt-get install -y \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy benchmark scripts and run-all script
COPY scripts/ scripts/
COPY run-all.sh .
RUN chmod +x run-all.sh
COPY README.md .

# Default: run all benchmarks
CMD ["./run-all.sh"]
