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

# Copy benchmark scripts
COPY scripts/ scripts/
COPY README.md .

# Default: run speed benchmark
CMD ["python3", "scripts/03_bench_speed.py"]
