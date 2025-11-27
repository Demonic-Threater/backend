# 1. Start from Python slim image
FROM python:3.12-slim

# 2. Install system dependencies and LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice unzip curl && \
    rm -rf /var/lib/apt/lists/*

# 3. Set working directory
WORKDIR /app

# 4. Copy everything to container
COPY . /app

# 5. Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# 6. Expose dynamic port (Railway ignores EXPOSE but good practice)
EXPOSE 8000

# 7. Start the server using environment variable for port
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
