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

# 6. Expose port for Railway
EXPOSE 8000

# 7. Start the server
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
