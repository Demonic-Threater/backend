FROM python:3.11-slim

RUN apt-get update && apt-get install -y libreoffice && apt-get clean

WORKDIR /app

COPY . .

RUN pip install --no-cache-dir -r requirements.txt

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080"]


