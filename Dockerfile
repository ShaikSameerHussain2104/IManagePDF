FROM python:3.11-slim

# Install LibreOffice and necessary dependencies
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Install Python dependencies
COPY requirements.txt .
RUN pip install -r requirements.txt

# Copy application code
COPY . /app

WORKDIR /app

CMD ["python", "app.py"]
