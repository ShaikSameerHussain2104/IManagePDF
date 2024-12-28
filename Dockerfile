# Use the official Python image from Docker Hub
FROM python:3.9-slim

# Set environment variables
ENV LIBREOFFICE_VERSION=7.4.3.2

# Install required system dependencies including LibreOffice and other packages
RUN apt-get update && \
    apt-get install -y \
    libreoffice \
    poppler-utils \
    gcc \
    g++ \
    python3-dev \
    libxml2-dev \
    libxslt1-dev \
    libjpeg-dev \
    libpq-dev \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Set working directory to /app
WORKDIR /app

# Copy the requirements.txt file into the container
COPY ./requirements.txt /app/requirements.txt

# Install Python dependencies
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copy the application code to the container
COPY . /app

# Expose the port Flask will run on
EXPOSE 5000

# Set environment variables for Flask
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0

# Command to run the Flask application
CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]
