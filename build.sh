#!/bin/bash

# Update system and install LibreOffice dependencies
apt-get update
apt-get install -y libreoffice

# Install dependencies
pip install -r requirements.txt

# Run the Flask app
python app.py
