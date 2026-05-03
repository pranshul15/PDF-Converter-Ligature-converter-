# Use a lightweight Python base image
FROM python:3.10-slim

# Install Ghostscript and clean up to keep the image small
RUN apt-get update --fix-missing && \
    apt-get install -y --no-install-recommends \
    ghostscript \
    libgl1 \
    libglib2.0-0 && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app

# Install Python dependencies
RUN pip install --no-cache-dir pdf2docx

# Copy your script into the container
COPY main.py .

# Command to run the script
CMD ["python", "main.py"]