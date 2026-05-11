FROM python:3.13-slim

# Install Node.js 20
RUN apt-get update && apt-get install -y curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Node dependencies
COPY package.json ./
RUN npm install

# Copy project
COPY . .

# Create required directories
RUN mkdir -p .tmp/data .tmp/thumbnails .tmp/charts output

CMD uvicorn api.main:app --host 0.0.0.0 --port $PORT
