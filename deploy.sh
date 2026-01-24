#!/bin/bash

# Configuration
APP_NAME="examtool"
PORT=5000

echo "🚀 Starting deployment for $APP_NAME..."

# Check if Docker is installed
if ! command -v docker &> /dev/null; then
    echo "❌ Docker is not installed. Please install Docker first."
    exit 1
fi

# Build the Docker image
echo "📦 Building Docker image..."
docker build -t $APP_NAME .

if [ $? -ne 0 ]; then
    echo "❌ Docker build failed."
    exit 1
fi

# Stop and remove existing container if it exists
if [ "$(docker ps -aq -f name=$APP_NAME)" ]; then
    echo "🛑 Stopping existing container..."
    docker stop $APP_NAME
    docker rm $APP_NAME
fi

# Run the new container
echo "▶️ Running container on port $PORT..."
docker run -d \
  --name $APP_NAME \
  --restart unless-stopped \
  -p $PORT:5000 \
  $APP_NAME

if [ $? -eq 0 ]; then
    echo "✅ Deployment successful!"
    echo "🌍 App is running at http://localhost:$PORT"
else
    echo "❌ Failed to run container."
    exit 1
fi
