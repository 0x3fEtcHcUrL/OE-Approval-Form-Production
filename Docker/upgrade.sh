#!/bin/bash

echo "🚀 Getting new database from repo..."
git pull
echo "✅ New code has been updated successfully."

echo "📦 Entering root mode to apply this update"
echo "📦 ....."
echo "📦 ........."

sudo docker stop leave-app && sudo docker rm leave-app
echo "✅ Access Granted..."
echo "🚧 Stopping current Docker container..."
echo "✅ Stopped and removed old container."

echo "📦 Rebuilding Docker image..."
sudo docker build -t leave-app-system .
echo "✅ Image rebuilt."

echo "🚀 Starting updated container on port 8553..."
sudo docker run -d -p 8553:80 --name leave-app leave-app-system
echo "✅ Container running: http://localhost:8553"

echo "📦 Deleting Docker Cache"
sudo docker system prune -a --volumes
echo "✅ Docker Cache Successfully Deleted!"