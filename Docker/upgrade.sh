#!/bin/bash

# Spinner function
spinner() {
  local pid=$1
  local delay=0.15
  local spin='|/-\'
  while kill -0 $pid 2>/dev/null; do
    for i in $(seq 0 3); do
      echo -ne "\r${2} ${spin:$i:1}"
      sleep $delay
    done
  done
  echo -ne "\r${2} ✅\n"
}

echo "🚀 Getting new database from repo..."
sleep 1
git pull
echo "✅ New code has been updated successfully."
sleep 1

echo "📦 Entering root mode to apply this update"
for dots in "....." ".........." "..............."; do
  echo "📦 $dots"
  sleep 1
done

echo "🚧 Stopping current Docker container..."
(sudo docker stop leave-app && sudo docker rm leave-app) &
spinner $! "Stopping & removing old container"
sleep 1

echo "📦 Rebuilding Docker image..."
(sudo docker build -t leave-app-system .) &
spinner $! "Building Docker image"
sleep 1

echo "🚀 Starting updated container on port 8553..."
(sudo docker run -d -p 8553:80 --name leave-app leave-app-system) &
spinner $! "Starting container"
sleep 1

echo "📦 Cleaning Docker Cache..."
(sudo docker system prune -a --volumes -f) &
spinner $! "Cleaning cache"
sleep 1

echo "✅ Container running: http://localhost:8553"
sleep 1

# Closing sequence
for msg in "✅ Closing Terminal..." "✅ ...." "✅ ..." "✅ .." "✅ Byee Byeee..."; do
  echo "$msg"
  sleep 1
done
