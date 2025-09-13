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

sudo docker exec -it leave-app touch /var/www/html/down/down.flag

echo "✅ Access Granted..."
sleep 1
echo "🔧 Entering maintenance mode..."
sleep 1
echo "🔧 Re-routed network begin in..."
sleep 1
for i in 3 2 1; do
  echo "🔧 $i"
  sleep 1
done

# Enable maintenance mode with spinner
(echo && sleep 3) &
spinner $! "🔧 Enabling maintenance mode"

echo "✅ Maintenance mode enabled. All requests will now show maintenance page."
sleep 1

# Closing sequence
for msg in "✅ Closing Terminal..." "✅ ...." "✅ ..." "✅ .." "✅ Byee Byeee..."; do
  echo "$msg"
  sleep 1
done
