#!/bin/bash

echo "✅ Disabling maintenance mode..."
sudo docker exec -it leave-app rm -f /var/www/html/down/down.flag
echo "🟢 Maintenance mode disabled."
echo "Your website is live again!"