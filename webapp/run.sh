#!/bin/bash
# Operating Model Observatory — Start script
cd "$(dirname "$0")"

echo "Starting Digital Products Observatory..."
cd frontend
npx vite --port 3000 --host --open
