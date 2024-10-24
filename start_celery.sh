#!/bin/bash

# Detect the port based on an environment variable or argument
if [ -z "$1" ]; then
    echo "Usage: $0 <port>"
    exit 1
fi

PORT=$1

# Set worker name based on the port number
if [ "$PORT" == "5000" ]; then
    WORKER_NAME="worker1"
elif [ "$PORT" == "5001" ]; then
    WORKER_NAME="worker2"
else
    WORKER_NAME="worker_$PORT"
fi

# Start Celery worker with unique hostname
celery -A app.celery worker --loglevel=info --hostname=${WORKER_NAME}@%h
