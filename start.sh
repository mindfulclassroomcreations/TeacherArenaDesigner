#!/bin/bash

# Start Redis (if using local Redis)
# redis-server --daemonize yes

# Start Celery worker in background with environment variable
# Increased concurrency to 4 to allow multiple simultaneous generations
export CELERY_WORKER=true
celery -A celery_app.celery_app worker --loglevel=info --concurrency=4 &

# Unset for web app
unset CELERY_WORKER

# Start Gunicorn web server
exec gunicorn app:app --bind 0.0.0.0:8080 --timeout 600 --workers 2 --worker-class sync
