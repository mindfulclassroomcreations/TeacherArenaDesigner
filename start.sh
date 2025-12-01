#!/bin/bash

# Start Redis (if using local Redis)
# redis-server --daemonize yes

# Start Celery worker in background with environment variable
export CELERY_WORKER=true
celery -A tasks worker --loglevel=info &

# Unset for web app
unset CELERY_WORKER

# Start Gunicorn web server
exec gunicorn app:app --bind 0.0.0.0:8080 --timeout 600 --workers 2 --worker-class sync
