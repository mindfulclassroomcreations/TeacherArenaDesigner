#!/bin/bash

# Start Redis (if using local Redis)
# redis-server --daemonize yes

# Start Celery worker in background
celery -A tasks worker --loglevel=info &

# Start Gunicorn web server
exec gunicorn app:app --bind 0.0.0.0:8080 --timeout 600 --workers 2 --worker-class sync
