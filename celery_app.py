from celery import Celery
import os

# Create standalone Celery app
redis_url = os.environ.get('REDIS_URL', 'redis://localhost:6379/0')

celery_app = Celery(
    'worksheet_tasks',
    broker=redis_url,
    backend=redis_url,
    include=['tasks']  # Auto-discover tasks module
)

celery_app.conf.update(
    task_serializer='json',
    accept_content=['json'],
    result_serializer='json',
    timezone='UTC',
    enable_utc=True,
    broker_connection_retry_on_startup=True,
    task_track_started=True,
    task_send_sent_event=True,
    result_extended=True,
)

def make_celery(app):
    """Configure Celery with Flask app context"""
    class ContextTask(celery_app.Task):
        def __call__(self, *args, **kwargs):
            with app.app_context():
                return self.run(*args, **kwargs)
    
    celery_app.Task = ContextTask
    return celery_app
