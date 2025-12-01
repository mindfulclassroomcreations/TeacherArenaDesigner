# Background Task Processing Setup

This application now uses Celery with Redis for background task processing, allowing users to browse while worksheets generate.

## Architecture

- **Flask App**: Handles web requests and serves UI
- **Celery Worker**: Processes worksheet generation in background
- **Redis**: Message broker for task queue

## Local Development Setup

1. **Install Redis** (if not already installed):
   ```bash
   # macOS
   brew install redis
   
   # Ubuntu/Debian
   sudo apt-get install redis-server
   ```

2. **Start Redis**:
   ```bash
   redis-server
   ```

3. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Start Celery worker** (in a separate terminal):
   ```bash
   celery -A tasks worker --loglevel=info
   ```

5. **Start Flask app**:
   ```bash
   python app.py
   ```

## Fly.io Deployment

### Option 1: Using Upstash Redis (Recommended)

1. **Create Upstash Redis instance**:
   ```bash
   flyctl redis create
   ```
   Follow prompts to create a Redis instance. This will automatically set the `REDIS_URL` secret.

2. **Deploy**:
   ```bash
   flyctl deploy
   ```

### Option 2: Set Redis URL Manually

If using an external Redis service (like Redis Cloud, AWS ElastiCache, etc.):

```bash
flyctl secrets set REDIS_URL="redis://your-redis-url:6379/0"
flyctl deploy
```

## How It Works

### Frontend (JavaScript)
1. User submits Excel file
2. JavaScript makes POST request to `/generate-academy-async` or `/generate-caterpillar-async`
3. Receives `task_id` immediately (non-blocking)
4. Polls `/task-status/{task_id}` every 2 seconds
5. Updates progress bar and shows completed worksheets as they finish
6. User can navigate away and come back - generation continues

### Backend (Python)
1. Flask endpoint saves file and creates Celery task
2. Returns task ID immediately (202 status)
3. Celery worker picks up task from Redis queue
4. Worker generates worksheets, updating progress in Redis
5. Individual lesson downloads become available as soon as each completes
6. Full ZIP available when all lessons finish

## Benefits

✅ **Non-blocking**: Users can browse other pages while generation happens
✅ **Real-time progress**: See which lessons are complete and download them immediately
✅ **Concurrent requests**: Multiple users can generate worksheets simultaneously
✅ **Resilient**: If connection drops, task continues; user can check status later
✅ **Scalable**: Add more Celery workers to handle more concurrent generations

## Monitoring

Access Celery Flower dashboard (development only):
```bash
celery -A tasks flower
```
Then visit: http://localhost:5555

## API Endpoints

### Async Generation
- `POST /generate-academy-async` - Start academy worksheet generation
- `POST /generate-caterpillar-async` - Start caterpillar worksheet generation

Returns:
```json
{
  "task_id": "abc-123-def",
  "status": "started",
  "status_url": "/task-status/abc-123-def"
}
```

### Check Progress
- `GET /task-status/{task_id}` - Check task progress

Returns:
```json
{
  "state": "PROGRESS",
  "status": "Completed: Lesson 3",
  "current": 3,
  "total": 10,
  "individual_files": [
    {
      "topic": "Lesson 1",
      "filename": "lesson_1.zip",
      "download_url": "/download/lesson_1.zip"
    }
  ]
}
```

### Download Files
- `GET /download/{filename}` - Download generated worksheet ZIP

## Environment Variables

- `REDIS_URL` - Redis connection URL (default: `redis://localhost:6379/0`)
- `OPENAI_API_KEY` - OpenAI API key for content generation
- `SECRET_KEY` - Flask secret key for sessions
- `DATABASE_URL` - PostgreSQL database URL

## Troubleshooting

### Celery worker not starting
- Check Redis is running: `redis-cli ping` (should return PONG)
- Check Redis URL is correct in logs
- Ensure all dependencies installed: `pip install -r requirements.txt`

### Tasks stuck in PENDING
- Celery worker may not be running
- Check worker logs for errors
- Verify Redis connection: `redis-cli -u $REDIS_URL ping`

### Memory issues
- Increase VM memory in fly.toml if needed
- Consider limiting concurrent tasks in Celery config
- Worksheets with many lessons may need more RAM
