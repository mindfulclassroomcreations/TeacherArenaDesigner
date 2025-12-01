# Deployment Guide

## PDF Generation
✅ **PDFs are created automatically** - No additional tools needed! The app uses ReportLab to generate PDFs from uploaded Excel files.

## Deployment Options

### Option 1: Render.com (Recommended - Free Tier)
1. Push code to GitHub
2. Go to [render.com](https://render.com)
3. Create new "Web Service"
4. Connect your GitHub repo
5. Set up:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
6. Add Environment Variables:
   - `DATABASE_URL` (if using PostgreSQL, or leave blank for SQLite)
   - `SECRET_KEY` (generate random string)

### Option 2: Railway.app
1. Push to GitHub
2. Go to [railway.app](https://railway.app)
3. "New Project" → "Deploy from GitHub"
4. Railway auto-detects Flask
5. Add PostgreSQL database (optional)
6. Set environment variables

### Option 3: Using Supabase Database
1. Create project at [supabase.com](https://supabase.com)
2. Get PostgreSQL connection string from Settings → Database
3. Set as `DATABASE_URL` environment variable:
   ```
   postgresql://user:password@host:5432/database
   ```

### Option 4: Heroku
1. Install Heroku CLI
2. Login: `heroku login`
3. Create app: `heroku create your-app-name`
4. Add PostgreSQL: `heroku addons:create heroku-postgresql:mini`
5. Deploy: `git push heroku main`

## GitHub Setup
```bash
cd "/Users/sankalpa/Desktop/untitled folder/01_03 WORKSHEETS + 30 TASK CARDS"
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

## Environment Variables Required
- `DATABASE_URL` - PostgreSQL connection string (optional, uses SQLite if not set)
- `SECRET_KEY` - Random secret key for Flask sessions
- OpenAI API Key set via admin panel after deployment

## Important Notes
- ⚠️ **Vercel is NOT suitable** for this app (long-running processes, file generation)
- ✅ PDF generation works automatically on all platforms
- SQLite works for small deployments, PostgreSQL recommended for production
- Admin credentials: username: `admin`, password: `admin123` (change after first login)
