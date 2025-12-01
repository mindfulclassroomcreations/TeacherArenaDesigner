# Render.com Deployment Guide

## Quick Setup Steps

### 1. Deploy to Render

1. Go to [Render.com](https://render.com) and sign up/login
2. Click **"New +"** → **"Web Service"**
3. Connect your GitHub account
4. Select the **`TeacherArenaDesigner`** repository
5. Render will auto-detect the `render.yaml` configuration
6. Click **"Create Web Service"**

### 2. Set Environment Variables

After deployment starts, go to your service's **Environment** tab and add:

- **`SECRET_KEY`** = (Auto-generated or set your own random string)
- **`DATABASE_URL`** = (Leave empty for SQLite, or add PostgreSQL URL for production)

### 3. Set OpenAI API Key

**IMPORTANT:** You must set the OpenAI API key through the admin panel:

1. Wait for deployment to complete (check the Logs tab)
2. Go to your app URL: `https://your-app-name.onrender.com`
3. Navigate to: `https://your-app-name.onrender.com/login`
4. **Default credentials:**
   - Username: `admin`
   - Password: `admin123`
5. Once logged in, you'll see the Admin Dashboard
6. **Paste your OpenAI API Key** in the input field
7. Click **"Update API Key"**

### 4. Test the Application

1. Go back to the main page
2. Click on **"Academy Ready"** or **"Dreaming Caterpillar"**
3. Upload your `details.xlsx` file
4. Click **"Generate Worksheets"**

## Troubleshooting

### Issue: "OpenAI API Key not set"
**Solution:** Login to `/login` and set the API key in the admin panel

### Issue: Generation starts but doesn't complete
**Possible causes:**
- Invalid OpenAI API key
- OpenAI API rate limits
- Render free tier timeout (services spin down after 15 min inactivity)

**Check the logs:**
- Go to Render Dashboard → Your Service → **Logs** tab
- Look for error messages

### Issue: 500 Internal Server Error
**Solution:** 
- Check Render logs for Python errors
- Ensure all dependencies in `requirements.txt` are installed
- Verify database is initialized properly

### Issue: App is slow or times out
**Free tier limitations:**
- Service spins down after 15 minutes of inactivity
- First request after spin-down takes 30-60 seconds to wake up
- Consider upgrading to a paid plan for production use

## Important Notes

1. **Change default password:** After first login, create a new admin user with a secure password
2. **OpenAI API costs:** Each worksheet generation uses the OpenAI API and will incur costs on your OpenAI account
3. **Free tier limits:** Render free tier has 750 hours/month and services spin down when inactive
4. **Database:** SQLite is used by default. For production, add a PostgreSQL database from Render

## Getting Your App URL

After deployment completes, your app will be available at:
```
https://your-app-name.onrender.com
```

You can find this URL in the Render dashboard at the top of your service page.

## Support

If you encounter issues:
1. Check the Render Logs tab for error messages
2. Verify OpenAI API key is set correctly
3. Ensure your Excel file follows the correct format
4. Check that your OpenAI account has available credits
