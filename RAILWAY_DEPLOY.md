# Railway Backend Deployment Guide

## Quick Deploy via Railway Dashboard (Recommended)

Your Railway project is ready! Deploy your backend in 3 minutes:

### Step 1: Open Railway Project
Visit: https://railway.com/project/8bc4b337-3b88-4fa3-bc3c-8bd78a7eabdc

### Step 2: Add GitHub Service
1. Click **"New"** button
2. Select **"Deploy from GitHub repo"**
3. Choose **"naik134234/ai-financial-modeler"**
4. Railway will auto-detect Python and start building from `backend/` directory

### Step 3: Set Root Directory
1. After service is created, go to **Settings**
2. Find **"Root Directory"**
3. Set to: `backend`
4. Save changes

### Step 4: Add Environment Variable
1. Go to **Variables** tab
2. Click **"New Variable"**
3. Add:
   - Key: `GEMINI_API_KEY`
   - Value: `your_gemini_api_key_here`
4. Click **"Add"**

### Step 5: Get Your Backend URL
1. Go to **Settings** → **Networking**
2. Click **"Generate Domain"**
3. Copy the URL (e.g., `https://your-app.up.railway.app`)

### Step 6: Update Frontend
1. Go to Vercel: https://vercel.com/naik134234s-projects/ai-financial-modeler
2. Settings → Environment Variables
3. Add:
   - Key: `NEXT_PUBLIC_API_URL`
   - Value: Your Railway backend URL
4. Redeploy: Deployments → Click "..." → Redeploy

## Done! 🎉

Your app will be fully functional with:
- Frontend: https://ai-financial-modeler.vercel.app
- Backend: Your Railway URL
- Code: https://github.com/naik134234/ai-financial-modeler
