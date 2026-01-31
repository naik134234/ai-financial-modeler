# 🚀 Deployment Guide: GitHub & Vercel

This project is configured as a monorepo, meaning you can deploy **both** the Frontend (Next.js) and Backend (FastAPI) to Vercel in a single project!

## 1. Prepare for GitHub

First, ensure all your code is committed:

```bash
git add .
git commit -m "Ready for deployment with OpenAI and Vercel config"
```

## 2. Push to GitHub

1.  Go to [GitHub.com/new](https://github.com/new) and create a new repository (e.g., `ai-financial-modeler`).
2.  Run the commands shown by GitHub to push your existing code:

```bash
git remote add origin https://github.com/YOUR_USERNAME/ai-financial-modeler.git
git branch -M main
git push -u origin main
```

## 3. Deploy to Vercel

1.  Go to [Vercel Dashboard](https://vercel.com/dashboard).
2.  Click **"Add New..."** -> **"Project"**.
3.  Import your `ai-financial-modeler` repository.
4.  **Framework Preset**: Select `Next.js` (or leave as Other, Vercel usually auto-detects `vercel.json`).
5.  **Root Directory**: Leave as `./` (Project Root).

### 🔑 Environment Variables
Expand the **"Environment Variables"** section and add:

| Key | Value |
|-----|-------|
| `GEMINI_API_KEY` | *(Your Gemini API Key)* |
| `OPENAI_API_KEY` | *(Your OpenAI API Key)* |
| `NEXT_PUBLIC_API_URL` | `/api` *(Important: Use relative path /api so frontend calls the serverless backend)* |

6.  Click **"Deploy"**.

## 4. Verification

Once deployed:
- **Frontend**: `https://your-project.vercel.app`
- **Backend API**: `https://your-project.vercel.app/api/docs` (Swagger UI)

Everything is pre-configured in `vercel.json` to route `/api/*` requests to your Python backend!
