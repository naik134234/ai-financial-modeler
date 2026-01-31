# Deployment Configurations

## Environment Variables Needed

### Frontend (.env.local)
```bash
NEXT_PUBLIC_API_URL=https://your-backend-url.railway.app
```

### Backend (.env)
```bash
GEMINI_API_KEY=your_gemini_api_key
ALLOWED_ORIGINS=https://your-app.vercel.app,http://localhost:3000
```

## Deployment Commands

### Frontend (Vercel)
```bash
cd frontend
vercel --prod
```

### Backend (Railway)
1. Install Railway CLI: `npm install -g @railway/cli`
2. Login: `railway login`
3. Initialize: `railway init`
4. Deploy: `railway up`

## Post-Deployment Steps

1. Get backend URL from Railway dashboard
2. Update frontend environment variable `NEXT_PUBLIC_API_URL`
3. Redeploy frontend with: `vercel --prod`
