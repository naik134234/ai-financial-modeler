@echo off
echo ===================================================
echo   AI Financial Modeler - Deployment Helper
echo ===================================================
echo.
echo This script will help you link your local project to GitHub.
echo.
echo 1. Go to https://github.com/new and create a repository named 'ai-financial-modeler'
echo 2. Do NOT initialize it with README, .gitignore or License (keep it empty)
echo 3. Copy the HTTPS URL (e.g., https://github.com/username/ai-financial-modeler.git)
echo.
set /p REPO_URL="Enter your GitHub Repository URL here: "

if "%REPO_URL%"=="" goto error

echo.
echo Linking remote origin...
git remote remove origin 2>nul
git remote add origin %REPO_URL%

echo.
echo Pushing code to main...
git branch -M main
git push -u origin main

echo.
echo ===================================================
echo   GitHub Upload Complete!
echo ===================================================
echo.
echo NOW DEPLOY TO VERCEL:
echo 1. Run 'npx vercel' in this terminal (if you have Vercel CLI)
echo    OR
echo 2. Go to https://vercel.com/new and import this repository.
echo.
pause
exit /b

:error
echo Error: Repository URL cannot be empty.
pause
