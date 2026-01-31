@echo off
echo ===================================================
echo   GitHub Permission Fix
echo ===================================================
echo.
echo Your current login is missing 'Repository' permissions. 
echo We need to refresh them to allow uploading code.
echo.
echo 1. I will launch the auth refresh command.
echo 2. A code will appear (e.g. ABCD-1234).
echo 3. A browser window will open.
echo 4. Paste the code and Approve.
echo.
pause

gh auth refresh -h github.com -s delete_repo,repo,workflow

echo.
echo Permissions refreshed! Now pushing code...
git push -u origin-new temp_deploy:main
echo.
echo DONE!
pause
