@echo off
echo 打刻時間の取得をしています・・・
powershell -NoProfile -ExecutionPolicy Unrestricted .\getTime.ps1
echo 完了しました
pause > nul
exit