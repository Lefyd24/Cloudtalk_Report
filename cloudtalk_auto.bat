@echo off
echo Running script at %date% %time% >> .log\batch_log.txt
call C:\Users\lefteris.fthenos\PY_Apps\Reports\Cloudtalk_Report\.venv\Scripts\activate
echo Virtual environment activated >> .log\batch_log.txt
C:\Users\lefteris.fthenos\PY_Apps\Reports\Cloudtalk_Report\.venv\Scripts\python.exe C:\Users\lefteris.fthenos\PY_Apps\Reports\Cloudtalk_Report\report_ppt.py >> .log\logfile.txt 2>&1
echo Script execution completed >> .log\batch_log.txt