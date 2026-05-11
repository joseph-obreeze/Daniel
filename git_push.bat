"C:\Program Files\Git\cmd\git.exe" init
"C:\Program Files\Git\cmd\git.exe" config user.name "joseph-obreeze"
"C:\Program Files\Git\cmd\git.exe" config user.email "joseph-obreeze@users.noreply.github.com"
"C:\Program Files\Git\cmd\git.exe" add .
"C:\Program Files\Git\cmd\git.exe" commit -m "chore: Update Sales Analyzer App v1.1.3"
"C:\Program Files\Git\cmd\git.exe" branch -M main
"C:\Program Files\Git\cmd\git.exe" remote remove origin
"C:\Program Files\Git\cmd\git.exe" remote add origin "https://github.com/joseph-obreeze/Daniel.git"
"C:\Program Files\Git\cmd\git.exe" push -u origin main
