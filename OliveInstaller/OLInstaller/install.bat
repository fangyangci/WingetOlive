.\OliveMsi\python.exe .\OliveMsi\get-pip.py --no-warn-script-location
.\OliveMsi\python.exe -m pip install virtualenv --no-warn-script-location
.\OliveMsi\python.exe -m virtualenv %1
cd %1
call ".\Scripts\activate.bat
pip install olive-ai[cpu,finetune]
pip install transformers==4.44.2