### A collection of my Python scripts that I think might be useful to someone.

## create virtual environment, met requirements and run (UNIX)
```bash
# you can clone whole repository or just one chosen
git clone https://github.com/mateuszferenc/my_pythons/<chosen-dir>
# change directory to chosen one
cd <chosen-dir>
python3 -m venv venv
source venv/bin/activate
# install required modules
pip3 install -r requirements.txt
# run main script, sometimes can be in child directories
python3 main.py 
```