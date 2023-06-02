### A collection of my Python scripts that I think might be useful to someone.

## create virtual environment, met requirements and run (UNIX)
```bash
# You can clone whole repository
git clone --depth=1 https://github.com/mateuszferenc/my_pythons

# Just one chosen directory (project)
git clone -n --depth=1 --filter=tree:0 https://github.com/mateuszferenc/my_pythons
cd my_pythons
git sparse-checkout set --no-cone <selected-directory>
git checkout

# Then you can create a virtual environment in the selected directory.
python3 -m venv venv
source venv/bin/activate
# Install required modules
pip3 install -r requirements.txt
# Run main script, sometimes can be in child directories
python3 main.py 
```