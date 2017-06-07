### Build for Unix like:

```bash
chmod +x build.sh
./build.sh
```


### Build for Windows:

```bash
pip install -r requirement.txt
python iconGenerator.py
pyinstaller -w -F --icon=logo.ico main.py
```