## Simple Python Parser using Pandas and openpyxl

[_metadata_:name]:- "Nino Casupanan"
[_metadata_:author]:- "nmcasupanan@medicardphils.com"


## Installation
- Python 3.10 or latest
- Install dependency from dependency.txt file by running below command.
```bash
pip3 install -r dependency.txt
```

## Installation as a CMD Command line (Windows)
- Open `Windows Command Line` as Administrator
- cd to `pyconverter\` folder
- run this command. 
```batch
C:Users\> install.bat
```
- You can now use the tool in any project inside your machine.

## Usage
### Single convertion
```bash
python pyconverter --json data.json --path ./ --name "CustomFileName"
```

### Multiple Convertion
- This will generate all json file in the specified folder and store it in the same directory.
```bash
python pyconverter --json-path ./json-folder
```