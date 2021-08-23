:: Build xls2gd with PyInstaller (Windows)
:: @copyright Hidden Moss
:: @see Hidden Moss: https://hiddenmoss.com/
:: @see git repo: https://github.com/hidden-moss/xls2gd
:: @author: Yuancheng Zhang, https://github.com/endaye
powershell -command "python .\gui.py .\tool_xls2gd.py;pyinstaller .\tool_xls2gd.spec;cp .\dist\xls2gd.exe .\xls2gd.exe"