@IF /I "%~1"=="-a" GOTO AHK
pyinstaller --noconfirm --onedir --console --icon "icon/tablet.ico" --upx-dir "../pycompile/upx-4.0.2-win64" --paths "Lib/site-packages" --add-data "D2Coding-01.ttf;." --dist "deploy" --name "sccore" "SnScanner.py"
:AHK
"C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /in SnScanner.ahk /out deploy\SnScanner.exe /icon icon\tablet.ico 
