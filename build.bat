copy GB_Titleblock.CATScript .\dist
copy arguments.txt .\dist
copy name_space.xlsx .\dist
python setup.py install
echo off python setup.py py2exe
python setup.py py2exe --dll-excludes=MSVCP90.DLL