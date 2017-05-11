#catia

1.安装py2exe，注意版本，这里使用0.6.9，如果使用0.9.2，则会出现错误ImportError:No module named machinery，因为0.9.2版本只适用于python3。
https://sourceforge.net/projects/py2exe/files/py2exe/0.6.9/

2.安装qrcode，使用命令easy_install --always-unzip qrcode
py2exe可以将Python的程序转换城生成window平台使用的可执行文件，从而可以脱离python环境单独运行。
但有时候用py2exe生成的文件会报can’t find module name等错误，原因很可能是这个模块是用egg安装的Egg类似Java的jar文件，是一种打包好的python库文件。
用easy_install安装这种格式的库很方便，但是当前版本的py2exe还不能找到egg中的模块，解决办法最简单的就是用不要用单独的egg库，而是将其解压安装：easy_install --always-unzip选项，如: easy_install --always-unzip qrcode。
如果直接用python的话，输入python setup.py install_lib（如果有install_data和install_scripts，也加上)代替install。