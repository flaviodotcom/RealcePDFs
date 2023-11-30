import PyInstaller.__main__

export_mode = '--onefile'

argumentos = [
    '--name=Destacar PDFs',
    '--windowed',
    '--noconfirm',
    '--clean',
    '--add-data=resources/images;resources/images',
    '--icon=resources/images/Taz.ico',
    export_mode, 'main.py'
]

if __name__ == '__main__':
    PyInstaller.__main__.run(argumentos)
