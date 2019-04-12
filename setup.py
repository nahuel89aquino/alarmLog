# Let's start with some default (for me) imports...
import sys
from cx_Freeze import setup, Executable


build_exe_options = {
"include_msvcr": True   #skip error msvcr100.dll missing
}
# Process the includes, excludes and packages first

carpeta = 'Conversor'  #nombre de la carpeta donde se instalara el programa

if 'bdist_msi' in sys.argv:
    sys.argv += ['--initial-target-dir', 'C:\Program File\\' + carpeta]

includes = ['wx','os','openpyxl'] #librerias que lleva tu proyecto separadas por comas entre comillas
excludes = []
packages = ['moduls']
path = []
include_files = ['img'] #carpetas que lleva tu aplicacion separadas por comas entre comillas
include_msvcr = ['networkChanger.exe.manifest']
base= None
if sys.platform == 'win32':
    base = 'Win32GUI'
if sys.platform == 'linux' or sys.platform == 'linux2':
    base = None

Conversor = Executable(
    # what to build
    script="Convert.py",
    initScript=None,
    base=base,
    icon='img\icon.ico', #ruta del icono del programa
    #shortcutName="DHCP",
    #shortcutDir="ProgramMenuFolder"
    )

setup(

    name="Conversor",
    version="1.0",
    description="Conversor de log.txt a archivos .xlsx",
    author="Aquino Nahuel",
    author_email="aquino_nahuel@hotmail.com",
    options={"build_exe": {"includes": includes,
                 "excludes": excludes,
                 "packages": packages,
                 "path": path,
                 "include_files": include_files,
                 "include_msvcr": include_msvcr,

                 }
           },

    executables=[Conversor]
    )
