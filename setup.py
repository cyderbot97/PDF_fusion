# commande à taper en ligne de commande après la sauvegarde de ce fichier:
# python setup.py build
from cx_Freeze import setup, Executable
  
executables = [
        Executable(script = "fusion.py", base = "Win32GUI" )
]
# ne pas mettre "base = ..." si le programme n'est pas en mode graphique, comme c'est le cas pour chiffrement.py.
  
buildOptions = dict( 
        includes = ["tkinter","PyPDF2","os","configparser","docx2pdf","pathlib"],
        include_files = ["config.ini"]
)
  
setup(
    name = "Fusion word pdf",
    version = "1.0",
    description = "description du programme",
    author = "cyril",
    options = dict(build_exe = buildOptions),
    executables = executables
)