from cx_Freeze import setup, Executable

exe = [Executable("fusion.py", base = "Win32GUI")]
setup(
    name = "salut",
    version = "0.1",
    description = "Ce programme vous dit bonjour",
    executables = exe,
)