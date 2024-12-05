from cx_Freeze import setup, Executable

script = "main.py" 

build_exe_options = {
    "packages": ["pandas", "numpy", 'cx_Frezze', "matlibplot", "tkinter"],
    "includes": ["icon.ico"]
}

setup(
    name="Heat Exchanger Sketch Generator",
    version="0.1.24",
    description="A tool to draw and visualize heat exchangers sketch",
    options={"build.exe": build_exe_options},
    executables=[Executable(script, base="Win32GUI", target_name="heat_exchanger_sketch.exe", icon="icon.ico")],
)
