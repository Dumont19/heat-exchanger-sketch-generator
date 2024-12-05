from cx_Freeze import setup, Executable

script = "main.py" 

setup(
    name="Heat Exchanger ",
    version="0.1.24",
    description="A tool to draw and visualize heat exchangers sketch",
    executables=[Executable(script, base=None, target_name="heat_exchanger_sketch.exe")],
)
