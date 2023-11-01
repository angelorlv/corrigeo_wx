import cx_Freeze

executables = [cx_Freeze.Executable(
    script="app.py",
    base="MacOSGUI",
    #icon="votre_icone.ico",  # Spécifiez une icône personnalisée (facultatif)
)]

cx_Freeze.setup(
    name="Corrigeo.wx (beta 0.)",
    version="0.1",
    description="Application de Correction de Thèse Universitaire (beta)",
    executables=executables
)