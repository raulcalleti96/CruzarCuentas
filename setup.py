from setuptools import setup

APP = ['Principal.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'icon.icns',  # Elimina esta línea si no tienes icono
    'packages': ['pandas'],
}

setup(
    app=APP,
    name='CruzarCuentas',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)