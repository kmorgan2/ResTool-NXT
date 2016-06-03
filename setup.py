from distutils.core import setup
import sys
# import py2exe

sys.argv.append('py2exe')

setup(
    options={'py2exe': {'bundle_files': 3, 'compressed': True}},
    windows=[{"script": "main.py", "icon_resources": [(1, "icon.ico")], 'uac_info': "requireAdministrator"}],
    zipfile=None, requires=['twisted', 'requests', 'pywinauto'])
