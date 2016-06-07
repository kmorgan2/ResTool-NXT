from distutils.core import setup
import sys
import py2exe

print py2exe.__version__  # This is exclusively so Pycharm won't complain about an unused libarary

sys.argv.append('py2exe')

setup(
    options={'py2exe': {'bundle_files': 3, 'compressed': True}},
    windows=[{"script": "main.py", "icon_resources": [(1, "icon.ico")], 'uac_info': "requireAdministrator"}],
    zipfile=None, requires=['twisted', 'requests', 'pywinauto', 'pywin32'])
