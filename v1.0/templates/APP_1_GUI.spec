# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_dynamic_libs

block_cipher = None

hiddenimports = collect_submodules("win32com") + [
	"pythoncom",
	"pywintypes",
]

binaries = collect_dynamic_libs("win32")

a = Analysis(['APP_1_GUI.py'],
			 pathex=[],
			 binaries=binaries,
			 datas=[
					 (r"mi_ruta\ico_app.ico", "."), 
					 (r"mi_ruta\PLANTILLA_CONTROL_VERSIONES.xlsb", "."),
					 (r"mi_ruta\PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb", "."),
					 (r"mi_ruta\PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb", "."), 
					],
			 hiddenimports=hiddenimports,
			 hookspath=[],
			 hooksconfig={},
			 runtime_hooks=[],
			 excludes=[],
			 win_no_prefer_redirects=False,
			 win_private_assemblies=False,
			 cipher=block_cipher,
			 noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
			 cipher=block_cipher)

exe = EXE(pyz,
		  a.scripts,
		  a.binaries,
		  a.zipfiles,
		  a.datas,  
		  [],
		  name='APP_CONTROL_VERSIONES_Y_DEPENDENCIAS_MS_ACCESS_Y_SQL_SERVER',
		  debug=False,
		  bootloader_ignore_signals=False,
		  strip=False,
		  upx=False,
		  runtime_tmpdir=None,
		  console=False,
		)

