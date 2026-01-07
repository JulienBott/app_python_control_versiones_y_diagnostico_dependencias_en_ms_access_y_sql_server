# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_dynamic_libs

block_cipher = None

hiddenimports = collect_submodules("win32com") + [
	"pythoncom",
	"pywintypes",
]

binaries = collect_dynamic_libs("win32")

a = Analysis(
	['APP_1_GUI.py'],
	pathex=[],
	binaries=binaries,
	datas=[
		(r"mi_ruta\ico_app.ico", "."),
		(r"mi_ruta\img_guia_usuario.png", "."),
		(r"mi_ruta\GUIA_USUARIO_V1.1.pdf", "."),
		(r"mi_ruta\img_boton_procesos.png", "."),
		(r"mi_ruta\img_boton_add.png", "."),
		(r"mi_ruta\img_boton_clear.png", "."),
		(r"mi_ruta\img_boton_sql_server_authentication.png", "."),
		(r"mi_ruta\img_seleccionar_all_none.png", "."),
		(r"mi_ruta\img_boton_dependencias_sql_server.png", "."),
		(r"mi_ruta\img_control_versiones_boton_ver.png", "."),
		(r"mi_ruta\img_control_versiones_boton_excel.png", "."),
		(r"mi_ruta\img_control_versiones_boton_migrar_lineas_codigo.png", "."),
		(r"mi_ruta\img_control_versiones_boton_merge_bbdd_fisica.png", "."),
		(r"mi_ruta\PLANTILLA_CONTROL_VERSIONES.xlsb", "."),
		(r"mi_ruta\PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb", "."),
		(r"mi_ruta\PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb", "."),
			
	],
	hiddenimports=hiddenimports,
	hookspath=[],
	runtime_hooks=[],
	excludes=[],
	win_no_prefer_redirects=False,
	win_private_assemblies=False,
	cipher=block_cipher,
	noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
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