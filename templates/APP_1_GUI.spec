# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['APP_1_GUI.py'],
             pathex=[],
             binaries=[],
             datas=[
			 
			 (r"mi_ruta\ico_app.ico", "."), 
			 (r"mi_ruta\PLANTILLA_CONTROL_VERSIONES.xlsb", "."),
			 (r"mi_ruta\PLANTILLA_DIAGNOSTICO_MS_ACCESS.xlsb", "."),
			 (r"mi_ruta\PLANTILLA_DIAGNOSTICO_SQL_SERVER.xlsb", ".")
			 
					],
             hiddenimports=[],
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
          name='APP_CONTROL_VERSIONES_MS_ACCESS_SQL_SERVER',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
