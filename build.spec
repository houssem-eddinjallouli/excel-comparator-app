# Add this to your existing code
block_cipher = None

a = Analysis(['main.py'],
             pathex=['D:\\workspace\\excel-comparator-app'],
             binaries=[],
             datas=[],
             hiddenimports=['pandas', 'openpyxl', 'xlrd'],
             hookspath=[],
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
          [],
          exclude_binaries=True,
          name='ExcelComparator',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,  # Set to True if you want to see console output
          icon='icon.ico')  # Optional icon
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='ExcelComparator')