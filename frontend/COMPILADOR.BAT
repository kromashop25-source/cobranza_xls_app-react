cd frontend && npm run build && cd ..
call .venv\Scripts\activate
python -m pip install -U pyinstaller pyinstaller-hooks-contrib
python -m PyInstaller backend\CobranzaTray.spec --clean --noconfirm