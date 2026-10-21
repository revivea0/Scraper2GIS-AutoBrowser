"""Build a native release on the target OS: python scripts/build_release.py."""
import hashlib
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
NAME = 'Scraper2GIS'


def main():
    os.chdir(ROOT)
    target = {'win32': 'Windows', 'darwin': 'macOS', 'linux': 'Linux'}[sys.platform]
    arch = 'arm64' if platform.machine().lower() in ('arm64', 'aarch64') else 'x64'
    command = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir',
               '--windowed', '--name', NAME, '--add-data', 'gui:gui',
               '--collect-all', 'playwright', '--collect-all', 'webview',
               '--recursive-copy-metadata', 'playwright', '--recursive-copy-metadata', 'pywebview',
               '--recursive-copy-metadata', 'openpyxl']
    if sys.platform == 'linux':
        command += ['--hidden-import', 'webview.platforms.qt']
    subprocess.run(command + ['main.py'], check=True)
    if sys.platform == 'darwin':
        app = ROOT / 'dist' / (NAME + '.app')
        executable = app / 'Contents' / 'MacOS' / NAME
    else:
        app = ROOT / 'dist' / NAME
        executable = app / (NAME + ('.exe' if sys.platform == 'win32' else ''))
    subprocess.run([str(executable), '--smoke-test'], check=True, timeout=60)
    release = ROOT / 'dist' / f'{NAME}-{target}-{arch}'
    release.mkdir(exist_ok=True)
    shutil.copytree(app, release / app.name, dirs_exist_ok=True)
    shutil.copy2(ROOT / 'LICENSE', release)
    shutil.copy2(ROOT / 'docs' / 'START.txt', release)
    if sys.platform == 'darwin':
        archive = release.with_suffix('.zip')
        subprocess.run(['ditto', '-c', '-k', '--sequesterRsrc', '--keepParent', str(release), str(archive)], check=True)
    else:
        archive = Path(shutil.make_archive(str(release), 'zip' if sys.platform == 'win32' else 'gztar',
                                           root_dir=release.parent, base_dir=release.name))
    with archive.open('rb') as stream:
        digest = hashlib.file_digest(stream, 'sha256').hexdigest()
    archive.with_name(archive.name + '.sha256').write_text(f'{digest}  {archive.name}\n', encoding='utf-8')
    print(f'Ready: {archive}')


if __name__ == '__main__':
    main()
