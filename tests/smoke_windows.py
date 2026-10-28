"""Exercise the packaged Windows update replacement on disposable files."""
import hashlib
from pathlib import Path
import subprocess
import tempfile
import time


exe = Path('dist/ExamAccessLogGenerator.exe').resolve()
with tempfile.TemporaryDirectory() as folder:
    target = Path(folder) / 'StudentSortUpdateSmoke.exe'
    target.write_bytes(b'MZ synthetic old executable')
    try:
        result = subprocess.run(
            [str(exe), '--apply-update', str(target), '4294967294'], timeout=90
        )
        assert result.returncode == 0, result.returncode
        assert hashlib.sha256(target.read_bytes()).digest() == hashlib.sha256(exe.read_bytes()).digest()
        time.sleep(5)
        running = subprocess.check_output(['tasklist', '/FO', 'CSV', '/FI', 'IMAGENAME eq StudentSortUpdateSmoke.exe'])
        assert b'StudentSortUpdateSmoke.exe' in running, running
    finally:
        subprocess.run(['taskkill', '/F', '/IM', 'StudentSortUpdateSmoke.exe'], check=False)
        time.sleep(2)
print('Packaged updater replaced the old EXE and started the updated GUI.')
