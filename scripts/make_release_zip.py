"""배포용 zip 만들기 — 한글 파일 이름이 Windows에서 깨지지 않게 UTF-8 표시를 붙여 압축한다.
사용: python scripts/make_release_zip.py <출력 zip 경로>
(git에 올라간 파일만 neovue-bu-tools/ 폴더로 묶음. 리눅스 zip 명령은 한글 이름에 UTF-8 표시를 안 붙여
 Windows 탐색기에서 spec·양식 파일 이름이 깨지거나 안 보일 수 있다)"""
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
out = Path(sys.argv[1])
files = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z"], capture_output=True, check=True).stdout.decode("utf-8").split("\0")
files = [f for f in files if f]
out.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for f in files:
        z.write(ROOT / f, f"neovue-bu-tools/{f}")           # 이름에 한글이 있으면 zipfile이 UTF-8 표시(0x800)를 붙임
with zipfile.ZipFile(out) as z:
    bad = [i.filename for i in z.infolist() if any(ord(c) > 127 for c in i.filename) and not i.flag_bits & 0x800]
    assert not bad, bad
    print(f"{out.name}: {len(z.infolist())}개 파일 (한글 이름 {sum(any(ord(c) > 127 for c in i.filename) for i in z.infolist())}개 UTF-8 표시)")
