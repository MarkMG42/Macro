"""Build and validate the current package using Python's standard library."""
from pathlib import Path
import hashlib
import json
import re
import zipfile
import build_release
ROOT = Path(__file__).resolve().parents[1]
def main():
    build_release.main()
    txt = (ROOT / 'MERGE_DRAWINGS_V16_10.txt').read_text(encoding='utf-8-sig')
    bas = (ROOT / 'MergeDrawingsV16_10.bas').read_text(encoding='cp1251')
    assert bas == 'Attribute VB_Name = "MergeDrawingsV16_10"\n' + txt
    assert max(map(len, txt.splitlines())) < 1023
    current = maximum = 0
    for line in txt.splitlines():
        current = current + 1 if line.rstrip().endswith(' _') else 0
        maximum = max(current, maximum)
    assert maximum <= 24
    names = re.findall(r'^(?:Public|Private) (?:Sub|Function) (\w+)', txt, re.M)
    assert len(names) == len(set(n.lower() for n in names))
    comparison = json.loads((ROOT / 'checks/comparison.json').read_text())
    for name, digest in comparison['restored_ui_functions_sha256'].items():
        proc = re.search(r'(?m)^Private Function ' + name + r'\b[\s\S]*?^End Function', txt).group()
        assert hashlib.sha256(proc.encode()).hexdigest() == digest, name
    assert not re.search(r'WScript|PowerShell|XMLHTTP|WinHTTP|URLDownload|CreateProcess|WinExec|cmd\.exe', txt, re.I)
    (ROOT / 'checks').mkdir(exist_ok=True)
    (ROOT / 'checks/static.json').write_text(json.dumps(dict(bas_txt_equivalent=True,
        procedures=len(names), max_continuations=maximum, ui_matches_v16_7=True,
        native_vba_compile=False, native_windows_ui_run=False, native_solidworks_run=False,
        antivirus_scan_performed=False), indent=2), encoding='utf-8')
    files = sorted(p for p in ROOT.rglob('*') if p.is_file() and '__pycache__' not in p.parts
                   and p.name != 'SHA256SUMS.txt')
    sums = ROOT / 'SHA256SUMS.txt'
    sums.write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.relative_to(ROOT).as_posix() + '\n'
                            for p in files), encoding='utf-8')
    files.append(sums)
    archive = ROOT.parent / 'SOLIDWORKS_MERGE_V16_10_RESTORED_SORT.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, ROOT.name + '/' + p.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        for p in files:
            assert z.read(ROOT.name + '/' + p.relative_to(ROOT).as_posix()) == p.read_bytes()
    print('Validated package:', archive.name)
    print('SHA256:', hashlib.sha256(archive.read_bytes()).hexdigest())
if __name__ == '__main__':
    main()
