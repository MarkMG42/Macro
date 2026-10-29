"""Assemble the standard module; companion frmDrawingOrder is required."""
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
def main():
    module = '\n\n'.join((ROOT / 'tools' / n).read_text(encoding='utf-8').strip()
                         for n in ('header.vba', 'merge_core.vba', 'source_properties.vba', 'linked_notes.vba', 'order_dialog.vba', 'helpers.vba')) + '\n'
    (ROOT / 'MERGE_DRAWINGS_V16_19.txt').write_bytes(module.replace('\n', '\r\n').encode('utf-8-sig'))
    (ROOT / 'MergeDrawingsV16_19.bas').write_bytes(('Attribute VB_Name = "MergeDrawingsV16_19"\n' + module).replace('\n', '\r\n').encode('cp1251'))
    print(f'Built {len(module.splitlines())} VBA lines')
if __name__ == '__main__':
    main()
