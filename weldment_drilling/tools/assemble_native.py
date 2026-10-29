"""Assemble the native VBA fragments into the single importable standard module.

The existing core stays editable in InsertWeldWithHoles.txt; only marked regions
are replaced. Run BEFORE grammar validation after changing either fragment.
"""
from pathlib import Path
import re
ROOT = Path(__file__).resolve().parents[1]
source = ROOT / 'InsertWeldWithHoles.txt'
code = source.read_text(encoding='utf-8')
for section in ('DECLARATIONS', 'ENGINE'):
    code = re.sub(rf"\n*' BEGIN NATIVE DIMENSIONS {section}\n.*?' END NATIVE DIMENSIONS {section}\n*", '\n\n', code, flags=re.S)
declarations = (ROOT / 'source/native_dimensions_declarations.vba').read_text(encoding='utf-8').rstrip()
engine = (ROOT / 'source/native_dimensions.vba').read_text(encoding='utf-8').rstrip()
code = code.replace('Public Sub start()', "' BEGIN NATIVE DIMENSIONS DECLARATIONS\n" + declarations + "\n' END NATIVE DIMENSIONS DECLARATIONS\n\nPublic Sub start()", 1)
code = code.rstrip() + "\n\n' BEGIN NATIVE DIMENSIONS ENGINE\n" + engine + "\n' END NATIVE DIMENSIONS ENGINE\n"
source.write_text(code, encoding='utf-8')
print('Assembled single-module native VBA release')
