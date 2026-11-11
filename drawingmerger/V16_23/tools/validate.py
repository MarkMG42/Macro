"""Validate supplied source/packaging. Does not execute VBA or SolidWorks.

Run from any directory: python tools/validate.py /path/to/generated/antlr/parser
Requires antlr4-python3-runtime==4.13.2 and grammars-v4 VBA7.1 generated parser.
"""
from pathlib import Path
import hashlib
import json
import re
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, sys.argv[1])
from antlr4 import InputStream, CommonTokenStream
from antlr4.error.ErrorListener import ErrorListener
from vbaLexer import vbaLexer
from vbaParser import vbaParser


class Errors(ErrorListener):
    def __init__(self):
        self.items = []

    def syntaxError(self, recognizer, offendingSymbol, line, column, msg, error):
        self.items.append(dict(line=line, column=column + 1, message=msg))


reports = []
for name, encoding in [('MergeDrawingsV16_23.bas', 'cp1251'), ('FORM_CODE.txt', 'utf-8')]:
    raw = (ROOT / name).read_bytes()
    text = raw.decode(encoding)
    # Pasted form code lacks the export-only module attribute required by this
    # grammar's startRule. Add it only to the parser input, never to the form.
    wrapper_lines = 0
    if name == 'FORM_CODE.txt':
        text = 'Attribute VB_Name = "frmDrawingOrder"\n' + text
        wrapper_lines = 1
    listener = Errors()
    lexer = vbaLexer(InputStream(text))
    lexer.removeErrorListeners()
    lexer.addErrorListener(listener)
    parser = vbaParser(CommonTokenStream(lexer))
    parser.removeErrorListeners()
    parser.addErrorListener(listener)
    parser.startRule()
    reports.append(dict(file=name, sha256=hashlib.sha256(raw).hexdigest(),
                        parser_wrapper_lines=wrapper_lines, errors=listener.items))

source = (ROOT / 'MergeDrawingsV16_23.bas').read_bytes().decode('cp1251')
assert source.replace('\r\n', '\n') == (ROOT / 'MergeDrawingsV16_23.txt').read_text()
assert '\n' not in source.replace('\r\n', '')


def procedures(text):
    return {m[1]: m[0].replace('\r\n', '\n') for m in re.finditer(
        r'^(?:Private|Public) (?:Sub|Function) (\w+)[\s\S]*?^End (?:Sub|Function)',
        text.replace('\r\n', '\n'), re.M)}


archive = ROOT.parents[0] / 'latestVerison/SOLIDWORKS_MERGE_V16_22_VBA_FORM.zip'
baseline_check = 'not available outside repository'
preserved = []
if archive.exists():
    with zipfile.ZipFile(archive) as z:
        baseline = z.read('MergeDrawingsV16_22.bas').decode('cp1251')
        assert z.read('FORM_CODE.txt') == (ROOT / 'FORM_CODE.txt').read_bytes()
    old, new = procedures(baseline), procedures(source)
    changed = {'main', 'AppendDrawing', 'MoveView', 'RebuildDrawing', 'InitRunFolder', 'ExportDrawingToPDF', 'CaptureDrawing'}
    removed = {'GetBaseName', 'ContainsSheetValue'}
    for name, body in old.items():
        if name not in changed | removed:
            assert new[name] == body, name
            preserved.append(name)
    assert set(old) - set(new) == removed
    baseline_check = 'unchanged form; all unaffected procedures byte-equivalent after newline normalization'

# Behavioral boundaries important to this change: only a drawing can reach the
# one explicit rebuild; every selected sheet is consumed once, missing names fail.
procs = procedures(source)
rebuild = procs['RebuildDrawing']
assert rebuild.index('model.GetType = swDocDRAWING') < rebuild.index('model.ForceRebuild3')
assert source.count('model.ForceRebuild3(') == 1
assert 'EditRebuild3' not in source
assert 'swDontRebuildActiveDoc' in procs['RequireActive']
assert 'RebuildDrawing' not in procs['main']
assert 'ContainsSheetValue' not in source
assert 'pending.Remove CStr(name)' in rebuild
assert 'Require pending.Count = 0' in rebuild
assert 'FINAL before save (complete)' in source
assert 'VERIFY_AFTER_SAVE' not in source
assert procs['main'].count('CaptureDrawing(') == 1
assert procs['main'].index('VerifyDrawing expectedResult, actual') < procs['main'].index('targetModel.Save3')
assert 'CaptureDrawing' not in procs['main'].split('targetModel.Save3', 1)[1]
assert 'Set actual = CaptureDrawing(targetModel)' in procs['main']
assert 'VerifyDrawing expectedResult, current' not in procs['AppendDrawing']
assert 'CaptureMappedSheets(targetModel, sheetMap)' in procs['AppendDrawing']
assert 'CaptureDrawing(targetModel)' not in procs['AppendDrawing']
assert 'CaptureDrawing(model, "", touched)' in procs['MoveView']
assert 'selectedSheets.Exists' in procs['CaptureDrawing']
assert 'MsgBox "выполнено успешно",' in source
assert 'MsgBox "ошибка, проверьте журнал",' in source
assert 'SetSheets(swExportData_ExportSpecifiedSheets, names)' in source
main = procs['main']
assert main.index('targetModel.Save3') < main.index('ExportDrawingToPDF') < main.index('PublishPDF') < main.index('WriteStatus "SUCCESS"')
assert 'WORK_PDF.pdf' in main
publish = procs['PublishPDF']
assert publish.index('CopyFile finalPath, previousPath, False') < publish.index('CopyFile stagedPath, finalPath, True')
assert 'CopyFile previousPath, finalPath, True' in publish
for dead in ['GetSaveFileNameW', 'OFN_OVERWRITEPROMPT', 'targetPath']:
    assert not re.search(r'\b' + dead + r'\b', source)

report = dict(grammar='grammars-v4/vba/vba7_1',
              grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',
              antlr_version='4.13.2', files=reports, baseline_check=baseline_check,
              preserved_procedures=preserved, static_invariants_passed=True,
              native_vba_compile=False, native_solidworks_run=False,
              pdf_content_verified=False, speedup_measured=False)
(ROOT / 'checks/validation.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n')
print(json.dumps(report, indent=2, ensure_ascii=False))
sys.exit(any(r['errors'] for r in reports))
