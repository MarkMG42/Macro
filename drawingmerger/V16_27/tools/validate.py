from pathlib import Path
import re,sys,json,hashlib
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,sys.argv[1])
from antlr4 import InputStream,CommonTokenStream
from antlr4.error.ErrorListener import ErrorListener
from vbaLexer import vbaLexer
from vbaParser import vbaParser
class Errors(ErrorListener):
 def __init__(self):self.items=[]
 def syntaxError(self,r,s,line,column,msg,error):self.items.append(dict(line=line,column=column+1,message=msg))
def procs(s):return {m[1]:m[0] for m in re.finditer(r'^(?:Private|Public) (?:Sub|Function) (\w+)[\s\S]*?^End (?:Sub|Function)',s,re.M)}
s=(ROOT/'MergeDrawingsV16_27.txt').read_text();raw=(ROOT/'MergeDrawingsV16_27.bas').read_bytes();old=(ROOT.parent/'V16_25/MergeDrawingsV16_25.txt').read_text();a,b=procs(old),procs(s)
assert raw.decode('cp1251').replace('\r\n','\n')==s
assert '\n' not in raw.decode('cp1251').replace('\r\n','')
changed={'main','RequireActive','CloseOwnedDrawing','ReleaseSource','CaptureDrawing','CaptureStructure'}
removed={'SuspendTargetGraphics','RestoreTargetGraphics'};preserved=[]
for name,body in a.items():
 if name not in changed|removed:
  assert b[name].replace('V16_27','V16_25').replace('V16.27','V16.25')==body,name;preserved.append(name)
assert set(a)-set(b)==removed
for name in ['CaptureDrawing','CaptureStructure']:
 assert re.sub(r'^\s*SuspendDocumentUI refDoc\n','',b[name],flags=re.M).split()==a[name].split(),name
for name in ['CloseOwnedDrawing','ReleaseSource']:
 assert re.sub(r'^\s*BeforeCloseUI model\n','',b[name],flags=re.M).split()==a[name].split(),name
assert 'StartUIExperiment' in b['main']
assert b['main'].index('GetOrderedSourceFiles') < b['main'].index('StartUIExperiment')
assert b['main'].index('RestoreAllUI',b['main'].index('PerfNext "UI_RESTORE"')) < b['main'].index('Set actual = CaptureDrawing(targetModel)') < b['main'].index('targetModel.Save3')
assert 'RestoreAllUI\n    PerfFinish "FAILED"' in b['main']
assert 'Finished:\n    On Error Resume Next\n    RestoreAllUI' in b['main']
assert b['RestoreAllUI'].index('uiEnabled = False') < b['RestoreAllUI'].index('TryRestoreUIState')
assert 'If uiStates.Exists(key) Then Exit Sub' in b['SuspendDocumentUI']
assert b['SuspendDocumentUI'].index('uiStates.Add key, state') < b['SuspendDocumentUI'].index('SuspendUIFlag')
assert b['SuspendUIFlag'].index('state(heldKey) = True') < b['SuspendUIFlag'].index('VbLet, False')
assert 'VbLet, CBool(state(previousKey))' in b['TryRestoreUIFlag']
assert b['TryRestoreUIFlag'].index('Require actual = CBool(state(previousKey))') < b['TryRestoreUIFlag'].index('state(heldKey) = False')
assert 'uiStates.Remove key' in b['TryRestoreUIState'] and 'If restored Then' in b['TryRestoreUIState']
assert 'UI STATE RESTORE FAILED' in b['TryRestoreUIState']
for prop in ['EnableFeatureTree','EnableFeatureTreeWindow','EnableConfigurationTree','EnableGraphicsUpdate']:
 assert prop in b['SuspendDocumentUI'] and prop in b['TryRestoreUIState']
lines=s.splitlines()
for i,line in enumerate(lines):
 if 'swApp.CloseDoc ' in line:assert lines[i-1].strip()=='BeforeCloseUI model'
 if 'SuspendDocumentUI refDoc' in line:assert 'Require Not refDoc Is Nothing' in lines[i-1]
assert 'RestoreReferenceUI' in b['BeforeCloseUI']
assert 'CallByName(model, propertyName, VbGet)' in b['TryGetUIManager']
assert b['RebuildDrawing']==a['RebuildDrawing'] and b['AppendDrawing']==a['AppendDrawing']
assert b['main'].count('Set actual = CaptureDrawing(targetModel)')==1
assert (ROOT/'FORM_CODE.txt').read_bytes()==(ROOT.parent/'V16_25/FORM_CODE.txt').read_bytes()
assert max(map(len,s.splitlines()))<1023
reports=[]
for name,encoding in [('MergeDrawingsV16_27.bas','cp1251'),('FORM_CODE.txt','utf-8')]:
 data=(ROOT/name).read_bytes();text=data.decode(encoding);wrapper=1 if name=='FORM_CODE.txt' else 0
 if wrapper:text='Attribute VB_Name = "frmDrawingOrder"\n'+text
 errors=Errors();lexer=vbaLexer(InputStream(text));lexer.removeErrorListeners();lexer.addErrorListener(errors)
 parser=vbaParser(CommonTokenStream(lexer));parser.removeErrorListeners();parser.addErrorListener(errors);parser.startRule()
 reports.append(dict(file=name,sha256=hashlib.sha256(data).hexdigest(),parser_wrapper_lines=wrapper,errors=errors.items))
report=dict(files=reports,preserved_procedures=preserved,static_invariants_passed=True,baseline='V16.25',grammar='grammars-v4/vba/vba7_1',grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',antlr_version='4.13.2',native_vba_compile=False,native_solidworks_run=False,native_ui_restoration_test=False,speedup_measured=False)
(ROOT/'checks/validation.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2));sys.exit(any(x['errors'] for x in reports))
