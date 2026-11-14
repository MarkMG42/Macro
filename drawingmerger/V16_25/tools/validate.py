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
s=(ROOT/'MergeDrawingsV16_25.txt').read_text();raw=(ROOT/'MergeDrawingsV16_25.bas').read_bytes();old=(ROOT.parent/'V16_24/MergeDrawingsV16_24.txt').read_text();a,b=procs(old),procs(s)
assert raw.decode('cp1251').replace('\r\n','\n')==s
assert '\n' not in raw.decode('cp1251').replace('\r\n','')
changed={'main','AppendDrawing','CloneInput','PrepareSource','RestoreImportedLinkedNotes','RebuildDrawing'}
preserved=[]
for name,body in a.items():
 if name not in changed:
  assert b[name].replace('DrawingMerger_V16_25','DrawingMerger_V16_24')==body,name
  preserved.append(name)
assert 'CaptureDrawing(original)' not in b['CloneInput']
assert 'OpenDoc6' not in b['CloneInput'] and 'CloseDoc' not in b['CloneInput']
assert 'VerifyDrawing baseline, actual' in b['CloneInput']
prep=b['PrepareSource']
assert prep.index('Set copy = CloneInput') < prep.index('ReleaseSource model, True') < prep.index('Set model = copy')
assert 'CloseOwnedDrawing copy' in prep
notes=b['RestoreImportedLinkedNotes']
assert notes.index('("LinkedNotes").Count > 0') < notes.index('Set view = FindView')
assert notes.index('GoTo NextNoteSheet') < notes.index('EnsureSheetActive')
rebuild=b['RebuildDrawing']
assert 'model.EditRebuild3()' in rebuild
assert 'If FORCE_ALL_REBUILDS Then' in rebuild
assert rebuild.index('model.EditRebuild3()') < rebuild.index('LogLine "REBUILD fallback=force')
assert 'swDocDRAWING' in rebuild and 'Require forced' in rebuild
main=b['main'];assert main.index('RestoreTargetGraphics') < main.index('Set actual = CaptureDrawing(targetModel)') < main.index('targetModel.Save3') < main.index('ExportDrawingToPDF')
assert main.count('Set actual = CaptureDrawing(targetModel)')==1
assert main.count('RestoreTargetGraphics')==3
assert 'graphicsHeld = True' in b['SuspendTargetGraphics']
assert b['SuspendTargetGraphics'].index('graphicsHeld = True') < b['SuspendTargetGraphics'].index('graphicsView.EnableGraphicsUpdate = False')
assert 'graphicsView.EnableGraphicsUpdate = graphicsPrevious' in b['RestoreTargetGraphics']
assert 'GRAPHICS RESTORE FAILED' in b['RestoreTargetGraphics']
assert 'DetailFinish "FAILED"' in b['AppendDrawing']
# Remove instrumentation and prove geometry transfer body stayed identical.
clean=re.sub(r'^    Detail(?:Start|Next|Finish) .*\n','',b['AppendDrawing'],flags=re.M)
assert clean==a['AppendDrawing']
assert b['CaptureDrawing']==a['CaptureDrawing'] and b['VerifyDrawing']==a['VerifyDrawing']
assert b['CaptureStructure']==a['CaptureStructure']
assert max(map(len,s.splitlines()))<1023
assert (ROOT/'FORM_CODE.txt').read_bytes()==(ROOT.parent/'V16_24/FORM_CODE.txt').read_bytes()
reports=[]
for name,encoding in [('MergeDrawingsV16_25.bas','cp1251'),('FORM_CODE.txt','utf-8')]:
 data=(ROOT/name).read_bytes();text=data.decode(encoding);wrapper=1 if name=='FORM_CODE.txt' else 0
 if wrapper:text='Attribute VB_Name = "frmDrawingOrder"\n'+text
 errors=Errors();lexer=vbaLexer(InputStream(text));lexer.removeErrorListeners();lexer.addErrorListener(errors)
 parser=vbaParser(CommonTokenStream(lexer));parser.removeErrorListeners();parser.addErrorListener(errors);parser.startRule()
 reports.append(dict(file=name,sha256=hashlib.sha256(data).hexdigest(),parser_wrapper_lines=wrapper,errors=errors.items))
report=dict(files=reports,preserved_procedures=preserved,static_invariants_passed=True,grammar='grammars-v4/vba/vba7_1',grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',antlr_version='4.13.2',native_vba_compile=False,native_solidworks_run=False,speedup_measured=False)
(ROOT/'checks/validation.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2));sys.exit(any(x['errors'] for x in reports))
