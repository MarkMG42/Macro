from pathlib import Path
import sys,re,json,hashlib
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,sys.argv[1])
from antlr4 import InputStream,CommonTokenStream
from antlr4.error.ErrorListener import ErrorListener
from vbaLexer import vbaLexer
from vbaParser import vbaParser
class Errors(ErrorListener):
 def __init__(self):self.items=[]
 def syntaxError(self,recognizer,symbol,line,column,msg,error):self.items.append(dict(line=line,column=column+1,message=msg))
def procs(s):
 return {m[1]:m[0] for m in re.finditer(r'^(?:Private|Public) (?:Sub|Function) (\w+)[\s\S]*?^End (?:Sub|Function)',s,re.M)}
s=(ROOT/'MergeDrawingsV16_24.txt').read_text();raw=(ROOT/'MergeDrawingsV16_24.bas').read_bytes()
assert raw.decode('cp1251').replace('\r\n','\n')==s
assert '\n' not in raw.decode('cp1251').replace('\r\n','')
old=(ROOT.parent/'V16_23/MergeDrawingsV16_23.txt').read_text();a,b=procs(old),procs(s)
changed={'InitFailureLog','CopyWholeSheet','main','FindView','MoveView','RestoreImportedLinkedNotes','RebuildDrawing','CloneInput','EnsureSheetActive','RequireActive','CloseOwnedDrawing','ReleaseSource','LocalizeTargetPropertyView'}
preserved=[]
for name,body in a.items():
 if name not in changed:
  assert b[name]==body,name;preserved.append(name)
assert b['InitFailureLog'].replace('V16_24','V16_23')==a['InitFailureLog']
assert b['VerifyStructure']==a['VerifyStructure']
assert b['CaptureDrawing']==a['CaptureDrawing']
assert b['MoveView'].replace('    InvalidateViewIndex\n','').replace('CaptureStructure(model, touched)','CaptureDrawing(model, "", touched)')==a['MoveView']
for name in ['CopyWholeSheet','RebuildDrawing','CloneInput','EnsureSheetActive','RequireActive','CloseOwnedDrawing','ReleaseSource','LocalizeTargetPropertyView']:
 assert re.sub(r'^\s*InvalidateViewIndex\n','',b[name],flags=re.M).split()==a[name].split(),name
assert 'expected.Count = 0 Then Exit Sub' in b['RestoreViewLinkedNotes']
assert 'expected.Count' not in b['RestoreImportedLinkedNotes']
assert 'original("Views").Keys' in b['RestoreImportedLinkedNotes']
for field in ['Type','Model','Configuration','Parent','Dimensions','ApprovedOmissions']:
 assert 'rec.Add "'+field+'"' in b['CaptureStructure']
for expensive in ['CaptureLinkedNotes','AddAnnotationStats','GetNotes','GetOutline','GetFaceHatchCount']:
 assert expensive not in b['CaptureStructure']
lines=s.splitlines()
for i,line in enumerate(lines):
 if any(t in line for t in ['model.EditCut','model.Paste','ok = targetDraw.PasteSheet(','forced = model.ForceRebuild3(','apiResult = drawing.ActivateSheet(','Set active = swApp.ActivateDoc3(','Set view = targetDraw.CreateDrawViewFromModelView3(']):
  assert lines[i-1].strip()=='InvalidateViewIndex',line
assert b['main'].count('Set actual = CaptureDrawing(targetModel)')==1
assert b['main'].index('VerifyDrawing expectedResult') < b['main'].index('targetModel.Save3')
assert s.count('model.ForceRebuild3(False)')==1
assert (ROOT/'FORM_CODE.txt').read_bytes()==(ROOT.parent/'V16_23/FORM_CODE.txt').read_bytes()
# Timing is isolated to main plus new helpers; preserve the optimization checks.
assert 'As LongLong' in s and 'GetTickCount64' in b['PerfClock']
main=b['main']
assert main.index('GetOrderedSourceFiles') < main.index('PerfStart') < main.index('CheckSavedInput CStr(sourcePath)')
assert main.index('PerfNext "FINAL_AUDIT"') < main.index('Set actual = CaptureDrawing(targetModel)')
assert main.index('WriteStatus "SUCCESS"') < main.index('PerfFinish "SUCCESS"') < main.index('MsgBox "выполнено успешно"')
assert main.index('failureNumber = Err.Number') < main.index('PerfFinish "FAILED"')
assert 'perfActive = False' in b['PerfFinish']
assert 'TIMING.tsv' in b['PerfFinish']
assert 'On Error GoTo Failed' in b['PerfFinish']
for stage in ['CHECK_INPUTS','PREPARE_MAIN','FINAL_AUDIT','SAVE3','PDF_EXPORT','PDF_PUBLISH']:
 assert 'PerfNext "'+stage+'"' in main
assert 'PerfNext "SOURCE_"' in main
assert 'Timer' not in b['PerfClock'] and 'Now' not in b['PerfClock']
reports=[]
for name,encoding in [('MergeDrawingsV16_24.bas','cp1251'),('FORM_CODE.txt','utf-8')]:
 data=(ROOT/name).read_bytes();text=data.decode(encoding);wrapper=1 if name=='FORM_CODE.txt' else 0
 if wrapper:text='Attribute VB_Name = "frmDrawingOrder"\n'+text
 err=Errors();lexer=vbaLexer(InputStream(text));lexer.removeErrorListeners();lexer.addErrorListener(err)
 parser=vbaParser(CommonTokenStream(lexer));parser.removeErrorListeners();parser.addErrorListener(err);parser.startRule()
 reports.append(dict(file=name,sha256=hashlib.sha256(data).hexdigest(),parser_wrapper_lines=wrapper,errors=err.items))
report=dict(files=reports,preserved_procedures=preserved,scope_1_2_4_passed=True,grammar='grammars-v4/vba/vba7_1',grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',antlr_version='4.13.2',native_vba_compile=False,native_solidworks_run=False,speedup_measured=False,timer_api="GetTickCount64",timer_native_test=False)
(ROOT/'checks/validation.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2))
sys.exit(any(r['errors'] for r in reports))
