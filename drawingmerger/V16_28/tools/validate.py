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
s=(ROOT/'MergeDrawingsV16_28.txt').read_text();raw=(ROOT/'MergeDrawingsV16_28.bas').read_bytes();old=(ROOT.parent/'V16_27/MergeDrawingsV16_27.txt').read_text();a,b=procs(old),procs(s)
assert raw.decode('cp1251').replace('\r\n','\n')==s
assert '\n' not in raw.decode('cp1251').replace('\r\n','')
preserved=[]
for name,body in a.items():
 if name not in {'main','AppendDrawing'}:
  assert b[name].replace('V16_28','V16_27').replace('V16.28','V16.27')==body,name;preserved.append(name)
assert re.sub(r'^        rec.Add "ReviewSource.*\n','',b['AppendDrawing'],flags=re.M)==a['AppendDrawing']
main=b['main'];start=main.index('    Set dimensionIssues =');end=main.index('    VerifyDrawing expectedResult',start);branch=main[start:end]
assert 'GoTo Finished' in branch
assert 'WriteStatus "REVIEW_REQUIRED"' in branch
assert 'PerfFinish "REVIEW_REQUIRED"' in branch
assert 'review_views.tsv' in branch
assert 'Save3(' not in branch and 'ExportDrawingToPDF' not in branch and 'PublishPDF' not in branch
assert 'EnsureSheetActive targetModel, firstReviewSheet' in branch
assert main.index('Set actual = CaptureDrawing(targetModel)')<start<end<main.index('targetModel.Save3')<main.index('ExportDrawingToPDF')
assert main.count('Set actual = CaptureDrawing(targetModel)')==1
collect=b['CollectDimensionLosses']
assert 'Array("Sheets", "Views")' in collect
assert 'RequiredDimensionCount(a) - RequiredDimensionCount(b)' in collect
assert 'If missing > 0 Then' in collect
assert 'issues.Add line' in collect and 'ReviewSourceView' in collect
assert 'actual("Order")(i)' in collect
assert 'actual(CStr(group)).Exists' in collect
assert 'REVIEW_REQUIRED.txt' in b['WriteDimensionReview']
assert 'автоматический аудит не пройден' in b['DimensionReviewText']
assert 'Остальные проверки прерваны' in b['DimensionReviewText']
for name in ['VerifyDrawing','VerifyStructure','VerifyAnnotations','SameViewFingerprint','ApprovedDimensionCount','RestoreAllUI']:
 assert b[name]==a[name],name
assert (ROOT/'FORM_CODE.txt').read_bytes()==(ROOT.parent/'V16_27/FORM_CODE.txt').read_bytes()
assert max(map(len,s.splitlines()))<1023
reports=[]
for name,encoding in [('MergeDrawingsV16_28.bas','cp1251'),('FORM_CODE.txt','utf-8')]:
 data=(ROOT/name).read_bytes();text=data.decode(encoding);wrapper=1 if name=='FORM_CODE.txt' else 0
 if wrapper:text='Attribute VB_Name = "frmDrawingOrder"\n'+text
 errors=Errors();lexer=vbaLexer(InputStream(text));lexer.removeErrorListeners();lexer.addErrorListener(errors)
 parser=vbaParser(CommonTokenStream(lexer));parser.removeErrorListeners();parser.addErrorListener(errors);parser.startRule()
 reports.append(dict(file=name,sha256=hashlib.sha256(data).hexdigest(),parser_wrapper_lines=wrapper,errors=errors.items))
report=dict(files=reports,preserved_procedures=preserved,strict_verifiers_unchanged=True,review_branch_bypasses_save_and_pdf=True,static_invariants_passed=True,grammar='grammars-v4/vba/vba7_1',grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',antlr_version='4.13.2',native_vba_compile=False,native_solidworks_run=False,native_manual_review_test=False)
(ROOT/'checks/validation.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2));sys.exit(any(x['errors'] for x in reports))
