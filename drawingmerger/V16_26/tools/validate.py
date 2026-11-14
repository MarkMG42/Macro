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
s=(ROOT/'MergeDrawingsV16_26.txt').read_text();raw=(ROOT/'MergeDrawingsV16_26.bas').read_bytes();old=(ROOT.parent/'V16_25/MergeDrawingsV16_25.txt').read_text();a,b=procs(old),procs(s)
assert raw.decode('cp1251').replace('\r\n','\n')==s
assert '\n' not in raw.decode('cp1251').replace('\r\n','')
changed={'AppendDrawing','MapCopiedSheet','CaptureDrawing','RestoreImportedPlacementAfterRebuild'}
preserved=[]
for name,body in a.items():
 if name not in changed:
  assert b[name].replace('V16_26','V16_25').replace('V16.26','V16.25')==body,name
  preserved.append(name)
# Prove the field reader and the rest of the complete capture are identical.
x=a['CaptureDrawing'];start=x.index('            Set rec = NewDictionary()');end=x.index('            Require Not views.Exists',start);record=x[start:end]
expected='\n'.join(line[8:] for line in record.rstrip().splitlines()).replace('sheet.GetName','sheetName')
assert expected in b['CaptureViewRecord']
assert b['CaptureDrawing'].replace('            Set rec = CaptureViewRecord(view, sheet.GetName)\n',record)==x
assert 'DependencyOrder' not in b['MapCopiedSheet']
assert 'viewMap, stagedOrder' in b['AppendDrawing']
assert 'Set stagedOrder = DependencyOrder(staged("Views"))' in b['AppendDrawing']
oldmap=a['MapCopiedSheet'];newmap=b['MapCopiedSheet']
assert oldmap.split('    Set used = NewDictionary()')[1]==newmap.split('    Set used = NewDictionary()')[1]
route=b['CanFuseImportedSheets'];assert route.count('("LinkedNotes").Count > 0 Then Exit Function')==2
assert 'CStr(parent("Sheet"))' in route
assert 'canFuse = CanFuseImportedSheets(original)' in b['AppendDrawing']
assert 'RestoreImportedLinkedNotes' in b['AppendDrawing']
assert 'CaptureMappedSheets(targetModel' not in b['AppendDrawing']
assert 'CaptureDrawing' not in b['CaptureImportedMetadata']
assert 'CaptureViewRecord(view, targetSheet)' in b['CaptureImportedMetadata']
assert 'views(name)("Label") = section.GetLabel' in b['CaptureImportedMetadata']
assert 'helpers(key)("Name")' in b['CaptureImportedMetadata']
fused=b['ProcessIndependentSheets']
for call in ['RestoreImportedLayout','RestoreImportedSheetProperties','RebuildDrawing','RestoreImportedPlacementAfterRebuild','CaptureImportedMetadata']:
 assert call in fused
assert fused.index('RestoreImportedLayout') < fused.index('RebuildDrawing') < fused.index('RestoreImportedPlacementAfterRebuild') < fused.index('CaptureImportedMetadata')
assert 'Require CanFuseImportedSheets(original)' in fused
assert 'oneSheet.Add CStr(sheetKey), CStr(sheetMap(CStr(sheetKey)))' in fused
assert 'Require SameText(ViewSheetName(view)' in b['RestoreImportedLayout']
assert b['main'].count('Set actual = CaptureDrawing(targetModel)')==1
assert b['main'].index('VerifyDrawing expectedResult') < b['main'].index('targetModel.Save3')
assert b['RebuildDrawing']==a['RebuildDrawing']
assert (ROOT/'FORM_CODE.txt').read_bytes()==(ROOT.parent/'V16_25/FORM_CODE.txt').read_bytes()
assert max(map(len,s.splitlines()))<1023
reports=[]
for name,encoding in [('MergeDrawingsV16_26.bas','cp1251'),('FORM_CODE.txt','utf-8')]:
 data=(ROOT/name).read_bytes();text=data.decode(encoding);wrapper=1 if name=='FORM_CODE.txt' else 0
 if wrapper:text='Attribute VB_Name = "frmDrawingOrder"\n'+text
 errors=Errors();lexer=vbaLexer(InputStream(text));lexer.removeErrorListeners();lexer.addErrorListener(errors)
 parser=vbaParser(CommonTokenStream(lexer));parser.removeErrorListeners();parser.addErrorListener(errors);parser.startRule()
 reports.append(dict(file=name,sha256=hashlib.sha256(data).hexdigest(),parser_wrapper_lines=wrapper,errors=errors.items))
report=dict(files=reports,preserved_procedures=preserved,full_reader_extraction_equivalent=True,static_invariants_passed=True,grammar='grammars-v4/vba/vba7_1',grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',antlr_version='4.13.2',native_vba_compile=False,native_solidworks_run=False,speedup_measured=False)
(ROOT/'checks/validation.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps(report,indent=2));sys.exit(any(x['errors'] for x in reports))
