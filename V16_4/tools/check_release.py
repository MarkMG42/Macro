"""Static checks only; this is NOT a SOLIDWORKS/VBA execution test.

Run from any cwd. ANTLR and interop metadata live in the local research folder;
they are intentionally not included in the end-user archive.
"""
from pathlib import Path
import hashlib
import json
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
RESEARCH = ROOT / "research"


def main():
    txt = (ROOT / "MERGE_DRAWINGS_V16.txt").read_text(encoding="utf-8-sig")
    bas = (ROOT / "MergeDrawingsV16.bas").read_text(encoding="cp1251")
    assert bas == 'Attribute VB_Name = "MergeDrawingsV16"\n' + txt
    assert len(max(bas.splitlines(), key=len)) < 1023
    continuation = maximum = 0
    for line in bas.splitlines():
        continuation = continuation + 1 if line.rstrip().endswith(" _") else 0
        maximum = max(maximum, continuation)
    assert maximum <= 24, maximum
    procedures = re.findall(r"^(?:Private|Public) (?:Sub|Function) (\w+)", txt, re.M)
    assert len(procedures) == len(set(n.lower() for n in procedures))
    assert not re.search(r"\b(DeleteFile|DeleteFolder|Kill|RmDir|EditPaste|SetBaseView)\b", re.sub(r"(?m)^\s*'.*", "", txt))
    original = (ROOT.parent / "outputs" / "SOLIDWORKS_MERGE_V12_WORKING_ORIGINAL.txt").read_text(encoding="utf-8-sig")
    export = lambda s: re.search(r"Private Function ExportDrawingToPDF\b[\s\S]*?End Function", s).group()
    assert export(txt) == export(original), "V12 PDF implementation changed"
    manifest = json.loads((RESEARCH / "input_manifest.json").read_text())
    # Accept the saved inventory's list/object format without altering it.
    entries = manifest if isinstance(manifest, list) else manifest.get("files", manifest)
    if isinstance(entries, dict):
        entries = [dict(path=k, sha256=v if isinstance(v, str) else v["sha256"]) for k, v in entries.items()]
    for item in entries:
        path = Path(item["path"])
        if not path.is_absolute():
            path = ROOT.parent / path
        assert hashlib.sha256(path.read_bytes()).hexdigest() == item["sha256"], path
    grammar = RESEARCH / "parser"
    sys.path[:0] = [str(grammar / "runtime"), str(grammar)]
    from antlr4 import InputStream, CommonTokenStream
    from antlr4.error.ErrorListener import ErrorListener
    from vbaLexer import vbaLexer
    from vbaParser import vbaParser
    errors = []

    class Listener(ErrorListener):
        def syntaxError(self, recognizer, offendingSymbol, line, column, msg, e):
            errors.append(dict(line=line, column=column, message=msg))

    lexer = vbaLexer(InputStream(bas))
    lexer.removeErrorListeners()
    lexer.addErrorListener(Listener())
    parser = vbaParser(CommonTokenStream(lexer))
    parser.removeErrorListeners()
    parser.addErrorListener(Listener())
    parser.startRule()
    result = dict(parser="ANTLR 4.13.2 / grammars-v4 VBA 7.1", errors=errors,
                  lines=len(bas.splitlines()), max_line_length=max(map(len, bas.splitlines())),
                  max_continuations=maximum, procedures=len(procedures),
                  original_files_unchanged=len(entries), pdf_matches_v12=True,
                  bas_txt_equivalent=True)
    (RESEARCH / "static_results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    assert not errors, "VBA grammar parse failed"


if __name__ == "__main__":
    main()
