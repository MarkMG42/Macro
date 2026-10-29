"""Syntax validation with externally generated grammars-v4 VBA 7.1 Python parser.

Dependencies: antlr4-python3-runtime==4.13.2; parser generated with ANTLR 4.13.2
from grammars-v4 commit 2766c2eed04f285dc763617c9459650892fcf5b9/vba/vba7_1.
Usage: python3 tools/check_vba.py /path/to/generated/parser
Do not replace this with parsing a previously generated grammar report.
"""
from pathlib import Path
import hashlib
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, sys.argv[1])
from antlr4 import FileStream, CommonTokenStream
from antlr4.error.ErrorListener import ErrorListener
from vbaLexer import vbaLexer
from vbaParser import vbaParser


class Errors(ErrorListener):
    def __init__(self):
        self.errors = []

    def syntaxError(self, recognizer, offendingSymbol, line, column, msg, error):
        self.errors.append(dict(line=line, column=column + 1, message=msg))


source = ROOT / 'InsertWeldWithHoles.txt'
errors = Errors()
lexer = vbaLexer(FileStream(str(source), encoding='utf-8'))
lexer.removeErrorListeners()
lexer.addErrorListener(errors)
parser = vbaParser(CommonTokenStream(lexer))
parser.removeErrorListeners()
parser.addErrorListener(errors)
parser.startRule()
report = dict(target=source.name, grammar='grammars-v4/vba/vba7_1',
              grammar_commit='2766c2eed04f285dc763617c9459650892fcf5b9',
              antlr_version='4.13.2', native_vba_compile=False,
              source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
              errors=errors.errors)
(ROOT / 'checks/grammar.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps(report, indent=2))
sys.exit(bool(errors.errors))
