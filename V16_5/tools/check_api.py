"""Check declared SOLIDWORKS member names against vendor interop metadata.

This checks existence, accessors and the ABI parameter counts in an explicit
call contract. It is not a VBA compiler or a native COM invocation.
"""
from pathlib import Path
import json
import re

ROOT = Path(__file__).resolve().parents[1]


def main():
    text = (ROOT / "MERGE_DRAWINGS_V16.txt").read_text(encoding="utf-8-sig")
    metadata = json.loads((ROOT / "research/interop_signatures.json").read_text())
    # Ignore string literals and comments, while preserving line boundaries.
    code = re.sub(r'"(?:[^"]|"")*"', '""', text)
    code = re.sub(r"'[^\n]*", "", code)
    code = code.replace(" _\n", " ")
    chunks = re.split(r"(?m)^(?=(?:Private|Public) (?:Sub|Function) )", code)
    declare = re.compile(r"\b(\w+)\s+As\s+SldWorks\.(\w+)", re.I)
    globals_ = {m.group(1).lower(): "I" + m.group(2) for m in declare.finditer(chunks[0])}
    checked = set()
    missing = []
    for chunk in chunks[1:]:
        types = dict(globals_)
        types.update({m.group(1).lower(): "I" + m.group(2) for m in declare.finditer(chunk)})
        # These variables are deliberately late-bound in the original V12 PDF function.
        if "Function ExportDrawingToPDF" in chunk:
            types.update(pdfdatalocal="IExportPdfData", swapplocal="ISldWorks", swmodel="IModelDoc2", swdrawlocal="IDrawingDoc")
        for m in re.finditer(r"\b(\w+)\.(\w+)(?:\.(\w+))?", chunk):
            var, member, child = m.groups()
            kind = types.get(var.lower())
            if not kind:
                continue
            members = {k.lower(): k for k in metadata[kind]}
            key = members.get(member.lower()) or members.get("get_" + member.lower())
            if not key and var.lower() == "swmodel" and "Function ExportDrawingToPDF" in chunk:
                # V12 intentionally uses the drawing's IDispatch for both interfaces.
                kind = "IDrawingDoc"
                members = {k.lower(): k for k in metadata[kind]}
                key = members.get(member.lower()) or members.get("get_" + member.lower())
            if not key:
                missing.append(f"{kind}.{member}")
                continue
            checked.add((kind, key))
            suffix = chunk[m.end():]
            prefix = chunk[chunk.rfind("\n", 0, m.start()) + 1:m.start()]
            is_assignment = not prefix.strip() or bool(re.search(r"\bThen\s*$", prefix, re.I))
            if key.startswith("get_") and not child and is_assignment and re.match(r"\s*=", suffix):
                setter = members.get("set_" + member.lower())
                if not setter:
                    missing.append(f"{kind}.set_{member}")
                else:
                    checked.add((kind, setter))
            if child:
                child_kind = {"extension": "IModelDocExtension", "getcurrentsheet": "ISheet"}.get(member.lower())
                if not child_kind:
                    missing.append(f"unknown chain {kind}.{member}.{child}")
                else:
                    child_members = {k.lower(): k for k in metadata[child_kind]}
                    child_key = child_members.get(child.lower()) or child_members.get("get_" + child.lower())
                    if child_key:
                        checked.add((child_kind, child_key))
                    else:
                        missing.append(f"{child_kind}.{child}")
    contracts = {
        "ISldWorks.OpenDoc6": 6, "ISldWorks.ActivateDoc3": 4,
        "ISldWorks.CloseDoc": 1, "ISldWorks.GetOpenDocumentByName": 1,
        "ISldWorks.SetCurrentWorkingDirectory": 1, "ISldWorks.GetExportFileData": 1,
        "IDrawingDoc.PasteSheet": 2, "IDrawingDoc.ActivateSheet": 1,
        "IDrawingDoc.get_Sheet": 1, "IModelDocExtension.SelectByID2": 9,
        "IModelDocExtension.SaveAs": 6, "IModelDoc2.Save3": 3,
        "IModelDoc2.ForceRebuild3": 1, "IModelDoc2.ClearSelection2": 1,
        "IModelDoc2.Paste": 0, "IModelDoc2.EditCut": 0, "IModelDoc2.EditCopy": 0,
    }
    for dotted, count in contracts.items():
        kind, member = dotted.split(".")
        params = [p for p in metadata[kind][member]["parameters"] if p["sequence"] > 0]
        assert len(params) == count, (dotted, params)
    report = dict(vendor="SolidWorks.Interop.sldworks 32.1.0 (SW 2024 ABI)",
                  checked_member_count=len(checked), checked_call_contracts=len(contracts),
                  missing=missing, members=[f"{a}.{b}" for a, b in sorted(checked)])
    (ROOT / "research/api_results.json").write_text(json.dumps(report, indent=2))
    print(json.dumps(report, indent=2))
    assert not missing


if __name__ == "__main__":
    main()
