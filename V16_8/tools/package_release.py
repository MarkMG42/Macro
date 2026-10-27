"""Create the reviewed, small release; exclude downloaded runtimes and CAD DLLs."""
from pathlib import Path
import hashlib
import json
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def main():
    checks = ROOT / "checks"
    checks.mkdir(exist_ok=True)
    for name in ("static_results.json", "api_results.json", "constant_results.json", "hta_results.json", "input_manifest.json"):
        (checks / name).write_bytes((ROOT / "research" / name).read_bytes())
    dependencies = {}
    for name in ("parser/vbaLexer.g4", "parser/vbaParser.g4", "sldworks32.1.nupkg", "swconst32.1.nupkg"):
        dependencies[name] = hashlib.sha256((ROOT / "research" / name).read_bytes()).hexdigest()
    (checks / "verification_dependencies.json").write_text(json.dumps(dependencies, indent=2))
    files = [ROOT / name for name in ("MergeDrawingsV16.bas", "MERGE_DRAWINGS_V16.txt", "README_RU.md", "RESEARCH.md", "VALIDATION.md", "STATUS.md")]
    files += sorted((ROOT / "tools").glob("*.py"))
    files += [ROOT / "tools" / name for name in ("merge_core.vba", "v12_header.txt", "v12_helpers.txt", "v12_ui.txt")]
    files += sorted(checks.glob("*.json"))
    lines = [hashlib.sha256(p.read_bytes()).hexdigest() + "  " + str(p.relative_to(ROOT)).replace("\\", "/") for p in files]
    sums = ROOT / "SHA256SUMS.txt"
    sums.write_text("\n".join(lines) + "\n")
    files.append(sums)
    archive = ROOT.parent / "outputs" / "SOLIDWORKS_MERGE_V16_LINKED_VIEWS.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for path in files:
            z.write(path, "V16/" + str(path.relative_to(ROOT)).replace("\\", "/"))
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        assert len(z.namelist()) == len(files)
        for path in files:
            assert z.read("V16/" + str(path.relative_to(ROOT)).replace("\\", "/")) == path.read_bytes()
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(digest + "  " + archive.name + "\n")
    (ROOT / "research" / "published_files.json").write_text(json.dumps([str(p.relative_to(ROOT)) for p in files]))
    print(json.dumps(dict(archive=str(archive), bytes=archive.stat().st_size, files=len(files), sha256=digest), ensure_ascii=False))


if __name__ == "__main__":
    main()
