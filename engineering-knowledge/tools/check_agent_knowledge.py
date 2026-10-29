"""Structural checks for repository knowledge; no network or SolidWorks required."""
import re
import json
import hashlib
import zipfile
import subprocess
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / "engineering-knowledge/topics"
errors = []
files = [ROOT / "AGENTS.md", *sorted(BASE.glob("*.md")),
         ROOT / "engineering-knowledge/references/autosizedrilling/README.md",
         ROOT / "engineering-knowledge/AGENTS.md", ROOT / "engineering-knowledge/README.md",
         ROOT / "drawingmerger/README.md", ROOT / "DrillingView/README.md",
         ROOT / "README.md"]
for path in files:
    content = path.read_text(encoding="utf-8")
    for target in re.findall(r"\]\(([^)]+)\)", content):
        if "://" in target or target.startswith("#"):
            continue
        if not (path.parent / target.split("#")[0]).resolve().exists():
            errors.append(f"Broken link: {path.relative_to(ROOT)} -> {target}")
config = tomllib.loads((ROOT / ".codex/config.toml").read_text())
roles = {key: val for key, val in config["agents"].items() if isinstance(val, dict)}
index = (BASE / "README.md").read_text()
for name, role in roles.items():
    path = ROOT / ".codex" / role["config_file"]
    if name not in index:
        errors.append(f"Role missing from catalog: {name}")
    if not path.is_file():
        errors.append(f"Missing role file: {path}")
        continue
    instruction = tomllib.loads(path.read_text()).get("developer_instructions", "")
    sources = re.findall(r"engineering-knowledge/topics/([\w-]+\.md)", instruction)
    if not sources or any(not (BASE / src).is_file() for src in sources):
        errors.append(f"Missing knowledge route: {name}")
if set(roles) != {p.stem for p in (ROOT / "engineering-knowledge/agents").glob("*.toml")}:
    errors.append("Unregistered or missing role config")
snapshot = ROOT / "engineering-knowledge/references/autosizedrilling"
manifest = json.loads((snapshot / "manifest.json").read_text())
archive = snapshot / manifest["archive"]
if hashlib.sha256(archive.read_bytes()).hexdigest() != manifest["archive_sha256"]:
    errors.append("Snapshot ZIP hash mismatch")
with zipfile.ZipFile(archive) as z:
    actual_files = {n for n in z.namelist() if not n.endswith("/")}
    if actual_files != set(manifest["files_sha256"]):
        errors.append("Snapshot file list mismatch")
    if z.testzip() is not None:
        errors.append("Snapshot CRC failure")
    for name, expected in manifest["files_sha256"].items():
        if name not in actual_files or hashlib.sha256(z.read(name)).hexdigest() != expected:
            errors.append(f"Snapshot source hash mismatch: {name}")
validation = json.loads((snapshot / "validation.json").read_text())
if validation["source_commit"] != manifest["source_commit"]:
    errors.append("Validation source commit mismatch")
if validation["native_solidworks_run"] is not False:
    errors.append("Review native-status evidence before changing this assertion")
# Saved macro copies and extracted AutoSizeDrilling must match their immutable sources.
merger = ROOT / "drawingmerger"
preserved = json.loads((merger / "PRESERVED_FILES.json").read_text())
for name, expected in preserved["files_sha256"].items():
    path = merger / name
    if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
        errors.append(f"Saved merger package differs: {name}")
    else:
        original = subprocess.run(
            ["git", "show", preserved["source_commit"] + ":" + name],
            cwd=ROOT, capture_output=True, check=False,
        )
        if original.returncode or original.stdout != path.read_bytes():
            errors.append(f"Preserved merger differs from original Git blob: {name}")
expected_folders = {Path(n).parts[0] for n in preserved["files_sha256"] if n.startswith("V16_")}
if {p.name for p in merger.glob("V16_*") if p.is_dir()} != expected_folders:
    errors.append("Merger version folders differ from original source tree")
for protected in preserved["protected_paths"]:
    paths = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", preserved["source_commit"], "--", protected],
        cwd=ROOT, text=True,
    ).splitlines()
    actual = {str(p.relative_to(ROOT)) for p in (ROOT / protected).rglob("*") if p.is_file()}
    if actual != set(paths):
        errors.append(f"Protected file list differs: {protected}")
    for name in paths:
        path = ROOT / name
        original = subprocess.check_output(["git", "show", preserved["source_commit"] + ":" + name], cwd=ROOT)
        if not path.is_file() or path.read_bytes() != original:
            errors.append(f"Protected file changed: {name}")
if list(ROOT.glob("V16*")) or list(ROOT.glob("SOLIDWORKS_MERGE_*.zip")) or (ROOT / "drawing-merger").exists():
    errors.append("Obsolete merger releases/duplicates remain in root")
for name, expected in manifest["files_sha256"].items():
    path = ROOT / name
    if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
        errors.append(f"Extracted AutoSizeDrilling file differs: {name}")
if errors:
    raise SystemExit("\n".join(errors))
print(f"OK: {len(files)} instruction/knowledge documents; {len(roles)} role routes; preserved merger Git blobs match; {len(actual_files)} snapshot files verified")
print("No native VBA/SolidWorks execution; no guarantee of client role activation")
