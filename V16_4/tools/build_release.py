"""Assemble the standalone module. Writes only inside macro/v16."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"


def main():
    header = (TOOLS / "v12_header.txt").read_text()
    core = (TOOLS / "merge_core.vba").read_text()
    ui = (TOOLS / "v12_ui.txt").read_text()
    helpers = (TOOLS / "v12_helpers.txt").read_text()
    # Complete the existing folder-selection UI; retain the drag/drop behavior.
    ui = ui.replace(
        "<div class='buttons'><button onclick='finish(true)'>",
        "<div class='buttons'><button onclick='addFolder()'>Добавить папку</button><button onclick='finish(true)'>",
    )
    ui = ui.replace(
        '"function renumber()',
        '"function addFolder(){var sh=new ActiveXObject(\'Shell.Application\');var fd=sh.BrowseForFolder(0,\'Выберите папку с чертежами\',0,0);if(!fd)return;var dir=fso.GetFolder(fd.Self.Path);var seen={};var rows=ul.getElementsByTagName(\'li\');for(var j=0;j<rows.length;j++)seen[norm(rows[j].getAttribute(\'data-path\'))]=true;var e=new Enumerator(dir.Files);for(;!e.atEnd();e.moveNext()){var f=e.item();var p=f.Path;if(f.Name.substr(0,2)!=\'~$\'&&fso.GetExtensionName(p).toLowerCase()==\'slddrw\'&&!seen[norm(p)]){addItem(p);seen[norm(p)]=true;}}renumber();}" & _\n'
        '        "function renumber()',
    )
    # Publish the result only after the HTA closes the file; prevents partial reads.
    ui = ui.replace("fso.CreateTextFile(outputPath,true,true)", "fso.CreateTextFile(outputPath+'.writing',true,true)")
    ui = ui.replace("w.Close();window.close();", "w.Close();fso.MoveFile(outputPath+'.writing',outputPath);window.close();")
    ui = ui.replace(
        'Debug.Print "Путь пропущен, файл не найден: " & resultText',
        'Err.Raise vbObjectError + 1601, , "Файл исчез из списка: " & resultText',
    )
    # Do not retain unused PDF-path dialog: outputs are confined to this run.
    start = helpers.index("Private Function PickPdfPath(")
    end = helpers.index("Private Function GetFolder(", start)
    helpers = helpers[:start] + helpers[end:]
    # Keep native export implementation unchanged; a new ExportPdfData has all sheets by default.
    module = header.rstrip() + "\n\n" + core.rstrip() + "\n\n" + ui.rstrip() + "\n\n" + helpers.strip() + "\n"
    (ROOT / "MERGE_DRAWINGS_V16.txt").write_bytes(module.replace("\n", "\r\n").encode("utf-8-sig"))
    bas = 'Attribute VB_Name = "MergeDrawingsV16"\n' + module
    (ROOT / "MergeDrawingsV16.bas").write_bytes(bas.replace("\n", "\r\n").encode("cp1251"))
    print(f"Built standalone TXT/BAS: {len(module.splitlines())} lines")


if __name__ == "__main__":
    main()
