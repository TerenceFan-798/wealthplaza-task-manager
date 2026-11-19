# -*- coding: utf-8 -*-
"""无需第三方依赖：直接解压 xlsx 读取指定列，输出 TSV 到工作区。"""
import zipfile, re, sys, json
from xml.etree import ElementTree as ET

SRC = r"C:\Users\Vincent\Desktop\11月浏览任务.xlsx"
OUT = r"C:\Users\Vincent\WorkBuddy\2026-11-18-16-08-38\task_data.tsv"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def col_letters(ref):
    return re.match(r"([A-Z]+)", ref).group(1)


def load_shared_strings(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall(NS + "si"):
        out.append("".join(t.text or "" for t in si.iter(NS + "t")))
    return out


def sheet_files(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2target = {r.get("Id"): r.get("Target") for r in rels}
    result = []
    for sh in wb.find(NS + "sheets"):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rid2target[rid]
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        result.append((sh.get("name"), target))
    return result


with zipfile.ZipFile(SRC) as z:
    sst = load_shared_strings(z)
    lines = []
    for name, path in sheet_files(z):
        root = ET.fromstring(z.read(path))
        lines.append("#SHEET\t" + name)
        for row in root.iter(NS + "row"):
            cells = {}
            for c in row.findall(NS + "c"):
                ref = c.get("r") or ""
                col = col_letters(ref) if ref else ""
                if col not in ("A", "B", "C"):
                    continue
                t = c.get("t")
                v = c.find(NS + "v")
                isel = c.find(NS + "is")
                if t == "s" and v is not None:
                    val = sst[int(v.text)]
                elif t == "inlineStr" and isel is not None:
                    val = "".join(x.text or "" for x in isel.iter(NS + "t"))
                else:
                    val = v.text if v is not None else ""
                cells[col] = (val or "").strip()
            lines.append("\t".join([row.get("r") or "", cells.get("A", ""), cells.get("B", ""), cells.get("C", "")]))

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("rows written:", len(lines), "->", OUT)
