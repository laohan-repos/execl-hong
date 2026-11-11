#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 customUI.xml 注入到 XLAM/XLSM（OPC/Open XML 包）中。

跨平台实现，PowerShell 版见 Build-ExcelToolbox.ps1 的 Add-CustomUIPart：
两者写入的是完全相同的三个部件，逻辑互相校验。

用法：
    python build/inject_ribbon.py <package.xlam> <customUI.xml> [--check-only]

注入内容：
  1. 新增部件  customUI/customUI.xml
  2. _rels/.rels 追加 ui/extensibility 关系
  3. [Content_Types].xml 追加 application/xml Override

不会改动 xl/vbaProject.bin、xl/workbook.xml 等任何已有部件的字节。
"""
from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

PART_NAME = "customUI/customUI.xml"
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RIBBON_NS_2006 = "http://schemas.microsoft.com/office/2006/01/customui"
RIBBON_NS_2009 = "http://schemas.microsoft.com/office/2009/07/customui"


class PackageError(RuntimeError):
    pass


def validate_ribbon_xml(data: bytes) -> list[str]:
    """校验 Ribbon XML 是否合法，返回其中声明的 onAction 回调名。"""
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    root = ET.fromstring(data)
    ns = root.tag.split("}")[0].lstrip("{")
    if ns not in (RIBBON_NS_2006, RIBBON_NS_2009):
        raise PackageError(f"customUI 命名空间不正确: {ns}")
    ribbon = root.find(f"{{{ns}}}ribbon")
    if ribbon is None:
        raise PackageError("缺少 <ribbon> 节点")
    tabs = ribbon.find(f"{{{ns}}}tabs")
    if tabs is None:
        raise PackageError("缺少 <tabs> 节点")
    callbacks: list[str] = []
    labels: list[str] = []
    for tab in tabs.findall(f"{{{ns}}}tab"):
        labels.append(tab.get("label", ""))
        for button in tab.iter(f"{{{ns}}}button"):
            if button.get("onAction"):
                callbacks.append(button.get("onAction"))
    if not callbacks:
        raise PackageError("没有找到任何 onAction 回调")
    on_load = root.get("onLoad")
    if on_load:
        callbacks.append(on_load)
    return callbacks


def inject(package: Path, ribbon_xml: Path, check_only: bool = False) -> dict:
    if not package.is_file():
        raise PackageError(f"找不到包文件: {package}")
    if not ribbon_xml.is_file():
        raise PackageError(f"找不到 Ribbon XML: {ribbon_xml}")

    ui_data = ribbon_xml.read_bytes()
    callbacks = validate_ribbon_xml(ui_data)

    with zipfile.ZipFile(package, "r") as zin:
        names = zin.namelist()
        for required in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"):
            if required not in names:
                raise PackageError(f"不是有效的 Excel OPC 包，缺少 {required}")
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]

    has_vba = any(n == "xl/vbaProject.bin" for n in names)
    by_name = {info.filename: data for info, data in entries}

    # --- [Content_Types].xml ---
    ct_root = ET.fromstring(by_name["[Content_Types].xml"])
    existing_overrides = {o.get("PartName") for o in ct_root.findall(f"{{{CT_NS}}}Override")}
    has_ct = "/" + PART_NAME in existing_overrides

    # --- _rels/.rels ---
    rel_root = ET.fromstring(by_name["_rels/.rels"])
    rel_ids = {r.get("Id") for r in rel_root.findall(f"{{{REL_NS}}}Relationship")}
    rel_match = [
        r for r in rel_root.findall(f"{{{REL_NS}}}Relationship")
        if r.get("Type") == REL_TYPE and r.get("Target", "").lstrip("/") == PART_NAME
    ]
    has_rel = bool(rel_match)
    has_part = PART_NAME in names

    if check_only:
        return {
            "part": has_part, "relationship": has_rel, "content_type": has_ct,
            "vba_project": has_vba, "callbacks": callbacks, "parts": len(names),
        }

    replacements: dict[str, bytes] = {PART_NAME: ui_data}

    # --- 构造新的 [Content_Types].xml ---
    if not has_ct:
        ov = ET.Element(f"{{{CT_NS}}}Override")
        ov.set("PartName", "/" + PART_NAME)
        ov.set("ContentType", "application/xml")
        ct_root.append(ov)
        ET.register_namespace("", CT_NS)
        replacements["[Content_Types].xml"] = ET.tostring(
            ct_root, encoding="utf-8", xml_declaration=True
        )

    # --- 构造新的 _rels/.rels ---
    if not has_rel:
        rid = "rIdExcelToolboxUI"
        i = 1
        while rid in rel_ids:
            rid = f"rIdExcelToolboxUI{i}"
            i += 1
        rel = ET.Element(f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", REL_TYPE)
        rel.set("Target", PART_NAME)
        rel_root.append(rel)
        ET.register_namespace("", REL_NS)
        replacements["_rels/.rels"] = ET.tostring(
            rel_root, encoding="utf-8", xml_declaration=True
        )

    # --- 重写包（保留原有 ZipInfo 元数据与部件顺序，一次成型） ---
    tmp = package.with_suffix(package.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, data in entries:
            zout.writestr(info, replacements.get(info.filename, data))
        if PART_NAME not in names:
            zout.writestr(PART_NAME, ui_data, compress_type=zipfile.ZIP_DEFLATED)

    shutil.move(str(tmp), str(package))

    result = check(package)
    result["callbacks"] = callbacks
    return result


def check(package: Path) -> dict:
    with zipfile.ZipFile(package, "r") as z:
        names = z.namelist()
        if z.testzip() is not None:
            raise PackageError("ZIP 校验失败：存在损坏的部件")
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        rels = ET.fromstring(z.read("_rels/.rels"))
        ui = z.read(PART_NAME) if PART_NAME in names else None

        has_ct = any(
            o.get("PartName") == "/" + PART_NAME for o in ct.findall(f"{{{CT_NS}}}Override")
        )
        has_rel = any(
            r.get("Type") == REL_TYPE and r.get("Target", "").lstrip("/") == PART_NAME
            for r in rels.findall(f"{{{REL_NS}}}Relationship")
        )
        if ui is not None:
            ET.fromstring(ui)  # well-formedness
        # 所有 xml 部件都必须能解析
        for n in names:
            if n.endswith((".xml", ".rels")):
                ET.fromstring(z.read(n))
        return {
            "part": PART_NAME in names,
            "relationship": has_rel,
            "content_type": has_ct,
            "vba_project": "xl/vbaProject.bin" in names,
            "parts": len(names),
            "size": package.stat().st_size,
        }


def main() -> int:
    ap = argparse.ArgumentParser(description="注入并校验 Excel Ribbon customUI 部件")
    ap.add_argument("package", type=Path)
    ap.add_argument("ribbon_xml", type=Path)
    ap.add_argument("--check-only", action="store_true")
    args = ap.parse_args()

    try:
        res = inject(args.package, args.ribbon_xml, args.check_only)
    except (PackageError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(f"[FAIL] {exc}")
        return 1

    ok = res["part"] and res["relationship"] and res["content_type"]
    print(f"包文件       : {args.package}")
    print(f"部件总数     : {res['parts']}")
    print(f"customUI 部件: {'OK' if res['part'] else 'MISSING'}")
    print(f"包关系       : {'OK' if res['relationship'] else 'MISSING'}")
    print(f"内容类型     : {'OK' if res['content_type'] else 'MISSING'}")
    print(f"vbaProject   : {'OK' if res['vba_project'] else '缺失（需 Excel 编译生成）'}")
    if res.get("callbacks"):
        print(f"回调         : {', '.join(res['callbacks'])}")
    print(f"所有 XML 部件: 均通过 well-formed 校验")
    print(("[OK] 校验通过" if ok else "[FAIL] 注入不完整"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
