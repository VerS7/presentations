"""
manage_slide.py - Standalone OOXML PPTX Packager & Slide Manager
Handles unpack, inspect, duplicate, and repack without external dependencies.
"""

import os
import sys
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'
}

def unpack_pptx(pptx_path: str, target_dir: str):
    """Extracts a PPTX file into a directory."""
    os.makedirs(target_dir, exist_ok=True)
    with zipfile.ZipFile(pptx_path, 'r') as z:
        z.extractall(target_dir)
    print(f"[OK] Unpacked {pptx_path} into {target_dir}")

def repack_pptx(source_dir: str, pptx_out: str):
    """Packs a directory back into a valid .pptx archive."""
    os.makedirs(os.path.dirname(os.path.abspath(pptx_out)), exist_ok=True)
    with zipfile.ZipFile(pptx_out, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, source_dir)
                z.write(abs_path, rel_path)
    print(f"[OK] Repacked {source_dir} into {pptx_out}")

def inspect_unpacked(source_dir: str):
    """Lists all slides and their titles/word counts."""
    p_dir = Path(source_dir)
    pres_xml = p_dir / "ppt" / "presentation.xml"
    if not pres_xml.exists():
        print(f"[ERROR] Not a valid unpacked presentation: {pres_xml} not found")
        return

    tree = ET.parse(pres_xml)
    root = tree.getroot()
    sld_ids = root.findall('.//p:sldId', NS)
    print(f"Total slides registered in presentation.xml: {len(sld_ids)}")
    for i, sld in enumerate(sld_ids, 1):
        rid = sld.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
        sid = sld.attrib.get('id', '')
        print(f"  Slide {i:2d} -> rId: {rid}, slideId: {sid}")

def duplicate_slide(source_dir: str, source_slide: str):
    """Duplicates a slide file (e.g. slide2.xml) and registers it in the manifest."""
    p_dir = Path(source_dir)
    slides_dir = p_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    pres_xml = p_dir / "ppt" / "presentation.xml"
    pres_rels = p_dir / "ppt" / "_rels" / "presentation.xml.rels"
    ct_xml = p_dir / "[Content_Types].xml"

    # 1. Determine next slide number
    existing_nums = []
    for f in slides_dir.glob("slide*.xml"):
        name = f.stem
        if name.startswith("slide"):
            try:
                num = int(name.replace("slide", ""))
                existing_nums.append(num)
            except ValueError:
                pass
    next_num = max(existing_nums, default=0) + 1
    new_slide_name = f"slide{next_num}.xml"

    # 2. Copy slide XML and rels
    src_path = slides_dir / source_slide
    dst_path = slides_dir / new_slide_name
    shutil.copy2(src_path, dst_path)

    src_rels = rels_dir / f"{source_slide}.rels"
    if src_rels.exists():
        dst_rels = rels_dir / f"{new_slide_name}.rels"
        shutil.copy2(src_rels, dst_rels)

    # 3. Register in [Content_Types].xml
    ct_tree = ET.parse(ct_xml)
    ct_root = ct_tree.getroot()
    override = ET.Element('{http://schemas.openxmlformats.org/package/2006/content-types}Override', {
        'PartName': f'/ppt/slides/{new_slide_name}',
        'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    })
    ct_root.append(override)
    ct_tree.write(ct_xml, encoding='utf-8', xml_declaration=True)

    # 4. Register in ppt/_rels/presentation.xml.rels
    pr_tree = ET.parse(pres_rels)
    pr_root = pr_tree.getroot()
    existing_rids = [r.attrib.get('Id', '') for r in pr_root]
    rid_nums = [int(r.replace('rId', '')) for r in existing_rids if r.startswith('rId') and r.replace('rId', '').isdigit()]
    next_rid = f"rId{max(rid_nums, default=0) + 1}"

    rel = ET.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', {
        'Id': next_rid,
        'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        'Target': f'slides/{new_slide_name}'
    })
    pr_root.append(rel)
    pr_tree.write(pres_rels, encoding='utf-8', xml_declaration=True)

    # 5. Register in ppt/presentation.xml
    p_tree = ET.parse(pres_xml)
    p_root = p_tree.getroot()
    sld_lst = p_root.find('.//p:sldIdLst', NS)
    if sld_lst is not None:
        existing_ids = [int(s.attrib.get('id', '256')) for s in sld_lst]
        next_id = str(max(existing_ids, default=255) + 1)
        new_sld = ET.Element('{http://schemas.openxmlformats.org/presentationml/2006/main}sldId', {
            'id': next_id,
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': next_rid
        })
        sld_lst.append(new_sld)
        p_tree.write(pres_xml, encoding='utf-8', xml_declaration=True)

    print(f"[OK] Successfully duplicated {source_slide} -> {new_slide_name} (rId: {next_rid})")
    return new_slide_name

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage:")
        print("  python manage_slide.py unpack <deck.pptx> <target_dir>")
        print("  python manage_slide.py repack <source_dir> <out.pptx>")
        print("  python manage_slide.py inspect <unpacked_dir>")
        print("  python manage_slide.py duplicate <unpacked_dir> <slideN.xml>")
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "unpack":
        unpack_pptx(sys.argv[2], sys.argv[3])
    elif cmd == "repack":
        repack_pptx(sys.argv[2], sys.argv[3])
    elif cmd == "inspect":
        inspect_unpacked(sys.argv[2])
    elif cmd == "duplicate":
        duplicate_slide(sys.argv[2], sys.argv[3])
    else:
        print(f"Unknown command: {cmd}")
