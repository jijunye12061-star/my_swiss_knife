"""
Excel 水印 + 工作表保护
- 背景水印（通过sheet背景图片实现）
- 审阅保护（防止编辑）
"""

import zipfile
import re
import shutil
import io
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook
from openpyxl.worksheet.protection import SheetProtection

# 注册命名空间
NAMESPACES = {
    '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
    'xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    'xr6': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision6',
    'xr10': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision10',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'


def create_watermark_image(text="内部资料 禁止外传", width=600, height=400,
                           font_size=60, color=(180, 180, 180, 240), angle=30):
    """生成斜体文字水印PNG"""
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    font = None
    for fp in [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
    ]:
        try:
            font = ImageFont.truetype(fp, font_size)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((width - tw) // 2, (height - th) // 2), text, font=font, fill=color)
    img = img.rotate(angle, expand=False)

    bg = Image.new("RGB", img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3])
    return bg


def add_background_watermark(input_xlsx, output_xlsx, watermark_text="内部资料 禁止外传"):
    """将水印图片设置为Excel Sheet背景"""
    # 生成水印
    wm_img = create_watermark_image(text=watermark_text)
    buf = io.BytesIO()
    wm_img.save(buf, format="PNG")
    img_data = buf.getvalue()

    shutil.copy2(input_xlsx, output_xlsx)

    # 读取所有文件
    with zipfile.ZipFile(output_xlsx, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    # 1. 注册png类型
    ct_root = ET.fromstring(files['[Content_Types].xml'])
    ET.register_namespace('', CT_NS)
    has_png = any(e.get('Extension', '').lower() == 'png' for e in ct_root)
    if not has_png:
        el = ET.SubElement(ct_root, f'{{{CT_NS}}}Default')
        el.set('Extension', 'png')
        el.set('ContentType', 'image/png')
    files['[Content_Types].xml'] = ET.tostring(ct_root, encoding='UTF-8', xml_declaration=True)

    # 2. 添加水印图片
    files['xl/media/watermark.png'] = img_data

    # 3. 处理每个sheet
    wm_rel_id = "rIdWatermark"
    sheet_paths = [n for n in files if re.match(r'xl/worksheets/sheet\d+\.xml', n)]

    for sheet_xml_path in sheet_paths:
        m = re.search(r'sheet(\d+)\.xml', sheet_xml_path)
        if not m:
            continue
        sheet_num = m.group(1)
        rel_path = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"

        # 更新/创建 rels
        if rel_path in files:
            rel_root = ET.fromstring(files[rel_path])
        else:
            ET.register_namespace('', PKG_REL_NS)
            rel_root = ET.fromstring(
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{PKG_REL_NS}"/>'
            )

        has_wm = any('watermark.png' in r.get('Target', '') for r in rel_root)
        if not has_wm:
            new_rel = ET.SubElement(rel_root, f'{{{PKG_REL_NS}}}Relationship')
            new_rel.set('Id', wm_rel_id)
            new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            new_rel.set('Target', '../media/watermark.png')

        ET.register_namespace('', PKG_REL_NS)
        files[rel_path] = ET.tostring(rel_root, encoding='UTF-8', xml_declaration=True)

        # 在sheet XML中插入 <picture>
        sheet_text = files[sheet_xml_path]
        if isinstance(sheet_text, bytes):
            sheet_text = sheet_text.decode('utf-8')

        if '<picture ' in sheet_text or '<picture>' in sheet_text:
            continue

        # 必须检查根标签内是否有xmlns:r，而非整个文档
        root_tag_end = sheet_text.index('>')
        root_tag = sheet_text[:root_tag_end]
        if 'xmlns:r=' not in root_tag:
            sheet_text = sheet_text.replace(
                f'xmlns="{SHEET_NS}"',
                f'xmlns="{SHEET_NS}" xmlns:r="{REL_NS}"'
            )

        picture_tag = f'<picture r:id="{wm_rel_id}"/>'
        if '<extLst' in sheet_text:
            sheet_text = sheet_text.replace('<extLst', picture_tag + '<extLst')
        elif '</worksheet>' in sheet_text:
            sheet_text = sheet_text.replace('</worksheet>', picture_tag + '</worksheet>')
        else:
            continue

        files[sheet_xml_path] = sheet_text.encode('utf-8')

    # 4. 重新打包
    with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f'✅ 背景水印已添加: {output_xlsx}')


def protect_all_sheets(xlsx_path, password=None):
    """保护所有工作表（审阅-保护工作表）"""
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        ws.protection = SheetProtection(
            sheet=True,
            password=password,
            formatCells=False,
            formatColumns=False,
            formatRows=False,
            sort=False,
            autoFilter=False,
        )
    wb.save(xlsx_path)
    print(f'✅ 工作表保护已启用: {xlsx_path}')


def apply_watermark_and_protection(input_xlsx, output_xlsx,
                                    watermark_text="内部资料 禁止外传",
                                    password=None):
    """一步完成保护 + 水印（先保护再加水印，避免XML解析问题）"""
    protect_all_sheets(input_xlsx, password)
    add_background_watermark(input_xlsx, output_xlsx, watermark_text)
    return output_xlsx
