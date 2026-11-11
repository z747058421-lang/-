# -*- coding: utf-8 -*-
"""
prepare_template.py — 从参考教案(docx)净化出「空白版式模板」。

做法：保留全部页面设置、表格结构、单元格边框/宽度/对齐、标签列文字、
以及每个区域的"样式样本段落"（供 build_lesson_plan.py 克隆），把可变
内容清空，并删除所有图片与对应媒体引用，使模板仅保留版式骨架。

输出：../assets/教案模板.docx
"""
import os, zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SRC = r"C:\Users\zhanggen\Desktop\五年级上册教案汇总\张根 五上信息科技电子教案-第3课.docx"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'assets', '教案模板.docx')

def q(t): return '{%s}%s' % (W, t)

def clone_el(el):
    return ET.fromstring(ET.tostring(el))

def strip_text(p):
    """删除段落内所有 w:t 的文本，但保留 run 的 rPr（样式样本）。"""
    for r in p.iter(q('r')):
        for t in r.findall(q('t')):
            t.text = ''
        # 清除 tab/br 等占位
        for br in r.findall(q('br')):
            r.remove(br)
    return p

def make_sample_para(bold):
    """构造教学设计区的样式样本段落（pPr + run rPr 已固定）。"""
    p = ET.Element(q('p'))
    pPr = ET.SubElement(p, q('pPr'))
    sp = ET.SubElement(pPr, q('spacing'))
    sp.set(q('line'), '400'); sp.set(q('lineRule'), 'auto')
    r = ET.SubElement(p, q('r'))
    rPr = ET.SubElement(r, q('rPr'))
    rf = ET.SubElement(rPr, q('rFonts'))
    rf.set(q('ascii'), '宋体'); rf.set(q('hAnsi'), '宋体'); rf.set(q('eastAsia'), '宋体')
    if bold:
        ET.SubElement(rPr, q('b'))
    col = ET.SubElement(rPr, q('color')); col.set(q('val'), '000000')
    sz = ET.SubElement(rPr, q('sz')); sz.set(q('val'), '24')
    szc = ET.SubElement(rPr, q('szCs')); szc.set(q('val'), '24')
    t = ET.SubElement(r, q('t'))
    t.text = '__SAMPLE__'
    return p

def main():
    z = zipfile.ZipFile(SRC)
    root = ET.fromstring(z.read('word/document.xml').decode('utf-8'))
    body = root.find(q('body'))

    # 1) 删除所有图片所在的 run
    for p in root.iter(q('p')):
        for r in list(p):
            if r.tag == q('r') and r.find(q('drawing')) is not None:
                p.remove(r)

    # 2) 处理正文顶部三段
    top_ps = list(body.findall(q('p')))
    # P0 标题保留、P1 空段保留；P2 信息行改为单个样式样本段（构建时重建）
    if len(top_ps) >= 3:
        p2 = top_ps[2]
        for child in list(p2):
            p2.remove(child)
        sample_rPr = ET.SubElement(ET.Element(q('r')), q('rPr'))
        rf = ET.SubElement(sample_rPr, q('rFonts'))
        rf.set(q('ascii'), '宋体'); rf.set(q('hAnsi'), '宋体'); rf.set(q('eastAsia'), '宋体')
        ET.SubElement(sample_rPr, q('b'))
        sz = ET.SubElement(sample_rPr, q('sz')); sz.set(q('val'), '24')
        szc = ET.SubElement(sample_rPr, q('szCs')); szc.set(q('val'), '24')
        r = ET.SubElement(p2, q('r'))
        r.append(clone_el(sample_rPr))
        t = ET.SubElement(r, q('t')); t.text = '设计者：'

    # 3) 处理表格
    tbl = body.find(q('tbl'))
    rows = tbl.findall(q('tr'))
    # 标签单元格：(row0based, cell0based) 需要保留文字
    KEEP_LABELS = {
        (0, 0): '班级', (0, 2): '课时计划',
        (1, 0): '课题',
        (2, 0): '教学目标', (3, 0): '重点难点', (4, 0): '教法学法',
        (5, 0): '教学准备', (6, 0): '板书设计',
        (7, 0): '教学设计', (7, 1): '个性化设计',
        (9, 0): '课后反思',
    }
    for ri, row in enumerate(rows):
        cells = row.findall(q('tc'))
        for ci, cell in enumerate(cells):
            if (ri, ci) in KEEP_LABELS:
                continue  # 保留标签/固定文字
            if ri == 8:  # 教学设计内容区
                continue  # 单独处理
            # 清空为样本段
            for para in list(cell.findall(q('p'))):
                cell.remove(para)
            new_p = ET.SubElement(cell, q('p'))
            src_p = None
            # 用同单元格任意已有段落的 pPr 作为样本（此处已被移除，故用通用）
            r = ET.SubElement(new_p, q('r'))
            rPr = ET.SubElement(r, q('rPr'))
            rf = ET.SubElement(rPr, q('rFonts'))
            rf.set(q('ascii'), '宋体'); rf.set(q('hAnsi'), '宋体'); rf.set(q('eastAsia'), '宋体')
            ET.SubElement(rPr, q('b'))
            sz = ET.SubElement(rPr, q('sz')); sz.set(q('val'), '24')
            szc = ET.SubElement(rPr, q('szCs')); szc.set(q('val'), '24')
            t = ET.SubElement(r, q('t')); t.text = ''

        if ri == 8:  # 教学设计：清空后放入两个样式样本段
            cell = cells[0]
            for para in list(cell.findall(q('p'))):
                cell.remove(para)
            cell.append(make_sample_para(bold=True))   # 环节标题样本
            cell.append(make_sample_para(bold=False))  # 正文样本

    # 4) 重新打包，去掉 media，清理 rels
    rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    RNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    for rel in list(rels):
        if 'media/' in (rel.get('Target') or ''):
            rels.remove(rel)

    new_files = {}
    for item in z.infolist():
        name = item.filename
        if name.startswith('word/media/'):
            continue
        if name == 'word/document.xml':
            new_files[name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        elif name == 'word/_rels/document.xml.rels':
            new_files[name] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
        else:
            new_files[name] = z.read(name)

    tmp = OUT + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z2:
        for name, data in new_files.items():
            z2.writestr(name, data)
    os.replace(tmp, os.path.normpath(OUT))
    print('template written ->', os.path.normpath(OUT))

if __name__ == '__main__':
    main()
