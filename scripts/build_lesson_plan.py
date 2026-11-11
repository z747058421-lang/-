# -*- coding: utf-8 -*-
"""
build_lesson_plan.py — 五年级信息科技电子教案生成器（通用版）

不依赖任何特定 AI 平台，纯 Python + 标准库运行。
分页校验需要 pywin32（可选），没有也能生成但不会自动校准页数。

用法:
    python build_lesson_plan.py --content 内容.json --out 输出.docx
    python build_lesson_plan.py --json '{...}' --out 输出.docx
    python build_lesson_plan.py --content 内容.json --out 输出.docx --template 自定义模板.docx
    python build_lesson_plan.py --content 内容.json --out 输出.docx --no-check  # 跳过分页校验

依赖:
    - Python 3.8+（仅标准库即可生成 docx）
    - pywin32（可选，用于 Windows 上通过 Word COM 精确校验页数）
    - python-docx 不需要
"""
import os, sys, json, math, subprocess, zipfile, argparse, re
from xml.etree import ElementTree as ET

# ==================== 配置 ====================
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE_NS = 'http://www.w3.org/XML/1998/namespace'

# 注册命名空间前缀
ET.register_namespace('w', W)
ET.register_namespace('xml', XML_SPACE_NS)

# 脚本所在目录
HERE = os.path.dirname(os.path.abspath(__file__))
# 默认模板路径（同级 assets 目录下）
DEFAULT_TEMPLATE = os.path.join(HERE, '..', 'assets', '教案模板.docx')
# 分页校验脚本路径
CHECK_PAGES_PY = os.path.join(HERE, 'check_pages.py')

# 版式常量
CHAR_PER_LINE = 28          # 教学设计正文每行约字数
LINE_400 = 400              # 固定行距 400 twips (20磅)
PAGE_AVAIL = 25227          # 两页可用高度 (twips)


def q(t):
    """构造 Word 命名空间的限定标签名。"""
    return '{%s}%s' % (W, t)


def xml_space(el):
    """设置 xml:space="preserve" 属性。"""
    el.set('{%s}space' % XML_SPACE_NS, 'preserve')


# ==================== XML 构造 ====================
def new_run(text, bold=False, underline=False, size=24):
    """创建一个 w:r (run) 元素。"""
    r = ET.Element(q('r'))
    rPr = ET.SubElement(r, q('rPr'))
    rf = ET.SubElement(rPr, q('rFonts'))
    rf.set(q('ascii'), '宋体')
    rf.set(q('hAnsi'), '宋体')
    rf.set(q('eastAsia'), '宋体')
    if bold:
        ET.SubElement(rPr, q('b'))
    if underline:
        u = ET.SubElement(rPr, q('u'))
        u.set(q('val'), 'single')
    col = ET.SubElement(rPr, q('color'))
    col.set(q('val'), '000000')
    sz = ET.SubElement(rPr, q('sz'))
    sz.set(q('val'), str(size))
    szc = ET.SubElement(rPr, q('szCs'))
    szc.set(q('val'), str(size))
    t = ET.SubElement(r, q('t'))
    xml_space(t)
    t.text = text
    return r


def new_para(text, bold=False, underline=False, line=LINE_400, line_rule='exact', size=24):
    """创建一个 w:p (段落) 元素。"""
    p = ET.Element(q('p'))
    pPr = ET.SubElement(p, q('pPr'))
    sp = ET.SubElement(pPr, q('spacing'))
    sp.set(q('line'), str(line))
    sp.set(q('lineRule'), line_rule)
    p.append(new_run(text, bold, underline, size))
    return p


def clone_sample(sample_p, text):
    """克隆模板里的样式样本段落，替换文字。"""
    p = ET.fromstring(ET.tostring(sample_p))
    for r in p.iter(q('r')):
        for t in r.findall(q('t')):
            xml_space(t)
            t.text = text
    return p


def fill_cell(cell, paras):
    """清空单元格内所有段落，填入新段落列表。"""
    for p in list(cell.findall(q('p'))):
        cell.remove(p)
    for p in paras:
        cell.append(p)


# ==================== 主构建 ====================
def build(content, template_path):
    """
    读取模板 docx，填充内容，返回修改后的 XML root。

    参数:
        content: dict，教案内容（见下方字段说明）
        template_path: str，模板 docx 路径

    内容 JSON 字段:
        designer        设计者（默认"张 根"）
        week            周次
        teach_date      授课时间
        class_no        班级
        lesson_ordinal  第几课
        lesson_period   课时数
        lesson_name     课题名
        objectives      教学目标（列表）
        key_difficult   重点难点（列表）
        methods         教法学法
        preparation     教学准备
        blackboard      板书设计（列表）
        personalized    个性化设计
        reflection      课后反思
        design          教学设计正文（列表，每项 {"kind":"title|sub|body", "text":"..."} ）
    """
    with zipfile.ZipFile(template_path) as z:
        root = ET.fromstring(z.read('word/document.xml').decode('utf-8'))

    body = root.find(q('body'))

    # --- 顶部信息行 P2 ---
    top_ps = body.findall(q('p'))
    p2 = top_ps[2]
    for c in list(p2):
        p2.remove(c)
    designer = content.get('designer', '张 根')
    week = content.get('week', '')
    teach_date = content.get('teach_date', '')
    p2.append(new_run('设计者：', bold=True))
    p2.append(new_run('    ' + designer + '              ', bold=True, underline=True))
    p2.append(new_run('第   ', bold=True))
    p2.append(new_run(str(week) + '   周        授课时间：', bold=True, underline=True))
    p2.append(new_run(teach_date + '           ', bold=True, underline=True))

    tbl = body.find(q('tbl'))
    rows = tbl.findall(q('tr'))

    def get_cell(r, c):
        return rows[r].findall(q('tc'))[c]

    class_no = content.get('class_no', '五年级     班')
    period = content.get('lesson_period', '1')
    ordinal = content.get('lesson_ordinal', '1')
    lesson_name = content.get('lesson_name', '')

    # R1: 班级 / 课时计划 / 第几课时
    fill_cell(get_cell(0, 1), [new_para(class_no, line=360, line_rule='auto')])
    fill_cell(get_cell(0, 3), [new_para('%s 课时' % period, line=360, line_rule='auto')])
    fill_cell(get_cell(0, 4), [new_para('第 %s 课时' % ordinal, line=360, line_rule='auto')])

    # R2: 课题
    fill_cell(get_cell(1, 1), [new_para('第%s课  %s' % (ordinal, lesson_name), line=360, line_rule='auto')])

    # R3: 教学目标
    fill_cell(get_cell(2, 1), [new_para(x, line=400, line_rule='exact') for x in content.get('objectives', [])])

    # R4: 重点难点
    fill_cell(get_cell(3, 1), [new_para(x, line=400, line_rule='exact') for x in content.get('key_difficult', [])])

    # R5: 教法学法
    fill_cell(get_cell(4, 1), [new_para(content.get('methods', ''), line=400, line_rule='exact')])

    # R6: 教学准备
    fill_cell(get_cell(5, 1), [new_para(content.get('preparation', ''), line=400, line_rule='exact')])

    # R7: 板书设计
    fill_cell(get_cell(6, 1), [new_para(x, line=400, line_rule='exact') for x in content.get('blackboard', [])])

    # R9(0-based=8): 教学设计内容区
    design_samples = rows[8].findall(q('tc'))[0].findall(q('p'))
    title_sample = design_samples[0]
    body_sample = design_samples[1]
    design_cell = rows[8].findall(q('tc'))[0]

    for p in list(design_cell.findall(q('p'))):
        design_cell.remove(p)

    for blk in content.get('design', []):
        kind = blk.get('kind', 'body')
        txt = blk.get('text', '')
        if kind in ('title', 'sub'):
            design_cell.append(clone_sample(title_sample, txt))
        else:
            design_cell.append(clone_sample(body_sample, txt))

    # 个性化设计列
    fill_cell(get_cell(8, 1), [new_para(content.get('personalized', ''), line=360, line_rule='auto')])

    # R10(0-based=9): 课后反思
    fill_cell(get_cell(9, 1), [new_para(content.get('reflection', ''), line=360, line_rule='auto')])

    return root


def set_r10_height(root, h):
    """设置课后反思行（第10行，0-based=9）的高度。"""
    body = root.find(q('body'))
    tbl = body.find(q('tbl'))
    rows = tbl.findall(q('tr'))
    tr = rows[9]
    trPr = tr.find(q('trPr'))
    if trPr is None:
        trPr = ET.SubElement(tr, q('trPr'))
    th = trPr.find(q('trHeight'))
    if th is None:
        th = ET.SubElement(trPr, q('trHeight'))
    th.set(q('val'), str(int(h)))
    th.set(q('hRule'), 'atLeast')


def write_docx(root, dest, template_path):
    """将修改后的 XML 写回 docx（其余部件从模板原样复制）。"""
    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    with zipfile.ZipFile(template_path) as z:
        with zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as z2:
            for item in z.infolist():
                if item.filename == 'word/document.xml':
                    z2.writestr(item, data)
                else:
                    z2.writestr(item, z.read(item.filename))


# ==================== 分页校验 ====================
def find_python_with_pywin32():
    """
    查找有 pywin32 的 Python 解释器。
    优先顺序: 当前 Python > 常见路径 > 失败返回 None。
    """
    # 先试当前 Python
    try:
        import win32com.client  # noqa
        return sys.executable
    except ImportError:
        pass

    # 尝试常见路径（仅 Windows）
    if sys.platform == 'win32':
        candidates = [
            r'D:\tools\python\python.exe',
            r'C:\Python39\python.exe',
            r'C:\Python310\python.exe',
            r'C:\Python311\python.exe',
            r'C:\Python312\python.exe',
            r'C:\Python313\python.exe',
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python', 'Python39', 'python.exe'),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python', 'Python310', 'python.exe'),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python', 'Python311', 'python.exe'),
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Python', 'Python312', 'python.exe'),
        ]
        for p in candidates:
            if os.path.exists(p):
                try:
                    r = subprocess.run(
                        [p, '-c', 'import win32com.client; print("OK")'],
                        capture_output=True, timeout=10)
                    if r.returncode == 0 and b'OK' in r.stdout:
                        return p
                except:
                    pass
    return None


def real_pages(path):
    """
    调用 Word COM 读取真实页数（需要 pywin32 + Windows + 已安装 Word）。
    失败抛 RuntimeError。
    """
    abs_path = os.path.abspath(path)
    python_exe = find_python_with_pywin32()
    if not python_exe:
        raise RuntimeError('未找到 pywin32，无法进行分页校验')

    if not os.path.exists(CHECK_PAGES_PY):
        raise RuntimeError('分页校验脚本不存在: ' + CHECK_PAGES_PY)

    r = subprocess.run(
        [python_exe, CHECK_PAGES_PY, abs_path],
        capture_output=True, timeout=120)
    s = (r.stdout or b'').decode('utf-8', errors='replace').strip()
    if not s:
        s = (r.stdout or b'').decode('gbk', errors='replace').strip()
    if not s or s.startswith('ERR'):
        stderr = (r.stderr or b'').decode('utf-8', errors='replace')
        raise RuntimeError('分页校验失败: ' + (s or stderr[:200]))
    m = re.search(r'\d+', s)
    if not m:
        raise RuntimeError('分页校验失败: 无法解析页数: ' + s)
    return int(m.group())


def estimate_pages(root):
    """
    无 Word COM 时的估算页数（基于内容行数推算）。
    返回估算页数，精度不如真实校验但可做参考。
    """
    body = root.find(q('body'))
    tbl = body.find(q('tbl'))
    rows = tbl.findall(q('tr'))
    design = rows[8].findall(q('tc'))[0]

    n_lines = 0
    for p in design.findall(q('p')):
        txt = ''.join(t.text or '' for t in p.iter(q('t')))
        n_lines += max(1, math.ceil(len(txt) / CHAR_PER_LINE))

    # 前8行固定高度约 4104 twips，表格起点 2689 twips
    h_fixed = 2689 + 4104
    h_design = n_lines * LINE_400
    # 每页版心高度 13958 twips
    page_h = 13958
    total_used = h_fixed + h_design
    # 课后反思行高（当前值）
    tr = rows[9]
    trPr = tr.find(q('trPr'))
    h_reflection = 2076  # 默认
    if trPr is not None:
        th = trPr.find(q('trHeight'))
        if th is not None:
            h_reflection = int(th.get(q('val')) or h_reflection)

    total = total_used + h_reflection
    est_pages = max(1, math.ceil(total / page_h))
    return est_pages


def fit_two_pages(root, out_final, template_path, use_check=True):
    """
    调整课后反思行高，使文档正好 2 页。
    如果 use_check=True 且 pywin32 可用，用 Word COM 精确校验。
    否则用估算模式设置一个合理行高。
    """
    tmp = out_final + '.pagetest.tmp.docx'

    # 估算教学设计行数
    design = root.find(q('body')).find(q('tbl')).findall(q('tr'))[8].findall(q('tc'))[0]
    n_lines = 0
    for p in design.findall(q('p')):
        txt = ''.join(t.text or '' for t in p.iter(q('t')))
        n_lines += max(1, math.ceil(len(txt) / CHAR_PER_LINE))

    h_fixed = 2689 + 4104
    h_design = n_lines * LINE_400
    h_est = PAGE_AVAIL - h_fixed - h_design
    h = max(300, min(20000, h_est))

    # 检查是否能用 Word COM
    can_check = use_check and find_python_with_pywin32() is not None

    if not can_check:
        # 无 Word COM，用估算值设置行高，直接输出
        set_r10_height(root, h)
        write_docx(root, out_final, template_path)
        if os.path.exists(tmp):
            os.remove(tmp)
        est = estimate_pages(root)
        print('[提示] 未检测到 pywin32 或 Word，已用估算行高 %d twips 生成（估算 %d 页）。' % (h, est))
        print('[提示] 如需精确保证两页，请在装有 Word 的 Windows 上安装 pywin32: pip install pywin32')
        return est, h

    # 有 Word COM，精确校验
    lo, hi = 200, 20000
    last_pages = 0
    best_2page_h = None

    for iteration in range(25):
        set_r10_height(root, h)
        write_docx(root, tmp, template_path)
        try:
            pages = real_pages(tmp)
        except RuntimeError:
            # COM 出错，清理后重试一次
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                              capture_output=True, timeout=10)
            except:
                pass
            import time
            time.sleep(2)
            try:
                pages = real_pages(tmp)
            except RuntimeError:
                # COM 仍然失败，退回估算模式
                set_r10_height(root, h)
                write_docx(root, out_final, template_path)
                if os.path.exists(tmp):
                    os.remove(tmp)
                est = estimate_pages(root)
                print('[警告] Word COM 调用失败，已用估算行高 %d twips 生成（估算 %d 页）。' % (h, est))
                return est, h

        last_pages = pages

        if pages == 2:
            best_2page_h = h
            # 微调：尝试增大行高撑满第二页
            for delta in [500, 300]:
                h_try = h + delta
                set_r10_height(root, h_try)
                write_docx(root, tmp, template_path)
                try:
                    p2 = real_pages(tmp)
                except RuntimeError:
                    p2 = 3
                if p2 == 2:
                    best_2page_h = h_try
                    h = h_try
                else:
                    break
            set_r10_height(root, best_2page_h)
            write_docx(root, out_final, template_path)
            if os.path.exists(tmp):
                os.remove(tmp)
            return 2, best_2page_h

        if pages < 2:
            lo = h
            if hi == 20000:
                h = h + 3000
            else:
                h = (lo + hi) // 2
                if h <= lo:
                    h = lo + 200
        else:
            hi = h
            h = (lo + hi) // 2
            if h >= hi:
                h = hi - 200

        if hi - lo <= 50:
            set_r10_height(root, h)
            write_docx(root, out_final, template_path)
            if os.path.exists(tmp):
                os.remove(tmp)
            return last_pages, h

    # 兜底
    if best_2page_h is not None:
        set_r10_height(root, best_2page_h)
        write_docx(root, out_final, template_path)
        if os.path.exists(tmp):
            os.remove(tmp)
        return 2, best_2page_h

    set_r10_height(root, h)
    write_docx(root, out_final, template_path)
    if os.path.exists(tmp):
        os.remove(tmp)
    if last_pages > 2:
        raise ValueError(
            '内容过长，两页排不下（当前 %d 页）。请精简教学设计正文（建议控制在 900 字 / 27 段左右）。' % last_pages)
    if last_pages < 2:
        raise ValueError(
            '内容不足两页（当前 %d 页）。请增加教学设计正文内容（建议 600-980 字 / 17-27 段）。' % last_pages)
    return last_pages, h


# ==================== 入口 ====================
def main():
    ap = argparse.ArgumentParser(
        description='生成五年级信息科技教案 docx（通用版）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python build_lesson_plan.py --content content.json --out 教案.docx
  python build_lesson_plan.py --json '{"lesson_ordinal":"3",...}' --out 教案.docx
  python build_lesson_plan.py --content content.json --out 教案.docx --no-check
  python build_lesson_plan.py --content content.json --out 教案.docx --template 自定义模板.docx
        """)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument('--content', help='内容 JSON 文件路径')
    g.add_argument('--json', help='内联 JSON 内容')
    ap.add_argument('--out', required=True, help='输出 .docx 路径')
    ap.add_argument('--template', default=None,
                    help='模板 docx 路径（默认: ../assets/教案模板.docx）')
    ap.add_argument('--no-check', action='store_true',
                    help='跳过 Word COM 分页校验（用估算行高）')
    args = ap.parse_args()

    # 确定模板路径
    template_path = args.template or DEFAULT_TEMPLATE
    if not os.path.exists(template_path):
        print('错误: 模板文件不存在: ' + template_path)
        sys.exit(1)

    # 读取内容
    if args.content:
        with open(args.content, 'r', encoding='utf-8') as f:
            content = json.load(f)
    else:
        content = json.loads(args.json)

    # 生成
    root = build(content, template_path)
    pages, h = fit_two_pages(root, args.out, template_path, use_check=not args.no_check)
    print('OK 生成: %s | 页数=%d | 课后反思行高=%d twips' % (args.out, pages, h))


if __name__ == '__main__':
    main()
