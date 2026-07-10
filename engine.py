"""
doc-formatter 核心排版引擎
支持三种模式：按指令修改、按规范排版、按标杆文档排版
"""

import re
import os
import shutil
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ─── 字体映射 ───
FONT_MAP = {
    '宋体': 'SimSun', '仿宋': 'FangSong', '黑体': 'SimHei',
    '楷体': 'KaiTi', '微软雅黑': 'Microsoft YaHei', 'Arial': 'Arial',
    'Times New Roman': 'Times New Roman',
    '小标宋': 'SimSun', '小标宋体': 'SimSun',
}

# ─── 字号映射（中文字号 → Pt） ───
SIZE_MAP = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24,
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5,
}


def _parse_size(value):
    """把 '三号' / '15' / '12pt' 统一转为 Pt 数值"""
    value = str(value).strip().lower().replace('pt', '').strip()
    if value in SIZE_MAP:
        return SIZE_MAP[value]
    try:
        return float(value)
    except ValueError:
        return 12  # 默认小四


def _parse_font(name):
    """中文字体名映射回系统字体名，找不到则原样返回"""
    return FONT_MAP.get(name, name)


def _get_pt(emu_value):
    """EMU → Pt"""
    if emu_value is None:
        return None
    return emu_value / 12700


def _set_cn_font(run, font_name):
    """同时设置西文和中文字体"""
    run.font.name = font_name
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_name}"/>')
        rPr.insert(0, rFonts)
    else:
        rFonts.set(qn('w:eastAsia'), font_name)


def _get_run_format(run):
    """安全的获取 run 的格式属性"""
    try:
        font_name = run.font.name
    except:
        font_name = None
    try:
        font_size = run.font.size
    except:
        font_size = None
    try:
        bold = run.font.bold
    except:
        bold = None
    try:
        italic = run.font.italic
    except:
        italic = None
    try:
        color = run.font.color.rgb
    except:
        color = None
    return {
        'name': font_name,
        'size': _get_pt(font_size) if font_size else None,
        'bold': bold,
        'italic': italic,
        'color': str(color) if color else None,
    }


# ═══════════════════════════════════════════════
# 模式一：按指令修改
# ═══════════════════════════════════════════════

def apply_instructions(doc_path, instructions, output_path):
    """
    解析自然语言指令并应用到文档
    支持的指令模式：
      - "把标题改成黑体三号加粗"
      - "把正文改成宋体小四号行距1.5倍"
      - "把页边距改成上3cm下3cm左2.5cm右2.5cm"
      - "把...部分的...改成..."
    """
    doc = Document(doc_path)
    changes = []
    warnings = []

    # 1. 页边距指令 (优先级最高，全局)
    margin_instructions = _extract_margin_instructions(instructions)
    if margin_instructions:
        for section in doc.sections:
            for key, val in margin_instructions.items():
                setattr(section, key, Cm(val))
        changes.append(f"[页边距] 已设置为「{_fmt_margins(margin_instructions)}」")

    # 2. 解析段落级指令
    para_rules = _parse_para_instructions(instructions)
    for rule in para_rules:
        count = 0
        scope = rule.get('scope', '全文')
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            # 范围过滤
            if scope == '封面':
                if para is not doc.paragraphs[0:5]:  # 简单处理前5段为封面
                    continue
            elif scope == '正文':
                if not _is_body_text(para, doc):
                    continue
            # 类型过滤
            elem_type = rule.get('element_type', '全文')
            if elem_type == '标题' and not _is_heading(para):
                continue
            if elem_type == '正文' and _is_heading(para):
                continue

            # 应用格式
            modified = _apply_format_to_para(para, rule)
            if modified:
                count += 1

        if count > 0:
            changes.append(f"[{rule.get('element_type', '全文')}] {scope}（共{count}处）：已改为「{_fmt_rule(rule)}」")

    if not changes:
        changes.append("[提示] 未匹配到可修改的格式项，请检查指令描述是否准确")

    doc.save(output_path)
    return _build_report(doc_path, '按指令修改', changes, warnings, instructions)


def _extract_margin_instructions(text):
    """提取页边距设置，如 '上3cm 下3cm 左2.5cm 右2.5cm'"""
    margins = {}
    pattern = r'(上|下|左|右|top|bottom|left|right)\s*[:：]?\s*(\d+\.?\d*)\s*cm'
    matches = re.findall(pattern, text.lower())
    key_map = {'上': 'top_margin', '下': 'bottom_margin',
               '左': 'left_margin', '右': 'right_margin',
               'top': 'top_margin', 'bottom': 'bottom_margin',
               'left': 'left_margin', 'right': 'right_margin'}
    for direction, val in matches:
        key = key_map.get(direction)
        if key:
            margins[key] = float(val)
    return margins


def _parse_para_instructions(text):
    """解析段落格式指令，可能含多条"""
    rules = []
    # 匹配: "把[范围][类型]改成[字体][字号][加粗/倾斜][行距X倍]"
    patterns = [
        r'(把|将)?(.{0,6})?(标题|正文|全文|封面|落款|一级标题|二级标题)?的?(字体|字号|行距|格式)?改成?'
        r'[\s]*(.{1,8})?[\s]*'
        r'([一二三四五六七八九小初号\d]+号?)?[\s]*'
        r'(加粗|倾斜|加粗倾斜|bold|italic)?[\s]*'
        r'(行距\s*(\d+\.?\d*)\s*倍)?',
    ]

    # 按分号/句号分割多条指令
    for segment in re.split(r'[；;。\n]', text):
        segment = segment.strip()
        if not segment:
            continue

        # 尝试匹配
        for pat in patterns:
            m = re.search(pat, segment)
            if m:
                rule = {}
                scope = m.group(2) if m.group(2) and m.group(2) not in ['把', '将', ''] else '全文'
                scope = scope.replace('的', '').strip()
                elem_type = m.group(3) or '全文'
                font_name = m.group(5)
                font_size = m.group(6)
                bold = '加粗' in (m.group(7) or '') or 'bold' in (m.group(7) or '').lower()
                line_spacing = m.group(9)

                # 特殊处理：如果整个约束是 "字体xxx" 格式
                if not font_name and not font_size:
                    # 尝试直接匹配 "黑体三号" 这种
                    fm = re.search(r'(.{1,8})?\s*([一二三四五六七八九小初号\d]+号?)', segment)
                    if fm:
                        font_name = fm.group(1) or font_name
                        font_size = fm.group(2) or font_size

                if font_name:
                    rule['font'] = _parse_font(font_name)
                if font_size:
                    rule['size'] = _parse_size(font_size)
                if bold:
                    rule['bold'] = True
                if line_spacing:
                    try:
                        rule['line_spacing'] = float(line_spacing)
                    except ValueError:
                        pass
                rule['scope'] = scope
                rule['element_type'] = elem_type
                if rule:
                    rules.append(rule)

                # 也尝试匹配简单的 "行距1.5倍"
                ls_match = re.search(r'行距\s*(\d+\.?\d*)\s*倍', segment)
                if ls_match and not rule.get('line_spacing'):
                    # 尝试补到上一规则
                    if rules:
                        try:
                            rules[-1]['line_spacing'] = float(ls_match.group(1))
                        except ValueError:
                            pass

    return rules


def _is_heading(para):
    """判断段落是否为标题"""
    if para.style.name.startswith('Heading') or para.style.name.startswith('heading'):
        return True
    return False


def _is_body_text(para, doc):
    """粗略判断是否为正文段落（非标题、非空白的段落）"""
    if _is_heading(para):
        return False
    if not para.text.strip():
        return False
    return True


def _apply_format_to_para(para, rule):
    """对段落应用格式规则，返回是否修改"""
    modified = False
    for run in para.runs:
        if 'font' in rule and rule['font']:
            try:
                _set_cn_font(run, rule['font'])
                modified = True
            except:
                pass
        if 'size' in rule and rule['size']:
            try:
                run.font.size = Pt(rule['size'])
                modified = True
            except:
                pass
        if 'bold' in rule and rule['bold']:
            try:
                run.font.bold = True
                modified = True
            except:
                pass
    if 'line_spacing' in rule and rule['line_spacing']:
        try:
            pf = para.paragraph_format
            pf.line_spacing = rule['line_spacing']
            modified = True
        except:
            pass
    return modified


def _fmt_margins(m):
    return f"上{m.get('top','?')}cm 下{m.get('bottom','?')}cm 左{m.get('left','?')}cm 右{m.get('right','?')}cm"


def _fmt_rule(rule):
    parts = []
    if 'font' in rule:
        rev_map = {v: k for k, v in FONT_MAP.items()}
        parts.append(rev_map.get(rule['font'], rule['font']))
    if 'size' in rule:
        rev_size = {v: k for k, v in SIZE_MAP.items()}
        parts.append(rev_size.get(rule['size'], f"{rule['size']}pt"))
    if rule.get('bold'):
        parts.append('加粗')
    if rule.get('line_spacing'):
        parts.append(f"{rule['line_spacing']}倍行距")
    return ' '.join(parts)


# ═══════════════════════════════════════════════
# 模式二：按规范排版
# ═══════════════════════════════════════════════

def _get_style_font_name(style):
    """安全地获取样式的字体名"""
    try:
        return style.font.name
    except:
        return None

def _get_style_font_size(style):
    """安全地获取样式的字号（Pt）"""
    try:
        sz = style.font.size
        return _get_pt(sz) if sz else None
    except:
        return None

def _get_style_bold(style):
    """安全地获取样式的加粗属性"""
    try:
        return style.font.bold
    except:
        return None

def _get_style_line_spacing(style):
    """安全地获取样式的行距"""
    try:
        pf = style.paragraph_format
        if pf and pf.line_spacing:
            return pf.line_spacing
    except:
        pass
    return None

def _get_style_east_asian_font(style):
    """获取样式中文字体名（从 XML 中读取）"""
    try:
        rPr = style.element.find(qn('w:rPr'))
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                ea = rFonts.get(qn('w:eastAsia'))
                if ea:
                    return ea
    except:
        pass
    return None


def _extract_styles_from_docx(docx_path):
    """
    从 .docx 文件的样式库中提取格式规则
    返回: { 'page_margins': {...}, 'heading_一级标题': {...}, 'heading_二级标题': {...}, 'body': {...} }
    """
    spec_doc = Document(docx_path)
    rules = {}

    # 1. 页边距
    for section in spec_doc.sections:
        rules['page_margins'] = {
            'top': round(section.top_margin / 360000, 2),
            'bottom': round(section.bottom_margin / 360000, 2),
            'left': round(section.left_margin / 360000, 2),
            'right': round(section.right_margin / 360000, 2),
        }
        break

    # 2. 标题样式: Heading 1 → 一级标题, Heading 2 → 二级标题, etc.
    style_names = {
        'heading 1': ('heading_一级标题', 1, '一级标题'),
        'heading 2': ('heading_二级标题', 2, '二级标题'),
        'heading 3': ('heading_三级标题', 3, '三级标题'),
    }

    for style in spec_doc.styles:
        sn = style.name.lower() if style.name else ''
        if sn in style_names:
            key, level, label = style_names[sn]
            rule = {}
            font_name = _get_style_font_name(style)
            ea_font = _get_style_east_asian_font(style)
            rule['font'] = ea_font or font_name or 'SimHei'
            sz = _get_style_font_size(style)
            if sz:
                rule['size'] = sz
            bold = _get_style_bold(style)
            if bold is not None:
                rule['bold'] = bold
            ls = _get_style_line_spacing(style)
            if ls:
                rule['line_spacing'] = ls
            rule['level'] = level
            rule['_label'] = label
            rules[key] = rule

    # 3. 正文字体（Normal 样式）
    normal_style = None
    try:
        normal_style = spec_doc.styles['Normal']
    except:
        pass

    if normal_style:
        body_rule = {}
        font_name = _get_style_font_name(normal_style)
        ea_font = _get_style_east_asian_font(normal_style)
        body_rule['font'] = ea_font or font_name or 'SimSun'
        sz = _get_style_font_size(normal_style)
        if sz:
            body_rule['size'] = sz
        bold = _get_style_bold(normal_style)
        if bold is not None:
            body_rule['bold'] = bold
        ls = _get_style_line_spacing(normal_style)
        if ls:
            body_rule['line_spacing'] = ls
        rules['body'] = body_rule

    return rules


def apply_specification(doc_path, spec_path, output_path):
    """
    从规范文件中提取格式规则并应用到文档
    支持两种规范来源：
      - .docx 文件：直接从样式库读取 Heading 1/2/3 + Normal 样式定义
      - .md / .pdf / 文本：从文字描述中正则提取格式规则
    """
    doc = Document(doc_path)
    changes = []
    warnings = []
    spec_rules = {}

    # 判断规范文件类型
    ext = os.path.splitext(spec_path)[1].lower()

    if ext == '.docx':
        spec_rules = _extract_styles_from_docx(spec_path)
        if spec_rules:
            changes.append(f"[来源] 已从 .docx 样式库读取格式定义")
    else:
        spec_rules = _parse_specification(spec_path)

    if not spec_rules:
        warnings.append("未能从规范文件中提取出格式规则")

    # 1. 全局设置 (页边距)
    if 'page_margins' in spec_rules:
        m = spec_rules['page_margins']
        for section in doc.sections:
            if 'top' in m:
                section.top_margin = Cm(m['top'])
            if 'bottom' in m:
                section.bottom_margin = Cm(m['bottom'])
            if 'left' in m:
                section.left_margin = Cm(m['left'])
            if 'right' in m:
                section.right_margin = Cm(m['right'])
        changes.append(f"[页边距] 已设置为「{_fmt_margins(m)}」")

    # 2. 标题格式
    heading_rules = {k: v for k, v in spec_rules.items() if k.startswith('heading')}
    heading_count = {k: 0 for k in heading_rules}
    for para in doc.paragraphs:
        for hk, hr in heading_rules.items():
            if para.style.name.startswith('Heading') or para.style.name.startswith('heading'):
                level = hr.get('level', 1)
                style_level = 1
                try:
                    style_level = int(para.style.name.replace('Heading ', '').replace('heading ', ''))
                except:
                    pass
                if style_level == level:
                    _apply_format_to_para(para, hr)
                    heading_count[hk] += 1

    for hk, cnt in heading_count.items():
        if cnt > 0:
            hr = heading_rules[hk]
            label = hr.get('_label', hk)
            changes.append(f"[标题] {label}（共{cnt}处）：已改为「{_fmt_rule(hr)}」")

    # 3. 正文格式
    if 'body' in spec_rules:
        br = spec_rules['body']
        count = 0
        for para in doc.paragraphs:
            if not _is_heading(para) and para.text.strip():
                _apply_format_to_para(para, br)
                count += 1
        if count > 0:
            changes.append(f"[正文] 全文（共{count}处）：已改为「{_fmt_rule(br)}」")

    if not changes:
        changes.append("[提示] 规范规则已读取，但文档未发现可修改元素")

    doc.save(output_path)
    return _build_report(doc_path, '按规范排版', changes, warnings, os.path.basename(spec_path))


def _parse_specification(spec_path):
    """
    从规范文件（Markdown / 文本）中提取格式规则
    支持格式：
      - 标准格式：页边距：上2.54cm 下2.54cm 左3.17cm 右3.17cm
      - 标准格式：一级标题：黑体 小三号 加粗
      - 国家标准用语：标题 一般用2号小标宋体字
      - 国家标准用语：正文 一般用3号仿宋体字
      - 国家标准用语：天头（上白边）为37 mm，订口（左白边）为28mm
      - 国家标准用语：公文格式各要素一般用3号仿宋体字
    """
    rules = {}

    with open(spec_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # ── 1. 页边距 ──
    margins = {}

    # 格式A: "页边距：上37mm 下35mm 左28mm 右26mm"
    margin_pat = r'页边距[：:。，,\s]*(?:上(\d+\.?\d*)\s*mm?)?[^。\n]*?(?:下(\d+\.?\d*)\s*mm?)?[^。\n]*?(?:左(\d+\.?\d*)\s*mm?)?[^。\n]*?(?:右(\d+\.?\d*)\s*mm?)?'
    m = re.search(margin_pat, content)
    if m:
        if m.group(1): margins['top'] = float(m.group(1)) / 10
        if m.group(2): margins['bottom'] = float(m.group(2)) / 10
        if m.group(3): margins['left'] = float(m.group(3)) / 10
        if m.group(4): margins['right'] = float(m.group(4)) / 10

    # 格式B: "天头（上白边）为37 mm±1 mm" "订口（左白边）为28mm±1mm"
    if not margins:
        # 上/天头
        m_top = re.search(r'(?:天头|上白边|上[^，。]*?边[距]?)[^。]*?为\s*(\d+\.?\d*)\s*mm', content)
        if m_top: margins['top'] = float(m_top.group(1)) / 10

        # 下
        m_bottom = re.search(r'(?:下白边|下[^，。]*?边[距]?)[^。]*?为\s*(\d+\.?\d*)\s*mm', content)
        if m_bottom: margins['bottom'] = float(m_bottom.group(1)) / 10

        # 左/订口
        m_left = re.search(r'(?:订口|左白边|左[^，。]*?边[距]?)[^。]*?为\s*(\d+\.?\d*)\s*mm', content)
        if m_left: margins['left'] = float(m_left.group(1)) / 10

        # 右
        m_right = re.search(r'(?:右白边|右[^，。]*?边[距]?)[^。]*?为\s*(\d+\.?\d*)\s*mm', content)
        if m_right: margins['right'] = float(m_right.group(1)) / 10

        # 如果只找到上左，通过版心推算下右（适用于国标场景）
        # 国标：A4=210mm×297mm，上37mm左28mm，版心156mm×225mm
        if 'top' in margins and 'left' in margins:
            margins['right'] = round(21.0 - margins['left'] - 15.6, 1)  # 210mm - 左白边 - 156mm版心宽
            margins['bottom'] = round(29.7 - margins['top'] - 22.5, 1)  # 297mm - 上白边 - 225mm版心高

    if margins:
        rules['page_margins'] = margins

    # ── 2. 标题规则 ──
    # 格式A: "一级标题[：:] 黑体 小三号 加粗"
    for level in ['一级标题', '二级标题', '三级标题', '四级标题']:
        pat = rf'{re.escape(level)}[：:。，\s]+(.+?)(?=\n|$)'
        m = re.search(pat, content)
        if m:
            rule_text = m.group(1)
            rule = _extract_format_from_text(rule_text)
            if rule:
                rule['level'] = ['一级标题', '二级标题', '三级标题', '四级标题'].index(level) + 1
                rules[f'heading_{level}'] = rule

    # 格式B: 国家标准用语 "标题 一般用2号小标宋体字" (注意提取的PDF中可能有空格: "2 号")
    if 'heading_一级标题' not in rules:
        title_match = re.search(r'(?:公文)?标题[：:。，\s]+(?:一般|通常)?用\s*(\S+?)\s*字号?\s*(\S+?)体字', content)
        if not title_match:
            title_match = re.search(r'(?:公文)?标题[：:。，\s]+\S*?用\S*?(\d+)\s*号\S*?(\S+?)体字', content)
        if not title_match:
            title_match = re.search(r'标题[^。]{0,40}?(\d+)\s*号\S*?(\S{2,4})体', content)
        if title_match:
            level_num = 1
            size_str = title_match.group(1).strip()
            font_str = title_match.group(2).strip()
            rule = {'level': level_num}
            # 字号映射
            size_map = {'一': '一号', '二': '二号', '三': '三号', '四': '四号',
                       '小一': '小一', '小二': '小二', '小三': '小三', '小四': '小四'}
            if size_str in size_map:
                rule['size'] = _parse_size(size_map[size_str])
            elif size_str.isdigit():
                size_lookup = {'1': 26, '2': 22, '3': 16, '4': 14}
                if size_str in size_lookup:
                    rule['size'] = size_lookup[size_str]
            # 字体
            font_map = {'小标宋': 'SimSun', '宋体': 'SimSun', '仿宋': 'FangSong',
                       '黑体': 'SimHei', '楷体': 'KaiTi'}
            full_font = font_str
            for k, v in font_map.items():
                if k in font_str:
                    full_font = v
                    break
            rule['font'] = full_font
            rules['heading_一级标题'] = rule

    # ── 3. 正文规则 ──
    # 格式A: "正文[：:] 宋体 小四号 1.5倍行距 首行缩进2字符"
    body_match = re.search(r'正文[：:。，\s]+(.+?)(?=\n|$)', content)
    if body_match:
        rule = _extract_format_from_text(body_match.group(1))
        if rule:
            rules['body'] = rule

    # 格式B: "公文格式各要素一般用3号仿宋体字" (国家标准，可能有空格如 "3 号")
    if 'body' not in rules:
        for pat in [
            r'公文格式各要素[^。]*?用\s*(\d+)\s*号\s*(\S+?)体字',
            r'正文[^。]*?用\s*(\d+)\s*号\s*(\S+?)体字',
            r'一般用\s*(\d+)\s*号\s*(\S+?)体字',
        ]:
            m = re.search(pat, content)
            if m:
                body_rule = {}
                size_str = m.group(1).strip()
                font_str = m.group(2).strip()
                size_lookup = {'1': 26, '2': 22, '3': 16, '4': 14, '5': 10.5}
                if size_str in size_lookup:
                    body_rule['size'] = size_lookup[size_str]
                font_map = {'仿宋': 'FangSong', '宋体': 'SimSun', '黑体': 'SimHei', '楷体': 'KaiTi'}
                for k, v in font_map.items():
                    if k in font_str:
                        body_rule['font'] = v
                        break
                if body_rule:
                    rules['body'] = body_rule
                    break

    # ── 4. 增强：从国标正文段落提取颜色规则 ──
    color_match = re.search(r'文字的颜色[：:。，\s]+([^。]{0,50}?黑[色]?)', content)
    if color_match:
        pass  # 颜色不直接用于 python-docx 段落设置

    # ── 5. 行数字数密度（仅记录，不直接应用） ──
    lines_per_page = re.search(r'每面排\s*(\d+)\s*行', content)
    chars_per_line = re.search(r'每行排\s*(\d+)\s*个字', content)
    if lines_per_page or chars_per_line:
        density = {}
        if lines_per_page: density['lines'] = int(lines_per_page.group(1))
        if chars_per_line: density['chars_per_line'] = int(chars_per_line.group(1))
        # 不直接应用，仅记录参考

    return rules


def _extract_format_from_text(text):
    """从文本片段中提取字体/字号/加粗信息"""
    rule = {}

    # 字体名
    for cn_name in ['宋体', '仿宋', '黑体', '楷体', '微软雅黑']:
        if cn_name in text:
            rule['font'] = _parse_font(cn_name)
            break

    # 字号
    for cn_size in ['小三', '小四', '小五', '小六', '小一', '小二', '小三', '小四',
                    '初号', '小初', '一号', '二号', '三号', '四号', '五号', '六号']:
        if cn_size in text:
            rule['size'] = _parse_size(cn_size)
            break

    # 加粗
    if '加粗' in text or 'bold' in text.lower():
        rule['bold'] = True

    # 行距
    ls_match = re.search(r'(\d+\.?\d*)\s*倍行距', text)
    if ls_match:
        try:
            rule['line_spacing'] = float(ls_match.group(1))
        except ValueError:
            pass

    return rule


# ═══════════════════════════════════════════════
# 模式三：按标杆文档排版
# ═══════════════════════════════════════════════

def apply_benchmark(doc_path, benchmark_path, output_path):
    """
    从标杆文档中提取格式特征，应用到目标文档
    对于 .docx 标杆：优先从样式库读取 Heading/Normal 样式定义
    对于其他格式：从文档段落内容中提取格式特征
    """
    doc = Document(doc_path)
    changes = []
    warnings = []

    # 判断文档类型
    ext = os.path.splitext(benchmark_path)[1].lower()

    if ext == '.docx':
        # 从样式库读取
        spec_rules = _extract_styles_from_docx(benchmark_path)

        # 1. 页边距
        if 'page_margins' in spec_rules:
            m = spec_rules['page_margins']
            for section in doc.sections:
                if 'top' in m: section.top_margin = Cm(m['top'])
                if 'bottom' in m: section.bottom_margin = Cm(m['bottom'])
                if 'left' in m: section.left_margin = Cm(m['left'])
                if 'right' in m: section.right_margin = Cm(m['right'])
            changes.append(f"[来源] 已从标杆 .docx 样式库读取格式定义")

        # 2. 标题格式
        heading_rules = {k: v for k, v in spec_rules.items() if k.startswith('heading')}
        heading_count = {k: 0 for k in heading_rules}
        for para in doc.paragraphs:
            for hk, hr in heading_rules.items():
                if para.style.name.startswith('Heading') or para.style.name.startswith('heading'):
                    level = hr.get('level', 1)
                    style_level = 1
                    try:
                        style_level = int(para.style.name.replace('Heading ', '').replace('heading ', ''))
                    except:
                        pass
                    if style_level == level:
                        _apply_format_to_para(para, hr)
                        heading_count[hk] += 1
        for hk, cnt in heading_count.items():
            if cnt > 0:
                hr = heading_rules[hk]
                label = hr.get('_label', hk)
                changes.append(f"[标题] {label}（共{cnt}处）：已按标杆样式库调整")

        # 3. 正文格式
        if 'body' in spec_rules:
            br = spec_rules['body']
            count = 0
            for para in doc.paragraphs:
                if not _is_heading(para) and para.text.strip():
                    _apply_format_to_para(para, br)
                    count += 1
            if count > 0:
                changes.append(f"[正文] 全文（共{count}处）：已按标杆样式库调整")

        if not changes:
            warnings.append("未从标杆样式库发现可用的格式定义")
    else:
        # 非 .docx：走原有段落特征提取
        benchmark = Document(benchmark_path)
        features = _extract_benchmark_features(benchmark)

        if not features:
            warnings.append("无法从标杆文档中提取足够的格式特征")

        # 1. 页边距
        bm_margins = features.get('margins', {})
        if bm_margins:
            for section in doc.sections:
                if 'top' in bm_margins:
                    section.top_margin = Emu(bm_margins['top'])
                if 'bottom' in bm_margins:
                    section.bottom_margin = Emu(bm_margins['bottom'])
                if 'left' in bm_margins:
                    section.left_margin = Emu(bm_margins['left'])
                if 'right' in bm_margins:
                    section.right_margin = Emu(bm_margins['right'])
            _m = {k: round(v/360000, 2) for k, v in bm_margins.items()}
            changes.append(f"[页边距] 已按标杆文档设置为「上{_m.get('top','?')}cm 下{_m.get('bottom','?')}cm 左{_m.get('left','?')}cm 右{_m.get('right','?')}cm」")

        # 2. 标题格式
        heading_patterns = features.get('headings', {})
        heading_count = 0
        for para in doc.paragraphs:
            if _is_heading(para):
                try:
                    level = int(para.style.name.replace('Heading ', '').replace('heading ', ''))
                except:
                    level = 1
                pattern = heading_patterns.get(level)
                if pattern:
                    for run in para.runs:
                        if pattern.get('name'):
                            try:
                                _set_cn_font(run, pattern['name'])
                            except:
                                pass
                        if pattern.get('size'):
                            try:
                                run.font.size = Pt(pattern['size'])
                            except:
                                pass
                        if pattern.get('bold') is not None:
                            try:
                                run.font.bold = pattern['bold']
                            except:
                                pass
                    heading_count += 1

        if heading_count > 0:
            changes.append(f"[标题] 全文（共{heading_count}处）：已按标杆文档格式调整")

        # 3. 正文格式
        body_pattern = features.get('body', {})
        if body_pattern:
            body_count = 0
            for para in doc.paragraphs:
                if not _is_heading(para) and para.text.strip():
                    for run in para.runs:
                        if body_pattern.get('name'):
                            try:
                                _set_cn_font(run, body_pattern['name'])
                            except:
                                pass
                        if body_pattern.get('size'):
                            try:
                                run.font.size = Pt(body_pattern['size'])
                            except:
                                pass
                    # 行距
                    if body_pattern.get('line_spacing'):
                        try:
                            para.paragraph_format.line_spacing = body_pattern['line_spacing']
                        except:
                            pass
                    body_count += 1
            if body_count > 0:
                changes.append(f"[正文] 全文（共{body_count}处）：已按标杆文档格式调整")

        if not changes:
            warnings.append("未发现需要修改的格式项（目标文档格式与标杆文档可能已一致）")

    doc.save(output_path)
    return _build_report(doc_path, '按标杆文档排版', changes, warnings, os.path.basename(benchmark_path))


def _extract_benchmark_features(benchmark_doc):
    """从标杆文档中提取格式特征"""
    features = {'margins': {}, 'headings': {}, 'body': {}}

    # 页边距
    for section in benchmark_doc.sections:
        features['margins']['top'] = section.top_margin
        features['margins']['bottom'] = section.bottom_margin
        features['margins']['left'] = section.left_margin
        features['margins']['right'] = section.right_margin
        break  # 只取第一个section

    # 标题 & 正文格式
    heading_formats = {}
    body_runs = []

    for para in benchmark_doc.paragraphs:
        if not para.runs:
            continue

        if _is_heading(para):
            try:
                level = int(para.style.name.replace('Heading ', '').replace('heading ', ''))
            except:
                level = 1
            if level not in heading_formats:
                fmt = _get_run_format(para.runs[0])
                if fmt['name'] or fmt['size']:
                    heading_formats[level] = fmt
        else:
            if para.text.strip():
                fmt = _get_run_format(para.runs[0])
                if fmt['name'] or fmt['size']:
                    body_runs.append(fmt)
                # 行距
                try:
                    ls = para.paragraph_format.line_spacing
                    if ls:
                        features['body']['line_spacing'] = ls
                except:
                    pass

    features['headings'] = heading_formats

    # 正文：取多数格式
    if body_runs:
        from collections import Counter
        names = [r['name'] for r in body_runs if r['name']]
        sizes = [r['size'] for r in body_runs if r['size']]
        if names:
            features['body']['name'] = Counter(names).most_common(1)[0][0]
        if sizes:
            features['body']['size'] = Counter(sizes).most_common(1)[0][0]

    return features


# ═══════════════════════════════════════════════
# 报告生成
# ═══════════════════════════════════════════════

def _build_report(doc_path, mode, changes, warnings, source):
    """生成排版变更报告"""
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    return {
        'success': True,
        'report': {
            'filename': os.path.basename(doc_path),
            'mode': mode,
            'source': source,
            'time': now,
            'changes': changes,
            'warnings': warnings,
        }
    }


# ═══════════════════════════════════════════════
# 知识库管理
# ═══════════════════════════════════════════════

KNOWLEDGE_BASE_PATH = os.path.join(os.path.dirname(__file__), 'knowledge_base')


def save_to_knowledge_base(uploaded_file_path, file_type='specification'):
    """
    将上传的文件保存到知识库
    file_type: 'specification' → 01-格式规范/, 'benchmark' → 02-标杆文档/
    """
    if not os.path.exists(uploaded_file_path):
        return None

    today = datetime.now().strftime('%Y%m%d')
    filename = os.path.basename(uploaded_file_path)

    if file_type == 'specification':
        target_dir = os.path.join(KNOWLEDGE_BASE_PATH, '01-格式规范')
        new_name = f'用户上传_{today}_{filename}'
    elif file_type == 'benchmark':
        target_dir = os.path.join(KNOWLEDGE_BASE_PATH, '02-标杆文档')
        new_name = f'标杆文档_{today}_{filename}'
    else:
        return None

    os.makedirs(target_dir, exist_ok=True)
    dest_path = os.path.join(target_dir, new_name)

    # 检查重复
    if os.path.exists(dest_path):
        base, ext = os.path.splitext(new_name)
        counter = 1
        while os.path.exists(dest_path):
            dest_path = os.path.join(target_dir, f'{base}_{counter}{ext}')
            counter += 1

    shutil.copy2(uploaded_file_path, dest_path)
    return dest_path
