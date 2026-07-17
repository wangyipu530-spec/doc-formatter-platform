"""
LLM 适配器 — 负责与 DeepSeek API 通信
功能：
  1. 收到 engine.py 提取的格式快照 + 规范要求
  2. 构造 prompt 发给 LLM
  3. 解析 LLM 返回的 JSON 指令
  4. 返回结构化指令供 engine.py 执行
"""
import json
import os
import yaml
import urllib.request
import urllib.error

# ── 加载配置 ──
_config = None


def _load_config():
    global _config
    if _config is not None:
        return _config
    cfg_path = os.path.join(os.path.dirname(__file__), 'config.yaml')
    with open(cfg_path, 'r', encoding='utf-8') as f:
        _config = yaml.safe_load(f)
    # 环境变量覆盖 API Key（部署用）
    env_key = os.environ.get('DEEPSEEK_API_KEY') or os.environ.get('DEEPSEEK_KEY') or os.environ.get('DEEPSEEK_APIKEY')
    if env_key:
        _config.setdefault('llm', {})['api_key'] = env_key
    return _config


def _call_llm(system_prompt, user_prompt):
    """调用 DeepSeek API，返回原始文本"""
    cfg = _load_config()
    llm_cfg = cfg['llm']

    payload = json.dumps({
        "model": llm_cfg['model'],
        "temperature": llm_cfg['temperature'],
        "max_tokens": llm_cfg['max_tokens'],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
    }).encode('utf-8')

    req = urllib.request.Request(
        f"{llm_cfg['base_url']}/v1/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {llm_cfg['api_key']}"
        }
    )

    try:
        resp = urllib.request.urlopen(req, timeout=60)
        result = json.loads(resp.read().decode('utf-8'))
        return result['choices'][0]['message']['content']
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        raise RuntimeError(f"LLM API 调用失败 (HTTP {e.code}): {body}")
    except Exception as e:
        raise RuntimeError(f"LLM API 调用异常: {str(e)}")


def _extract_page_margins(doc):
    """从 Document 对象提取页边距"""
    for section in doc.sections:
        return {
            'top_cm': round(section.top_margin / 360000, 2),
            'bottom_cm': round(section.bottom_margin / 360000, 2),
            'left_cm': round(section.left_margin / 360000, 2),
            'right_cm': round(section.right_margin / 360000, 2),
        }
    return {}


def _safe_font_name(run):
    try:
        return run.font.name or ''
    except:
        return ''


def _safe_font_size(run):
    try:
        sz = run.font.size
        return round(sz / 12700, 1) if sz else None
    except:
        return None


def _safe_bold(run):
    try:
        return run.font.bold
    except:
        return None


def _safe_italic(run):
    try:
        return run.font.italic
    except:
        return None


def _safe_color(run):
    try:
        rgb = run.font.color.rgb
        return str(rgb) if rgb else None
    except:
        return None


def _get_para_alignment(para):
    try:
        align = para.alignment
        if align is None:
            return 'LEFT'
        names = {
            0: 'LEFT', 1: 'CENTER', 2: 'RIGHT',
            3: 'BOTH', 4: 'MEDIUM', 5: 'DISTRIBUTE',
            7: 'RIGHT_NUMBER'
        }
        return names.get(align, 'LEFT')
    except:
        return 'LEFT'


def _get_line_spacing(para):
    try:
        ls = para.paragraph_format.line_spacing
        if ls:
            return round(ls, 2)
    except:
        pass
    return None


def _get_first_line_indent(para):
    try:
        indent = para.paragraph_format.first_line_indent
        if indent:
            return round(indent / 12700, 1)
    except:
        pass
    return None


def _is_heading(para):
    name = para.style.name.lower() if para.style.name else ''
    return name.startswith('heading')


def _get_heading_level(para):
    if _is_heading(para):
        try:
            return int(para.style.name.replace('Heading ', '').replace('heading ', ''))
        except:
            return 1
    return 0


def extract_document_snapshot(doc):
    """
    从 python-docx Document 提取完整的格式快照，供 LLM 分析。

    返回:
    {
        'margins': {...},
        'paragraphs': [
            {
                'index': 0,
                'text_preview': '标题内容...',
                'type': 'heading',       # heading / body / empty
                'heading_level': 1,       # 0 表示非标题
                'style_name': 'Heading 1',
                'runs': [
                    {'font': 'SimHei', 'size': 16, 'bold': True, ...}
                ],
                'alignment': 'LEFT',
                'line_spacing': 1.5,
                'first_line_indent_cm': 0.74,
            },
            ...
        ],
        'paragraph_count': 10,
    }
    """
    snapshot = {
        'margins': _extract_page_margins(doc),
        'paragraphs': [],
        'paragraph_count': len(doc.paragraphs),
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        para_info = {
            'index': i,
            'text_preview': text[:100] if text else '',
            'type': 'heading' if _is_heading(para) else ('empty' if not text else 'body'),
            'heading_level': _get_heading_level(para),
            'style_name': para.style.name or '',
            'runs': [],
            'alignment': _get_para_alignment(para),
            'line_spacing': _get_line_spacing(para),
            'first_line_indent_cm': _get_first_line_indent(para),
        }

        for run in para.runs:
            run_info = {
                'font': _safe_font_name(run),
                'size_pt': _safe_font_size(run),
                'bold': _safe_bold(run),
                'italic': _safe_italic(run),
                'color': _safe_color(run),
                'text': run.text[:50] if run.text else '',
            }
            para_info['runs'].append(run_info)

        snapshot['paragraphs'].append(para_info)

    return snapshot


def _json_to_instructions(llm_text):
    """从 LLM 返回文本中提取 JSON 指令块"""
    # 尝试直接解析
    text = llm_text.strip()

    # 找 ```json ... ``` 块
    if '```json' in text:
        text = text.split('```json')[1].split('```')[0].strip()
    elif '```' in text:
        text = text.split('```')[1].split('```')[0].strip()

    # 找 { } 包裹的最外层 JSON
    start = text.find('{')
    end = text.rfind('}')
    if start >= 0 and end > start:
        text = text[start:end+1]

    return json.loads(text)


def analyze_and_plan(mode, target_snapshot, spec_text, user_instructions=''):
    """
    核心函数：将格式快照 + 规范要求发给 LLM，返回结构化指令。

    参数:
        mode: 'instruction' / 'specification' / 'benchmark'
        target_snapshot: extract_document_snapshot() 的返回值
        spec_text: 规范/标杆的文本内容（或为空）
        user_instructions: 用户的文字指令

    返回:
        {
            'analysis': str,           # 差异分析文字
            'issues': [str],           # 发现的问题列表
            'instructions': [...],     # 引擎可执行的指令数组
            'suggestions': [str],      # 需人工复核的建议
        }
    """
    mode_names = {
        'instruction': '按指令修改',
        'specification': '按规范排版',
        'benchmark': '按标杆文档排版',
    }
    mode_name = mode_names.get(mode, mode)

    # 构造快照文本
    para_lines = []
    for p in target_snapshot['paragraphs']:
        runs_detail = '; '.join([
            f"字体={r['font'] or '?原文?'}, 字号={r['size_pt'] or '?'}pt, 加粗={r['bold']}, "
            f"斜体={r['italic']}, 颜色={r['color'] or '?'}"
            for r in p['runs']
        ]) if p['runs'] else '(无run)'

        para_lines.append(
            f"  [{p['index']}] 类型={p['type']} 样式={p['style_name']} "
            f"层级={p['heading_level']} 对齐={p['alignment']} "
            f"行距={p['line_spacing']} 首行缩进={p['first_line_indent_cm']}cm\n"
            f"     文字预览: {p['text_preview'][:60]}\n"
            f"     runs: {runs_detail[:120]}"
        )

    margins = target_snapshot['margins']
    margins_text = f"上{margins.get('top_cm','?')}cm 下{margins.get('bottom_cm','?')}cm " \
                   f"左{margins.get('left_cm','?')}cm 右{margins.get('right_cm','?')}cm"

    doc_snapshot_text = f"""
===== 目标文档格式快照 =====
总段落数: {target_snapshot['paragraph_count']}
页边距: {margins_text}

各段落格式:
"""
    doc_snapshot_text += '\n'.join(para_lines)

    # 构造规范/指令描述
    spec_text_block = f"""
===== 格式要求 =====
{mode_name} 模式

"""
    if mode == 'instruction':
        spec_text_block += f"用户指令：\n{user_instructions}\n"
    elif mode == 'specification':
        spec_text_block += f"格式规范内容：\n{spec_text}\n"
    elif mode == 'benchmark':
        spec_text_block += f"标杆文档内容/特征：\n{spec_text}\n"

    system_prompt = """你是一个专业的 Word 文档排版分析专家。你的任务是分析文档的格式快照，与格式要求进行对比，找出差异，并生成精确的修改指令。

## 输出格式

你必须返回一个标准的 JSON 对象（不要包含任何额外说明文字），格式如下：

```json
{
  "analysis": "简要分析当前文档格式与要求的差距",
  "issues": [
    "问题1：具体描述",
    "问题2：具体描述"
  ],
  "paragraph_instructions": [
    {
      "indices": [0, 1],
      "font": "SimHei",
      "size_pt": 16,
      "bold": true,
      "italic": null,
      "color": null,
      "alignment": "CENTER",
      "line_spacing": 1.5,
      "first_line_indent_cm": null
    }
  ],
  "global_instructions": {
    "margins": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17}
  },
  "suggestions": [
    "需人工处理的内容1"
  ]
}
```

## 规则

1. **indices** 数组指定要修改的段落索引号。留空 [] 表示所有匹配类型的段落（如所有标题、所有正文）。
2. 只修改与要求**不一致**的格式属性。保持一致的属性设为 `null`。
3. **重点：你必须同时检查字体属性和段落格式属性，两者都需在输出中体现。**
4. **font**: 使用系统字体名（SimHei=黑体, SimSun=宋体, FangSong=仿宋, KaiTi=楷体, Microsoft YaHei=微软雅黑, Times New Roman=Arial替代等）。不修改设为 null。
5. **size_pt**: 字号（磅值）。例如 二号=22, 小二号=18, 三号=16, 小三=15, 四号=14, 小四=12, 五号=10.5。不修改设为 null。
6. **bold/italic**: true/false。不修改设为 null。
7. **alignment**: LEFT / CENTER / RIGHT / BOTH。与要求一致时才设为 null。
8. **line_spacing**: **必须填写**行距倍数。如 1.5。如果文档当前行距与要求一致才设为 null。
9. **first_line_indent_cm**: **必须填写**首行缩进厘米数。如 0.74（约2字符）。如果文档当前缩进与要求一致才设为 null。
10. 如果多个段落需要同样的修改，合并到同一个 instruction 中。
11. **global_instructions.margins** 中的属性也用 null 表示不修改。
12. 如果有任何文字描述的建议（如"建议检查表格格式"），放在 suggestions 数组中。
"""

    user_prompt = f"请分析以下文档格式快照，对比格式要求，输出 JSON 格式的修改指令。\n\n{doc_snapshot_text}\n\n{spec_text_block}"

    # 调用 API
    llm_response = _call_llm(system_prompt, user_prompt)

    # 解析 JSON
    try:
        result = _json_to_instructions(llm_response)
    except (json.JSONDecodeError, ValueError) as e:
        # 如果解析失败，返回原始响应
        return {
            'analysis': f'LLM 响应解析失败: {str(e)}',
            'issues': ['LLM 响应格式异常'],
            'paragraph_instructions': [],
            'global_instructions': {},
            'suggestions': [f'原始响应: {llm_response[:500]}'],
            '_raw': llm_response,
        }

    # 确保所有字段存在
    result.setdefault('analysis', '')
    result.setdefault('issues', [])
    result.setdefault('paragraph_instructions', [])
    result.setdefault('global_instructions', {})
    result.setdefault('suggestions', [])
    result['_mode'] = mode
    result['_raw'] = llm_response

    return result
