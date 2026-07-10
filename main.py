"""
doc-formatter Web 平台 - FastAPI 后端
"""
import os
import uuid
import shutil
from datetime import datetime
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi import Request
from fastapi.templating import Jinja2Templates

from engine import (
    apply_instructions,
    apply_specification,
    apply_benchmark,
    save_to_knowledge_base,
)
from docx import Document
from docx.shared import Pt


def _md_to_docx(md_path):
    """将 Markdown 规范/标杆文档转为临时 .docx，供排版引擎读取"""
    doc = Document()
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    for line in content.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('# ') or line.startswith('## ') or line.startswith('### '):
            level = len(line.split()[0]) - 1
            doc.add_heading(line.lstrip('#').strip(), level=level)
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(line, style='List Bullet')
        else:
            doc.add_paragraph(line)

    tmp_path = md_path.with_suffix('.docx.tmp')
    doc.save(str(tmp_path))
    return tmp_path


def _clean_tmp(base_path):
    """清理由 _md_to_docx / _pdf_text_to_file 生成的临时文件"""
    for suffix in ('.docx.tmp', '.md.tmp'):
        tmp_path = Path(str(base_path) + suffix) if not str(base_path).endswith(suffix) else Path(str(base_path))
        if tmp_path.exists():
            tmp_path.unlink()


def _pdf_text_to_file(pdf_path):
    """从 PDF 提取文字并写出到临时 Markdown 文件，供规范/标杆模式使用"""
    try:
        import pymupdf
    except ImportError:
        return None
    doc = pymupdf.open(str(pdf_path))
    text = ''
    for page in doc:
        text += page.get_text()
    doc.close()

    tmp_path = pdf_path.with_suffix('.md.tmp')
    with open(str(tmp_path), 'w', encoding='utf-8') as f:
        f.write(text)
    return tmp_path

app = FastAPI(title="文档自动排版平台")

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / 'uploads'
STATIC_DIR = BASE_DIR / 'static'
TEMPLATE_DIR = BASE_DIR / 'templates'

# 确保目录存在
for d in [UPLOAD_DIR / 'target', UPLOAD_DIR / 'specification',
          UPLOAD_DIR / 'benchmark', UPLOAD_DIR / 'output']:
    d.mkdir(parents=True, exist_ok=True)

app.mount('/static', StaticFiles(directory=str(STATIC_DIR)), name='static')
templates = Jinja2Templates(directory=str(TEMPLATE_DIR))


@app.get('/', response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse('index.html', {'request': request})


@app.post('/api/upload/{file_type}')
async def upload_file(file_type: str, file: UploadFile = File(...)):
    """上传文件"""
    valid_types = {'target', 'specification', 'benchmark'}

    if file_type not in valid_types:
        raise HTTPException(400, f'无效的文件类型，支持: {valid_types}')

    # target 只允许 .docx；规范/标杆允许 .docx/.md/.pdf
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if file_type == 'target' and ext != 'docx':
        raise HTTPException(400, '待排版文档仅支持 .docx 格式')
    if file_type in ('specification', 'benchmark') and ext not in ('docx', 'md', 'txt', 'pdf'):
        raise HTTPException(400, '规范/标杆文档支持 .docx / .md / .pdf 格式')

    ext = file.filename.rsplit('.', 1)[-1]
    unique_name = f'{uuid.uuid4().hex}_{datetime.now().strftime("%Y%m%d%H%M%S")}.{ext}'
    dest = UPLOAD_DIR / file_type / unique_name

    with open(dest, 'wb') as f:
        content = await file.read()
        f.write(content)

    return {
        'filename': file.filename,
        'stored_name': unique_name,
        'path': str(dest),
        'size': len(content),
    }


@app.post('/api/process')
async def process_document(data: dict):
    """
    处理文档排版
    body: {
        "mode": "instruction" | "specification" | "benchmark",
        "target_file": "上传后的文件名",
        "instructions": "格式要求（模式一用）",
        "spec_file": "规范文件名（模式二用）",
        "benchmark_file": "标杆文档名（模式三用）",
        "save_to_kb": true/false  # 是否保存到知识库
    }
    """
    mode = data.get('mode')
    target_file = data.get('target_file')
    instructions = data.get('instructions', '')
    spec_file = data.get('spec_file')
    benchmark_file = data.get('benchmark_file')
    save_to_kb = data.get('save_to_kb', False)
    original_filename = data.get('original_filename', target_file)  # 用户原始文件名

    if not target_file:
        raise HTTPException(400, '请上传待排版的文档')

    target_path = UPLOAD_DIR / 'target' / target_file
    if not target_path.exists():
        raise HTTPException(404, '目标文件未找到，请重新上传')

    # 生成输出路径（用原始文件名）
    base_name = os.path.splitext(original_filename)[0]
    output_name = f'{base_name}-排版后.docx'
    output_path = UPLOAD_DIR / 'output' / output_name

    try:
        if mode == 'instruction':
            result = apply_instructions(str(target_path), instructions, str(output_path))
        elif mode == 'specification':
            spec_path = None
            source_name = ''
            if spec_file:
                # 先在上传目录找，再在知识库目录找
                spec_path_candidate = UPLOAD_DIR / 'specification' / spec_file
                if spec_path_candidate.exists():
                    spec_path = spec_path_candidate
                    source_name = f'用户上传 ({spec_file})'
                else:
                    # 知识库中查找
                    kb_spec_dir = BASE_DIR / 'knowledge_base' / '01-格式规范'
                    kb_candidate = kb_spec_dir / spec_file
                    if kb_candidate.exists():
                        if kb_candidate.suffix == '.pdf':
                            tmp_md = _pdf_text_to_file(kb_candidate)
                            if tmp_md:
                                spec_path = tmp_md
                                source_name = f'知识库 ({spec_file}, PDF→文本)'
                        elif kb_candidate.suffix == '.md':
                            spec_path = kb_candidate
                            source_name = f'知识库 ({spec_file})'
                        else:
                            spec_path = kb_candidate
                            source_name = f'知识库 ({spec_file})'
                if save_to_kb and spec_path and spec_path.suffix != '.tmp':
                    kb_path = save_to_knowledge_base(str(spec_path), 'specification')
                    if kb_path:
                        result_extra_kb = f'已保存至知识库 01-格式规范/{os.path.basename(kb_path)}'

            if spec_path is None:
                kb_spec = _find_default_spec()
                if kb_spec:
                    if kb_spec.suffix == '.pdf':
                        tmp_md = _pdf_text_to_file(kb_spec)
                        if tmp_md:
                            spec_path = tmp_md
                            source_name = f'{kb_spec.name} (PDF→文本)'
                        else:
                            raise HTTPException(500, '无法读取 PDF 内容（需安装 pymupdf）')
                    else:
                        spec_path = kb_spec
                        source_name = kb_spec.name
                else:
                    raise HTTPException(400, '请上传格式规范文件（未找到知识库默认规范）')

            result = apply_specification(str(target_path), str(spec_path), str(output_path))
            result['report']['source'] = f'知识库规范 ({source_name})'
            # 清理临时转换文件 (PDF→文本)
            if spec_path and str(spec_path).endswith('.tmp'):
                try:
                    spec_path.unlink()
                except:
                    pass

        elif mode == 'benchmark':
            bm_path = None
            source_name = ''
            if benchmark_file:
                bm_path_candidate = UPLOAD_DIR / 'benchmark' / benchmark_file
                if bm_path_candidate.exists():
                    bm_path = bm_path_candidate
                    source_name = f'用户上传 ({benchmark_file})'
                else:
                    kb_bm_dir = BASE_DIR / 'knowledge_base' / '02-标杆文档'
                    kb_candidate = kb_bm_dir / benchmark_file
                    if kb_candidate.exists():
                        if kb_candidate.suffix == '.md':
                            tmp_docx = _md_to_docx(kb_candidate)
                            bm_path = tmp_docx
                            source_name = f'知识库 ({benchmark_file}, Markdown→Word)'
                        elif kb_candidate.suffix == '.pdf':
                            tmp_md = _pdf_text_to_file(kb_candidate)
                            if tmp_md:
                                tmp_docx = _md_to_docx(tmp_md)
                                bm_path = tmp_docx
                                source_name = f'知识库 ({benchmark_file}, PDF→Word)'
                                _clean_tmp(tmp_md)
                        else:
                            bm_path = kb_candidate
                            source_name = f'知识库 ({benchmark_file})'
                if save_to_kb and bm_path and bm_path.suffix != '.tmp':
                    kb_path = save_to_knowledge_base(str(bm_path), 'benchmark')
                    if kb_path:
                        result_extra_kb = f'已保存至知识库 02-标杆文档/{os.path.basename(kb_path)}'

            if bm_path is None:
                kb_bm = _find_default_benchmark()
                if kb_bm:
                    if kb_bm.suffix == '.md':
                        tmp_docx = _md_to_docx(kb_bm)
                        bm_path = tmp_docx
                        source_name = f'{kb_bm.name} (Markdown→Word)'
                    elif kb_bm.suffix == '.pdf':
                        tmp_md = _pdf_text_to_file(kb_bm)
                        if tmp_md:
                            tmp_docx = _md_to_docx(tmp_md)
                            bm_path = tmp_docx
                            source_name = f'{kb_bm.name} (PDF→文本→Word)'
                            # 清理中间临时文件
                            _clean_tmp(tmp_md)
                        else:
                            raise HTTPException(500, '无法读取 PDF 内容（需安装 pymupdf）')
                    else:
                        bm_path = kb_bm
                        source_name = kb_bm.name
                else:
                    raise HTTPException(400, '请上传标杆文档（未找到知识库默认标杆）')

            result = apply_benchmark(str(target_path), str(bm_path), str(output_path))
            result['report']['source'] = f'知识库标杆 ({source_name})'
            # 清理临时转换文件
            if str(bm_path).endswith('.tmp'):
                try:
                    bm_path.unlink()
                except:
                    pass
        else:
            raise HTTPException(400, '无效的模式，请选择 instruction / specification / benchmark')

        result['output_file'] = output_name
        # 修正报告中的文件名为原始文件名（非 UUID）
        result['report']['filename'] = original_filename

        # 自动保存副本到桌面「排版后」目录
        try:
            auto_save_dir = Path('C:/Users/王逸朴/Desktop/智能体大赛/排版后')
            auto_save_dir.mkdir(parents=True, exist_ok=True)
            src = UPLOAD_DIR / 'output' / output_name
            if src.exists():
                import shutil
                dst = auto_save_dir / output_name
                shutil.copy2(str(src), str(dst))
                result['auto_saved_to'] = str(dst)
        except Exception as e:
            # 保存失败不影响主流程
            pass

        return result

    except Exception as e:
        raise HTTPException(500, f'排版处理出错: {str(e)}')


@app.get('/api/download/{filename}')
async def download(filename: str):
    """下载排版后的文档"""
    file_path = UPLOAD_DIR / 'output' / filename
    if not file_path.exists():
        raise HTTPException(404, '文件未找到')
    return FileResponse(
        str(file_path),
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename=filename,
    )


@app.get('/api/knowledge-base')
async def list_knowledge_base():
    """列出知识库中的格式规范和标杆文档"""
    kb_path = BASE_DIR / 'knowledge_base'
    result = {'specifications': [], 'benchmarks': []}

    spec_dir = kb_path / '01-格式规范'
    if spec_dir.exists():
        for f in sorted(spec_dir.iterdir()):
            if f.suffix in ('.md', '.docx', '.txt', '.pdf'):
                ext_info = '📋' if f.suffix in ('.md', '.docx', '.txt') else '📕'
                result['specifications'].append({
                    'name': f.name,
                    'ext': f.suffix,
                    'icon': ext_info,
                    'modified': datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M'),
                })

    bm_dir = kb_path / '02-标杆文档'
    if bm_dir.exists():
        for f in sorted(bm_dir.iterdir()):
            if f.suffix in ('.md', '.docx', '.txt', '.pdf'):
                ext_info = '🏷️' if f.suffix in ('.md', '.docx', '.txt') else '📕'
                result['benchmarks'].append({
                    'name': f.name,
                    'ext': f.suffix,
                    'icon': ext_info,
                    'modified': datetime.fromtimestamp(f.stat().st_mtime).strftime('%Y-%m-%d %H:%M'),
                })

    return result


@app.get('/api/knowledge-base/download/{category}/{filename}')
async def download_kb_file(category: str, filename: str):
    """下载知识库中的文件"""
    category_map = {'spec': '01-格式规范', 'benchmark': '02-标杆文档'}
    dir_name = category_map.get(category)
    if not dir_name:
        raise HTTPException(400, '无效的分类')
    file_path = BASE_DIR / 'knowledge_base' / dir_name / filename
    if not file_path.exists():
        # 也检查原始知识库目录
        alt_path = Path('C:/Users/王逸朴/Desktop/智能体大赛/知识库') / dir_name / filename
        if alt_path.exists():
            file_path = alt_path
    if not file_path.exists():
        raise HTTPException(404, '文件未找到')
    return FileResponse(
        str(file_path),
        filename=filename,
    )


def _find_default_spec():
    """查找知识库中默认的格式规范（优先用 .docx 样式库文件，其次 .md）"""
    spec_dir = BASE_DIR / 'knowledge_base' / '01-格式规范'
    if spec_dir.exists():
        files = list(spec_dir.glob('*.*'))
        docx_files = [f for f in files if f.suffix == '.docx']
        if docx_files:
            return docx_files[0]
        md_files = [f for f in files if f.suffix == '.md']
        if md_files:
            return md_files[0]
        pdf_files = [f for f in files if f.suffix == '.pdf']
        if pdf_files:
            return pdf_files[0]
        if files:
            return files[0]
    return None


def _find_default_benchmark():
    """查找知识库中默认的标杆文档（优先用 .docx，其次 .md）"""
    bm_dir = BASE_DIR / 'knowledge_base' / '02-标杆文档'
    if bm_dir.exists():
        files = list(bm_dir.glob('*.*'))
        docx_files = [f for f in files if f.suffix == '.docx']
        if docx_files:
            return docx_files[0]
        md_files = [f for f in files if f.suffix == '.md']
        if md_files:
            return md_files[0]
        pdf_files = [f for f in files if f.suffix == '.pdf']
        if pdf_files:
            return pdf_files[0]
        if files:
            return files[0]
    return None


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=8000)
