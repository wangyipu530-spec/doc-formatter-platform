/**
 * doc-formatter 平台前端逻辑
 */

// ─── 状态 ───
const state = {
    targetFile: null,      // { name, stored_name }
    specFile: null,        // { name, stored_name, source: 'upload'|'kb' }
    benchmarkFile: null,   // { name, stored_name, source: 'upload'|'kb' }
    processing: false,
};

// ─── DOM 引用 ───
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

// ─── 上传区域功能 ───
function setupUploadZone(zoneId, inputId, statusId, fileNameId, fileSizeId, onUpload) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    const statusEl = document.getElementById(statusId);

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });

    zone.addEventListener('dragleave', () => {
        zone.classList.remove('dragover');
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            input.dispatchEvent(new Event('change'));
        }
    });

    input.addEventListener('change', async () => {
        const file = input.files[0];
        if (!file) return;

        // 如果不是 .docx 且是规范或标杆，可以允许.md/.txt
        const fileType = zoneId.includes('spec') ? 'specification'
                       : zoneId.includes('benchmark') ? 'benchmark'
                       : 'target';

        const allowedExts = fileType === 'target' ? ['.docx'] : ['.docx', '.md', '.txt'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!allowedExts.includes(ext)) {
            showStatus(`仅支持 ${allowedExts.join(', ')} 格式`, 'error');
            return;
        }

        statusEl.textContent = '⏳ 上传中...';

        try {
            const formData = new FormData();
            formData.append('file', file);
            const res = await fetch(`/api/upload/${fileType}`, {
                method: 'POST',
                body: formData,
            });
            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || '上传失败');
            }
            const data = await res.json();

            zone.classList.add('has-file');
            if (fileNameId) {
                const el = document.getElementById(fileNameId);
                if (el) el.textContent = file.name;
            }
            if (fileSizeId) {
                const el = document.getElementById(fileSizeId);
                if (el) el.textContent = `(${(file.size / 1024).toFixed(1)} KB)`;
            }
            statusEl.textContent = '✅ 上传成功';

            if (onUpload) onUpload({ name: file.name, stored_name: data.stored_name });

        } catch (err) {
            statusEl.textContent = `❌ ${err.message}`;
            showStatus(err.message, 'error');
        }
    });
}

// ─── 初始化上传区域 ───
setupUploadZone('targetZone', 'targetFile', 'targetStatus', 'targetFileName', 'targetFileSize',
    (f) => { state.targetFile = f; updateProcessBtn(); });

setupUploadZone('specZone', 'specFile', null, 'specFileName', null,
    (f) => { state.specFile = f; state.specFile.source = 'upload'; updateProcessBtn(); });

setupUploadZone('benchmarkZone', 'benchmarkFile', null, 'benchmarkFileName', null,
    (f) => { state.benchmarkFile = f; state.benchmarkFile.source = 'upload'; updateProcessBtn(); });

// ─── 模式切换 ───
$$('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $$('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        $$('.tab-content').forEach(c => c.classList.remove('active'));
        const tabId = 'tab-' + btn.dataset.tab;
        document.getElementById(tabId).classList.add('active');
        updateProcessBtn();
    });
});

// ─── 知识库选择 ───
document.getElementById('specKbSelect').addEventListener('change', function() {
    if (this.value) {
        // 清空上传文件状态
        state.specFile = {
            name: this.options[this.selectedIndex].text,
            stored_name: this.value,
            source: 'kb',
        };
    } else {
        state.specFile = null;
    }
    updateProcessBtn();
});

document.getElementById('bmKbSelect').addEventListener('change', function() {
    if (this.value) {
        state.benchmarkFile = {
            name: this.options[this.selectedIndex].text,
            stored_name: this.value,
            source: 'kb',
        };
    } else {
        state.benchmarkFile = null;
    }
    updateProcessBtn();
});

// ─── 处理按钮 ───
function getActiveMode() {
    const active = document.querySelector('.tab-btn.active');
    return active ? active.dataset.tab : 'instruction';
}

function updateProcessBtn() {
    const btn = document.getElementById('processBtn');
    const mode = getActiveMode();

    if (!state.targetFile) {
        btn.disabled = true;
        btn.textContent = '⏳ 请先上传待排版文档';
        return;
    }

    if (mode === 'specification' && !state.specFile) {
        btn.disabled = true;
        btn.textContent = '⏳ 请上传格式规范或选择知识库规范';
        return;
    }

    if (mode === 'benchmark' && !state.benchmarkFile) {
        btn.disabled = true;
        btn.textContent = '⏳ 请上传标杆文档或选择知识库标杆';
        return;
    }

    btn.disabled = false;
    btn.textContent = '🚀 开始排版';
}

document.getElementById('processBtn').addEventListener('click', async () => {
    if (state.processing) return;

    const mode = getActiveMode();
    const saveToKb = document.getElementById('saveToKb').checked;

    const payload = {
        mode: mode,
        target_file: state.targetFile.stored_name,
        original_filename: state.targetFile.name,  // 传原始文件名
        instructions: document.getElementById('instructions').value,
        save_to_kb: saveToKb,
    };

    if (mode === 'specification' && state.specFile) {
        payload.spec_file = state.specFile.stored_name;
    }

    if (mode === 'benchmark' && state.benchmarkFile) {
        payload.benchmark_file = state.benchmarkFile.stored_name;
    }

    state.processing = true;
    document.getElementById('loadingOverlay').classList.add('show');

    try {
        const res = await fetch('/api/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || '排版处理失败');
        }

        const data = await res.json();
        showResult(data);
        showStatus('✅ 排版完成！', 'success');

    } catch (err) {
        showStatus(`❌ ${err.message}`, 'error');
    } finally {
        state.processing = false;
        document.getElementById('loadingOverlay').classList.remove('show');
    }
});

// ─── 显示结果 ───
function showResult(data) {
    const area = document.getElementById('resultArea');
    const empty = document.getElementById('resultEmpty');

    empty.style.display = 'none';
    area.classList.add('show');

    // 下载链接
    const dlLink = document.getElementById('downloadLink');
    dlLink.href = `/api/download/${data.output_file}`;
    dlLink.download = data.output_file;

    // 变更报告
    const report = data.report || {};
    const reportEl = document.getElementById('reportContent');

    let html = '';
    html += `<div class="section-title">📄 文档：${report.filename || '-'}</div>`;
    html += `<div class="info">🎯 模式：${report.mode || '-'}</div>`;
    html += `<div class="info">📎 参考来源：${report.source || '-'}</div>`;
    html += `<div class="info">🕐 处理时间：${report.time || '-'}</div>`;

    if (report.changes && report.changes.length > 0) {
        html += `\n<div class="section-title">✅ 已修改项</div>`;
        report.changes.forEach(c => {
            html += `  <span class="success">•</span> ${c}\n`;
        });
    }

    if (report.warnings && report.warnings.length > 0) {
        html += `\n<div class="section-title">⚠️ 需注意</div>`;
        report.warnings.forEach(w => {
            html += `  <span class="warning">•</span> ${w}\n`;
        });
    }

    if (report.suggestions && report.suggestions.length > 0) {
        html += `\n<div class="section-title">💡 AI 建议</div>`;
        report.suggestions.forEach(s => {
            html += `  <span class="info">•</span> ${s}\n`;
        });
    }

    if (report.knowledge_base_update) {
        html += `\n<div class="section-title">📌 知识库更新</div>`;
        html += `  <span class="info">•</span> ${report.knowledge_base_update}\n`;
    }

    if (!report.changes?.length && !report.warnings?.length) {
        html += `\n<span class="info">✅ 文档已符合规范，无需修改</span>`;
    }

    reportEl.innerHTML = html;
}

// ─── 状态栏 ───
function showStatus(msg, type = 'info') {
    const bar = document.getElementById('statusBar');
    bar.textContent = msg;
    bar.className = 'status-bar show ' + type;

    clearTimeout(bar._timer);
    bar._timer = setTimeout(() => {
        bar.classList.remove('show');
    }, 4000);
}

// ─── 知识库 Modal ───
async function toggleKbModal() {
    const modal = document.getElementById('kbModal');
    if (modal.classList.contains('show')) {
        modal.classList.remove('show');
        return;
    }

    modal.classList.add('show');
    const content = document.getElementById('kbContent');
    content.innerHTML = '<div style="text-align:center;padding:20px;color:var(--gray-400)">⏳ 加载中...</div>';

    try {
        const res = await fetch('/api/knowledge-base');
        const data = await res.json();

        let html = '';

        if (data.specifications && data.specifications.length > 0) {
            html += `<h4 style="margin-bottom:8px;font-size:14px;color:var(--gray-600);">📋 格式规范</h4>`;
            html += `<ul class="kb-list">`;
            data.specifications.forEach(s => {
                const extBadge = s.ext === '.pdf' ? '📕 PDF文档' : '格式规范';
                const downloadUrl = `/api/knowledge-base/download/spec/${encodeURIComponent(s.name)}`;
                html += `<li>
                    <div>
                        <div class="name"><a href="${downloadUrl}" target="_blank" style="color:var(--primary);text-decoration:none;">${s.name}</a></div>
                        <div class="meta">${s.modified} · <a href="${downloadUrl}" style="color:var(--gray-400);font-size:11px;" download>⬇️ 下载</a></div>
                    </div>
                    <span class="badge">${extBadge}</span>
                </li>`;
            });
            html += `</ul>`;
        }

        if (data.benchmarks && data.benchmarks.length > 0) {
            html += `<h4 style="margin:12px 0 8px;font-size:14px;color:var(--gray-600);">🏷️ 标杆文档</h4>`;
            html += `<ul class="kb-list">`;
            data.benchmarks.forEach(b => {
                const extBadge = b.ext === '.pdf' ? '📕 PDF文档' : '标杆文档';
                const downloadUrl = `/api/knowledge-base/download/benchmark/${encodeURIComponent(b.name)}`;
                html += `<li>
                    <div>
                        <div class="name"><a href="${downloadUrl}" target="_blank" style="color:var(--primary);text-decoration:none;">${b.name}</a></div>
                        <div class="meta">${b.modified} · <a href="${downloadUrl}" style="color:var(--gray-400);font-size:11px;" download>⬇️ 下载</a></div>
                    </div>
                    <span class="badge">${extBadge}</span>
                </li>`;
            });
            html += `</ul>`;
        }

        if (!html) {
            html = '<div class="empty-state"><div class="icon">📚</div><div class="text">知识库暂无文件，上传时勾选「保存到知识库」即可添加</div></div>';
        }

        content.innerHTML = html;

    } catch (err) {
        content.innerHTML = `<div style="text-align:center;padding:20px;color:var(--danger)">❌ 加载失败: ${err.message}</div>`;
    }
}

// ─── 初始化：加载知识库列表 ───
async function loadKbOptions() {
    try {
        const res = await fetch('/api/knowledge-base');
        const data = await res.json();

        const specSelect = document.getElementById('specKbSelect');
        if (data.specifications) {
            data.specifications.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s.name;
                opt.textContent = s.name;
                specSelect.appendChild(opt);
            });
        }

        const bmSelect = document.getElementById('bmKbSelect');
        if (data.benchmarks) {
            data.benchmarks.forEach(b => {
                const opt = document.createElement('option');
                opt.value = b.name;
                opt.textContent = b.name;
                bmSelect.appendChild(opt);
            });
        }
    } catch (err) {
        console.warn('加载知识库选项失败:', err);
    }
}

loadKbOptions();
