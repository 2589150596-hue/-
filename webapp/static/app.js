const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

let sessionId = generateId();
let currentTaskName = '';
let allServerFiles = [];
let allSummaries = [];
let resultData = {};

// 设置日期选择器默认值为今天
$('#taskDate').value = new Date().toISOString().split('T')[0];

// ========== 页面初始化 ==========
refreshFileList();
refreshSummaries();

$('#btnRefreshFiles').addEventListener('click', refreshFileList);
$('#btnRefreshSummaries').addEventListener('click', refreshSummaries);

// SSE 实时刷新 + 提示音
const evtSource = new EventSource('/api/file-events');
let sseFirstLoad = true;
evtSource.onmessage = () => {
    refreshFileList(); refreshSummaries();
    if (!sseFirstLoad) playBeep();
    sseFirstLoad = false;
};

function playBeep() {
    try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain); gain.connect(ctx.destination);
        osc.frequency.value = 800; osc.type = 'sine';
        gain.gain.value = 0.3;
        osc.start(); osc.stop(ctx.currentTime + 0.15);
    } catch(e) {}
}

// ========== 文件列表 ==========
async function refreshFileList() {
    try {
        const res = await fetch('/api/files');
        const data = await res.json();
        allServerFiles = data.files || [];
        renderFileTable();
        updateFileCount();
    } catch (e) { console.error(e); }
}

function renderFileTable() {
    const tbody = document.querySelector('#fileSelectTable tbody');
    if (allServerFiles.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;padding:30px">暂无上传文件</td></tr>';
        return;
    }
    const groups = {};
    allServerFiles.forEach((f, i) => {
        const dateKey = f.task_date || '未知日期';
        if (!groups[dateKey]) groups[dateKey] = [];
        groups[dateKey].push({...f, _idx: i});
    });
    let html = '';
    const orderedDates = Object.keys(groups);
    for (const dateKey of orderedDates) {
        const files = groups[dateKey];
        html += `<tr class="date-group"><td colspan="4"><span class="date-tag">&#128197; ${dateKey}</span><span class="date-count">${files.length} 个文件</span></td></tr>`;
        for (const f of files) {
            const timeStr = new Date(f.mtime * 1000).toLocaleTimeString('zh-CN', {hour:'2-digit',minute:'2-digit'});
            const encFn = encodeURIComponent(f.filename);
            const badge = f.processed ? '<span class="processed-badge">已处理</span>' : '';
            html += `<tr class="${f.processed ? 'row-processed' : ''}">
                <td><input type="checkbox" class="file-check" data-idx="${f._idx}" ${f._checked ? 'checked' : ''}></td>
                <td>${f.filename} ${badge}</td>
                <td style="color:#999">${timeStr}</td>
                <td><a href="/api/download-file/${f.session_id}/${encFn}" class="btn-op" title="下载">&#128229;</a>
                    <span class="btn-op btn-del" data-sid="${f.session_id}" data-fn="${f.filename}" title="删除">&#128465;</span></td>
            </tr>`;
        }
    }
    tbody.innerHTML = html;
    tbody.querySelectorAll('.file-check').forEach(cb => {
        cb.addEventListener('change', () => {
            allServerFiles[parseInt(cb.dataset.idx)]._checked = cb.checked;
            updateFileCount(); updateProcessBtn(); updateCheckAll();
        });
    });
    tbody.querySelectorAll('.btn-del').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            e.stopPropagation();
            if (!confirm(`确定删除 "${btn.dataset.fn}" 吗？`)) return;
            await fetch('/api/delete-file', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:btn.dataset.sid,filename:btn.dataset.fn})});
            refreshFileList();
        });
    });
    updateCheckAll();
    updateProcessBtn();
}

function updateFileCount() {
    const checked = allServerFiles.filter(f => f._checked).length;
    $('#fileCount').textContent = `已选 ${checked} / ${allServerFiles.length} 个文件`;
}
function updateProcessBtn() {
    $('#btnProcess').disabled = !allServerFiles.some(f => f._checked);
}
function updateCheckAll() {
    const c = allServerFiles.filter(f => f._checked).length;
    $('#checkAll').checked = c > 0 && c === allServerFiles.length;
    $('#checkAll').indeterminate = c > 0 && c < allServerFiles.length;
}

$('#checkAll').addEventListener('change', function() {
    allServerFiles.forEach(f => f._checked = this.checked);
    renderFileTable(); updateFileCount();
});
$('#btnSelectAll').addEventListener('click', () => { allServerFiles.forEach(f => f._checked = true); renderFileTable(); updateFileCount(); });
$('#btnDeselectAll').addEventListener('click', () => { allServerFiles.forEach(f => f._checked = false); renderFileTable(); updateFileCount(); });

// 批量下载
$('#btnBatchDownload').addEventListener('click', async () => {
    const sel = allServerFiles.filter(f => f._checked).map(f => ({session_id:f.session_id,filename:f.filename}));
    if (!sel.length) { alert('请先勾选文件'); return; }
    const res = await fetch('/api/batch-download', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({files:sel})});
    if (res.ok) { const b=await res.blob(); const a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download='批量下载.zip'; a.click(); }
    else alert('下载失败');
});
// 批量删除
$('#btnBatchDelete').addEventListener('click', async () => {
    const sel = allServerFiles.filter(f => f._checked).map(f => ({session_id:f.session_id,filename:f.filename}));
    if (!sel.length) { alert('请先勾选文件'); return; }
    if (!confirm(`确定删除选中的 ${sel.length} 个文件吗？`)) return;
    await fetch('/api/batch-delete', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({files:sel})});
    refreshFileList();
});

// ========== 顾客上传 ==========
const customerUploadZone = $('#customerUploadZone');
const customerFileInput = $('#customerFileInput');
customerUploadZone.addEventListener('click', () => customerFileInput.click());
customerUploadZone.addEventListener('dragover', e => { e.preventDefault(); customerUploadZone.style.borderColor = '#1890ff'; });
customerUploadZone.addEventListener('dragleave', () => { customerUploadZone.style.borderColor = '#d9d9d9'; });
customerUploadZone.addEventListener('drop', e => { e.preventDefault(); customerUploadZone.style.borderColor = '#d9d9d9'; customerFileInput.files = e.dataTransfer.files; doCustomerUpload(); });
customerFileInput.addEventListener('change', doCustomerUpload);

async function doCustomerUpload() {
    const customerName = $('#customerName').value.trim();
    if (!customerName) { alert('请先填写顾客名称'); return; }
    const files = customerFileInput.files;
    if (!files || !files.length) return;
    const fd = new FormData();
    fd.append('customer_name', customerName);
    fd.append('task_date', $('#taskDate').value || new Date().toISOString().split('T')[0]);
    fd.append('session_id', sessionId);
    let n = 0;
    for (const f of files) { if (f.name.match(/\.xlsx?$/i)) { fd.append('files', f); n++; } }
    if (!n) return;
    showMsg(`正在上传 ${n} 个文件...`, 'processing');
    $('#btnCustomerUpload').disabled = true;
    const res = await fetch('/api/upload', {method:'POST',body:fd});
    const data = await res.json();
    if (data.error) { showMsg(data.error, 'error'); alert('上传失败：'+data.error); }
    else { showMsg(`上传成功：${data.count} 个文件`, 'done'); alert(`上传成功！\n任务日期：${$('#taskDate').value}\n共 ${data.count} 个文件`); customerFileInput.value = ''; }
    $('#btnCustomerUpload').disabled = false;
}
$('#btnCustomerUpload').addEventListener('click', doCustomerUpload);
function showMsg(msg, type) { const el = $('#customerUploadMsg'); el.textContent = msg; el.className = 'upload-msg msg-' + type; }

// ========== 处理 ==========
$('#btnProcess').addEventListener('click', async () => {
    const sel = allServerFiles.filter(f => f._checked).map(f => ({session_id:f.session_id,filename:f.filename}));
    if (!sel.length) { alert('请先勾选文件'); return; }
    setStatus(`正在处理 ${sel.length} 个文件...`, 'processing');
    $('#btnProcess').disabled = true; $('#btnDownload').disabled = true;
    const res = await fetch('/api/process', {method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({session_id:sessionId, start_time:$('#startTime').value||'08:00', end_time:$('#endTime').value||'23:00',
            num_workers:parseInt($('#numWorkers').value)||1, min_interval:parseFloat($('#minInterval').value)||20, selected_files:sel})});
    const data = await res.json();
    if (data.error) { setStatus('处理失败', 'error'); alert(data.error); $('#btnProcess').disabled = false; return; }
    resultData = data;
    currentTaskName = data.task_name;
    $('#statOrders').textContent = data.total_orders;
    $('#statFiles').textContent = data.total_files;
    $('#statCustomers').textContent = data.total_customers;
    $('#statShops').textContent = data.total_shops;
    $('#resultArea').style.display = 'block';
    $('#btnDownload').disabled = false;
    $('#btnProcess').disabled = false;
    switchTab('summary');
    setStatus(`处理完成！${data.total_rows} 行 → ${currentTaskName}`, 'done');

    // 异常弹窗提醒
    if (data.anomalies && data.anomalies.length > 0) {
        const msgs = data.anomalies.map(a => `• ${a['文件名']}：${a['异常原因']}`).join('\n');
        alert(`⚠ 发现 ${data.anomalies.length} 个异常，已跳过处理：\n\n${msgs}`);
    }
});

// ========== 下载最新汇总 ==========
$('#btnDownload').addEventListener('click', () => {
    if (currentTaskName) window.open('/api/download-summary/' + encodeURIComponent(currentTaskName), '_blank');
});

// ========== 汇总历史 ==========
async function refreshSummaries() {
    try {
        const res = await fetch('/api/summaries');
        const data = await res.json();
        allSummaries = data.summaries || [];
        renderSummaryTable();
    } catch (e) { console.error(e); }
}

function renderSummaryTable() {
    const tbody = document.querySelector('#summaryTable tbody');
    if (!allSummaries.length) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;padding:30px">暂无汇总记录</td></tr>';
        return;
    }
    const groups = {};
    allSummaries.forEach((s, i) => {
        const d = new Date(s.mtime * 1000);
        const dateKey = d.toLocaleDateString('zh-CN');
        if (!groups[dateKey]) groups[dateKey] = [];
        groups[dateKey].push({...s, _idx: i});
    });
    let html = '';
    const orderedDates = Object.keys(groups);
    for (const dateKey of orderedDates) {
        const items = groups[dateKey];
        html += `<tr class="date-group"><td colspan="4"><span class="date-tag">&#128197; ${dateKey}</span><span class="date-count">${items.length} 个汇总</span></td></tr>`;
        for (const s of items) {
            const timeStr = new Date(s.mtime * 1000).toLocaleTimeString('zh-CN', {hour:'2-digit',minute:'2-digit'});
            const sizeStr = (s.size / 1024 / 1024).toFixed(1) + ' MB';
            const encFn = encodeURIComponent(s.filename);
            html += `<tr>
                <td><input type="checkbox" class="sum-check" data-idx="${s._idx}" ${s._checked ? 'checked' : ''}></td>
                <td>${s.filename} <span style="color:#999;font-size:11px">(${sizeStr})</span></td>
                <td style="color:#999">${timeStr}</td>
                <td><a href="/api/download-summary/${encFn}" class="btn-op" title="下载">&#128229;</a>
                    <span class="btn-op btn-del" data-fn="${s.filename}" title="删除">&#128465;</span></td>
            </tr>`;
        }
    }
    tbody.innerHTML = html;
    tbody.querySelectorAll('.sum-check').forEach(cb => {
        cb.addEventListener('change', () => {
            allSummaries[parseInt(cb.dataset.idx)]._checked = cb.checked;
            updateSumCount();
        });
    });
    tbody.querySelectorAll('.btn-del').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            e.stopPropagation();
            if (!confirm(`确定删除 "${btn.dataset.fn}" 吗？`)) return;
            await fetch('/api/delete-summary', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:btn.dataset.fn})});
            refreshSummaries();
        });
    });
    updateSumCount();
}

function updateSumCount() {
    const c = allSummaries.filter(s => s._checked).length;
    $('#sumCount').textContent = `已选 ${c} / ${allSummaries.length} 个汇总`;
}

$('#checkAllSum').addEventListener('change', function() {
    allSummaries.forEach(s => s._checked = this.checked);
    renderSummaryTable();
});
$('#btnSelectAllSum').addEventListener('click', () => { allSummaries.forEach(s => s._checked = true); renderSummaryTable(); });
$('#btnDeselectAllSum').addEventListener('click', () => { allSummaries.forEach(s => s._checked = false); renderSummaryTable(); });
$('#btnBatchDownloadSum').addEventListener('click', () => {
    const sel = allSummaries.filter(s => s._checked);
    if (!sel.length) { alert('请先勾选汇总表'); return; }
    sel.forEach(s => window.open('/api/download-summary/' + encodeURIComponent(s.filename), '_blank'));
});
$('#btnBatchDeleteSum').addEventListener('click', async () => {
    const sel = allSummaries.filter(s => s._checked).map(s => s.filename);
    if (!sel.length) { alert('请先勾选汇总表'); return; }
    if (!confirm(`确定删除选中的 ${sel.length} 个汇总表吗？`)) return;
    await fetch('/api/batch-delete-summaries', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filenames:sel})});
    refreshSummaries();
});

// ========== Tab 切换 ==========
$$('.tab').forEach(tab => tab.addEventListener('click', () => switchTab(tab.dataset.tab)));
function switchTab(name) {
    $$('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
    let rows=[], cols=[];
    switch(name) {
        case 'summary': rows=resultData.summary||[]; cols=resultData.summary_columns||[]; break;
        case 'shop': rows=resultData.shop_schedule||[]; cols=['店铺名称','店铺总单量','平均间隔(分钟)','预计完成时间']; break;
        case 'file': rows=resultData.file_stats||[]; cols=['文件名','客户名','补单数量']; break;
        case 'customer': rows=resultData.customer_stats||[]; cols=['客户名','补单数量']; break;
        case 'anomaly': rows=resultData.anomalies||[]; cols=['文件名','异常原因']; break;
    }
    const thead=document.querySelector('#dataTable thead'), tbody=document.querySelector('#dataTable tbody');
    thead.innerHTML='<tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr>';
    tbody.innerHTML=rows.length?rows.map(r=>'<tr>'+cols.map(c=>`<td>${r[c]!==undefined?r[c]:''}</td>`).join('')+'</tr>').join(''):'<tr><td colspan="'+cols.length+'" style="text-align:center;color:#999;padding:40px">暂无数据</td></tr>';
}

function setStatus(msg,cls) { const el=$('#statusBar'); el.textContent=msg; el.className='status-bar '+(cls||''); }
function generateId() { return 'xxxxxxxxxxxx4xxxyxxxxxxxxxxxxxxx'.replace(/[xy]/g,c=>{const r=Math.random()*16|0,v=c==='x'?r:(r&0x3|0x8);return v.toString(16);}); }
