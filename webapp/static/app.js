const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

let sessionId = '';
let outputId = '';
let uploadedFiles = [];
let resultData = {};

// ========== 上传 ==========
const uploadZone = $('#uploadZone');
const fileInput = $('#fileInput');
const fileList = $('#fileList');
const statusBar = $('#statusBar');
const btnProcess = $('#btnProcess');
const btnDownload = $('#btnDownload');

uploadZone.addEventListener('click', () => fileInput.click());
uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('dragover'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
uploadZone.addEventListener('drop', e => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', () => handleFiles(fileInput.files));

async function handleFiles(fileList) {
    const formData = new FormData();
    let count = 0;
    for (const f of fileList) {
        if (f.name.toLowerCase().endsWith('.xlsx') || f.name.toLowerCase().endsWith('.xls')) {
            formData.append('files', f);
            count++;
        }
    }
    if (count === 0) return;

    if (!sessionId) sessionId = generateId();
    formData.append('session_id', sessionId);

    setStatus('正在上传...', 'processing');
    try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await res.json();
        if (data.error) { setStatus(data.error, 'error'); return; }
        uploadedFiles = data.files;
        sessionId = data.session_id;
        renderFileTags();
        uploadZone.classList.add('has-files');
        btnProcess.disabled = false;
        setStatus(`已上传 ${uploadedFiles.length} 个文件，可以开始处理`, '');
    } catch (e) {
        setStatus('上传失败: ' + e.message, 'error');
    }
}

function renderFileTags() {
    fileList.innerHTML = uploadedFiles.map((f, i) =>
        `<span class="file-tag">${f}<span class="remove" data-idx="${i}">&times;</span></span>`
    ).join('');
    fileList.querySelectorAll('.remove').forEach(el => {
        el.addEventListener('click', e => {
            e.stopPropagation();
            const idx = parseInt(el.dataset.idx);
            uploadedFiles.splice(idx, 1);
            if (uploadedFiles.length === 0) {
                uploadZone.classList.remove('has-files');
                btnProcess.disabled = true;
            }
            renderFileTags();
        });
    });
}

// ========== 处理 ==========
btnProcess.addEventListener('click', async () => {
    const startTime = $('#startTime').value || '08:00';
    const endTime = $('#endTime').value || '23:00';
    const numWorkers = parseInt($('#numWorkers').value) || 1;
    const minInterval = parseFloat($('#minInterval').value) || 20;

    setStatus('正在处理中，大文件可能需要数十秒...', 'processing');
    btnProcess.disabled = true;
    btnDownload.disabled = true;

    try {
        const res = await fetch('/api/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: sessionId, start_time: startTime, end_time: endTime, num_workers: numWorkers, min_interval: minInterval })
        });
        const data = await res.json();
        if (data.error) { setStatus('处理失败: ' + data.error, 'error'); btnProcess.disabled = false; return; }

        resultData = data;
        outputId = data.output_id;

        $('#statOrders').textContent = data.total_orders;
        $('#statFiles').textContent = data.total_files;
        $('#statCustomers').textContent = data.total_customers;
        $('#statShops').textContent = data.total_shops;

        $('#resultArea').style.display = 'block';
        btnDownload.disabled = false;
        btnProcess.disabled = false;

        switchTab('summary');
        setStatus(`处理完成！共 ${data.total_rows} 行数据`, 'done');
    } catch (e) {
        setStatus('处理失败: ' + e.message, 'error');
        btnProcess.disabled = false;
    }
});

// ========== Tab 切换 ==========
$$('.tab').forEach(tab => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
});

function switchTab(name) {
    $$('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
    let rows = [], cols = [];
    switch (name) {
        case 'summary': rows = resultData.summary || []; cols = resultData.summary_columns || []; break;
        case 'shop': rows = resultData.shop_schedule || []; cols = ['店铺名称', '店铺总单量', '平均间隔(分钟)', '预计完成时间']; break;
        case 'file': rows = resultData.file_stats || []; cols = ['文件名', '客户名', '补单数量']; break;
        case 'customer': rows = resultData.customer_stats || []; cols = ['客户名', '补单数量']; break;
        case 'anomaly': rows = resultData.anomalies || []; cols = ['文件名', '异常原因']; break;
    }
    renderTable(cols, rows);
}

function renderTable(cols, rows) {
    const thead = document.querySelector('#dataTable thead');
    const tbody = document.querySelector('#dataTable tbody');
    thead.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
    tbody.innerHTML = rows.length === 0
        ? '<tr><td colspan="' + cols.length + '" style="text-align:center;color:#999;padding:40px">暂无数据</td></tr>'
        : rows.map(row => '<tr>' + cols.map(c => `<td>${row[c] !== undefined ? row[c] : ''}</td>`).join('') + '</tr>').join('');
}

// ========== 下载 ==========
btnDownload.addEventListener('click', () => {
    if (sessionId && outputId) {
        window.open(`/api/download/${sessionId}/${outputId}`, '_blank');
    }
});

// ========== 清空 ==========
$('#btnClear').addEventListener('click', () => {
    sessionId = '';
    outputId = '';
    uploadedFiles = [];
    resultData = {};
    fileList.innerHTML = '';
    uploadZone.classList.remove('has-files');
    btnProcess.disabled = true;
    btnDownload.disabled = true;
    $('#resultArea').style.display = 'none';
    setStatus('就绪 - 请先上传 Excel 文件', '');
});

// ========== 工具 ==========
function setStatus(msg, cls) {
    statusBar.textContent = msg;
    statusBar.className = 'status-bar ' + cls;
}

function generateId() {
    return 'xxxxxxxxxxxx4xxxyxxxxxxxxxxxxxxx'.replace(/[xy]/g, c => {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}
