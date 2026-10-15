/*!
 * 办公文档转换器 — 前端交互
 *
 * 结构：
 *   1. 工具函数 / 状态      2. 主题       3. 上传与文件列表
 *   4. 各功能模块           5. 操作分发   6. 结果与错误展示
 *   7. PDF 编辑器
 *
 * 后端接口契约（与 main.py 一致，勿随意改动参数名）：
 *   GET  /api/formats   /api/config
 *   POST /api/convert            file, target
 *   POST /api/compress           file, level
 *   POST /api/batch              files[], mode, target|level
 *   POST /api/pdf-folder-to-jpg  files[]
 *   POST /api/scan-photo         file
 *   POST /api/pdf-merge          files[], mode, insert_index(0-based)
 *   POST /api/pdf-preview        file
 *   POST /api/pdf-edit           file, delete_pages(1-based), rotations, page_order(1-based), annotations
 *   POST /api/translate          file, output, domain, glossary, translate_names, skip_references
 *   GET  /api/translate/progress/{task_id}
 *   GET  /api/download/{task_id}
 */

/* ============================================================
 * 1. 工具函数与全局状态
 * ============================================================ */
const $ = (id) => document.getElementById(id);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const MODES = ['convert', 'compress', 'pdfjpg', 'scan', 'pdfmerge', 'pdfedit', 'translate'];
const TRANSLATE_EXTS = ['pdf', 'docx', 'txt', 'md', 'markdown'];
const SCAN_EXTS = ['png', 'jpg', 'jpeg', 'jfif', 'bmp', 'webp', 'gif', 'tiff'];
const COMPRESS_LEVELS = ['80', '60', '40', '20'];
/** 合并接口的 insert_index 是 0-based「插入到第 N 页之后」，界面上按 1-based 展示。 */
const INSERT_LABEL_TO_API = (value) => Math.max(0, Math.min(1e6, Math.trunc(value))) - 1;

const state = {
  formats: {},
  limits: { maxFileMb: 200, maxBatchMb: 500 },
  translateReady: false,
  translateModel: '',
  mode: 'convert',
  files: [],            // 当前已选文件（顺序即处理顺序）
  est: null,            // 压缩预估缓存 { key, seq, results, error }
  busy: false,
  editorLoaded: false,
  lastError: '',
};

/* ---------- 小工具 ---------- */
function el(tag, props = {}, children = []) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (key === 'class') node.className = value;
    else if (key === 'text') node.textContent = value;
    else if (key === 'html') node.innerHTML = value;
    else if (key === 'dataset') Object.assign(node.dataset, value);
    else if (key.startsWith('on') && typeof value === 'function') node.addEventListener(key.slice(2), value);
    else if (value === true) node.setAttribute(key, '');
    else if (value !== false && value != null) node.setAttribute(key, value);
  }
  for (const child of [].concat(children)) {
    if (child == null) continue;
    node.append(child.nodeType ? child : document.createTextNode(String(child)));
  }
  return node;
}

function ext(name) {
  const idx = String(name || '').lastIndexOf('.');
  return idx < 0 ? '' : String(name).slice(idx + 1).toLowerCase();
}

function humanSize(bytes) {
  const n = Number(bytes) || 0;
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
  if (n < 1024 * 1024 * 1024) return (n / 1048576).toFixed(2) + ' MB';
  return (n / 1073741824).toFixed(2) + ' GB';
}

function targetLabel(target) {
  const names = { docx: 'Word（可编辑 .docx）', pdf: 'PDF', txt: '纯文本 .txt', html: '网页 .html',
    csv: 'CSV 表格', xlsx: 'Excel .xlsx', png: 'PNG 图片', jpg: 'JPG 图片', ico: 'ICO 图标',
    gif: 'GIF', bmp: 'BMP', webp: 'WEBP', tiff: 'TIFF', jfif: 'JFIF', pptx: 'PPT .pptx' };
  return names[target] || ('.' + target);
}

function download(url, filename) {
  const a = el('a', { href: url, download: filename || '' });
  document.body.append(a);
  a.click();
  a.remove();
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

/* ---------- 旋转坐标换算 ----------
 * 标注统一以「未旋转的 PDF 坐标」存储（后端也按此写入）。
 * 页面旋转后，显示区域是旋转后的矩形，画布点击与绘制都需要在这两个坐标系之间转换。
 */
function displayDims(page) {
  const raw = page.raw || { w: page.width, h: page.height };
  const quarter = ((page.rotation || 0) % 360 + 360) % 360;
  return (quarter === 90 || quarter === 270) ? { w: raw.h, h: raw.w } : { w: raw.w, h: raw.h };
}

function pointToPdf(page, x, y) {
  const raw = page.raw || { w: page.width, h: page.height };
  const quarter = ((page.rotation || 0) % 360 + 360) % 360;
  if (quarter === 90) return [y, raw.w - x];
  if (quarter === 180) return [raw.w - x, raw.h - y];
  if (quarter === 270) return [raw.h - y, x];
  return [x, y];
}

function pointToDisplay(page, x, y) {
  const raw = page.raw || { w: page.width, h: page.height };
  const quarter = ((page.rotation || 0) % 360 + 360) % 360;
  if (quarter === 90) return [raw.w - y, x];
  if (quarter === 180) return [raw.w - x, raw.h - y];
  if (quarter === 270) return [y, raw.h - x];
  return [x, y];
}

function rectToDisplay(page, rect) {
  const [x0, y0, x1, y1] = rect;
  const [dx0, dy0] = pointToDisplay(page, x0, y0);
  const [dx1, dy1] = pointToDisplay(page, x1, y1);
  return [Math.min(dx0, dx1), Math.min(dy0, dy1), Math.max(dx0, dx1), Math.max(dy0, dy1)];
}

/* ---------- Toast ---------- */
function toast(message, kind = 'info', timeout = 5200) {
  const icons = { info: 'ℹ️', ok: '✅', warn: '⚠️', error: '❌' };
  const node = el('div', { class: 'toast ' + kind }, [
    el('span', { class: 't-icon', text: icons[kind] || icons.info }),
    el('div', { text: message }),
  ]);
  $('toasts').append(node);
  const remove = () => {
    node.classList.add('leaving');
    setTimeout(() => node.remove(), 220);
  };
  node.addEventListener('click', remove);
  setTimeout(remove, timeout);
}

/* ---------- 统一请求处理 ---------- */
async function readError(resp) {
  try {
    const data = await resp.json();
    if (typeof data?.detail === 'string') return data.detail;
    if (Array.isArray(data?.detail)) return data.detail.map((d) => d.msg || JSON.stringify(d)).join('；');
    return JSON.stringify(data);
  } catch {
    try { return (await resp.text()).slice(0, 300) || `HTTP ${resp.status}`; }
    catch { return `HTTP ${resp.status}`; }
  }
}

async function request(url, { body, method = 'POST' } = {}) {
  let resp;
  try {
    resp = await fetch(url, { method, body });
  } catch (err) {
    throw new Error('无法连接到后端服务，请确认服务仍在运行（' + err.message + '）');
  }
  if (!resp.ok) throw new Error(await readError(resp));
  if (resp.status === 204) return {};
  return resp.json();
}

/* ============================================================
 * 2. 主题
 * ============================================================ */
const THEME_KEY = 'doc-converter-theme';

function applyTheme(theme) {
  const dark = theme === 'dark';
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  $('themeToggle').textContent = dark ? '☀️' : '🌙';
  $('themeToggle').title = dark ? '切换到浅色主题' : '切换到深色主题';
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  applyTheme(saved || (prefersDark ? 'dark' : 'light'));
  $('themeToggle').addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    localStorage.setItem(THEME_KEY, next);
    applyTheme(next);
  });
}

/* ============================================================
 * 3. 上传与文件列表
 * ============================================================ */
const dropzone = $('dropzone');
const fileInput = $('fileInput');
const folderInput = $('folderInput');

function validateFiles(list, currentMode = state.mode) {
  if (!list.length) return '没有有效的文件';
  const maxFile = state.limits.maxFileMb * 1024 * 1024;
  const tooBig = list.find((f) => f.size > maxFile);
  if (tooBig) return `文件「${tooBig.name}」超过单文件 ${state.limits.maxFileMb}MB 上限`;
  const total = list.reduce((sum, f) => sum + f.size, 0);
  if (total > state.limits.maxBatchMb * 1024 * 1024) {
    return `所选文件共 ${humanSize(total)}，超过批量 ${state.limits.maxBatchMb}MB 上限`;
  }
  if (currentMode === 'scan' && list.length !== 1) return '扫描照片一次只能处理一张图片';
  if (currentMode === 'pdfedit' && list.length !== 1) return 'PDF 编辑一次只能处理一个文件';
  if (currentMode === 'translate' && list.length !== 1) return '文献翻译一次只能处理一个文件';
  return '';
}

function acceptFiles(list) {
  const filtered = list.filter((f) => f.name.includes('.'));
  if (!filtered.length) {
    showActionError('没有识别到带扩展名的文件');
    toast('没有识别到带扩展名的文件', 'warn');
    return;
  }
  const problem = validateFiles(filtered, state.mode);
  if (problem) {
    showActionError(problem);
    toast(problem, 'error');
    return;
  }
  state.files = filtered;
  // 换了文件，之前针对旧文件的预估与编辑结果全部作废
  state.est = null;
  resetEditor();
  clearResults();
  renderFileBar();
  renderEstimate();
  updateModeUI();
  if (state.mode === 'pdfedit') loadEditor();
}

function clearFiles() {
  state.files = [];
  state.est = null;
  fileInput.value = '';
  folderInput.value = '';
  resetEditor();
  clearResults();
  renderFileBar();
  renderEstimate();
  updateModeUI();
}

function renderFileBar() {
  const files = state.files;
  const bar = $('fileBar');
  const listBox = $('fileList');
  if (!files.length) {
    bar.classList.add('hidden');
    listBox.classList.add('hidden');
    $('pdfQuickActions').classList.add('hidden');
    return;
  }

  const total = files.reduce((sum, f) => sum + f.size, 0);
  $('fileName').textContent = files.length === 1 ? files[0].name : `已选择 ${files.length} 个文件`;
  $('fileSub').textContent = files.length === 1 ? '' : '按以下顺序处理';
  const chips = $('fileChips');
  chips.replaceChildren();
  if (files.length === 1) {
    chips.append(el('span', { class: 'chip fmt', text: '.' + ext(files[0].name) }));
    chips.append(el('span', { class: 'chip size', text: humanSize(files[0].size) }));
  } else {
    const exts = [...new Set(files.map((f) => '.' + ext(f.name)))];
    chips.append(el('span', { class: 'chip fmt', text: exts.length <= 4 ? exts.join(' ') : `${exts.length} 种格式` }));
    chips.append(el('span', { class: 'chip size', text: '共 ' + humanSize(total) }));
    const pdfCount = files.filter((f) => ext(f.name) === 'pdf').length;
    if (pdfCount) chips.append(el('span', { class: 'chip', text: pdfCount + ' 个 PDF' }));
  }
  bar.classList.remove('hidden');

  const reorderable = state.mode === 'pdfmerge';
  const body = $('fileListBody');
  body.replaceChildren();
  files.forEach((file, index) => {
    const row = el('div', { class: 'file-row', dataset: { index: String(index) } }, [
      el('span', { class: 'fr-name', text: `${index + 1}. ${file.name}`, title: file.name }),
      el('span', { class: 'fr-meta', text: humanSize(file.size) }),
    ]);
    if (reorderable) {
      row.setAttribute('draggable', 'true');
      row.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('text/plain', String(index));
        e.dataTransfer.effectAllowed = 'move';
        row.style.opacity = '.4';
      });
      row.addEventListener('dragend', () => { row.style.opacity = ''; });
      row.addEventListener('dragover', (e) => { e.preventDefault(); row.style.borderTop = '2px solid var(--brand-1)'; });
      row.addEventListener('dragleave', () => { row.style.borderTop = ''; });
      row.addEventListener('drop', (e) => {
        e.preventDefault();
        row.style.borderTop = '';
        const from = Number(e.dataTransfer.getData('text/plain'));
        if (!Number.isFinite(from) || from === index) return;
        const [moved] = state.files.splice(from, 1);
        state.files.splice(index, 0, moved);
        state.est = null;
        renderFileBar();
        updateModeUI();
      });
    }
    const remove = el('button', { class: 'fr-drop', type: 'button', title: '移除这个文件', text: '✕' });
    remove.addEventListener('click', () => {
      state.files.splice(index, 1);
      state.est = null;
      if (state.mode === 'pdfedit') resetEditor();
      clearResults();
      renderFileBar();
      renderEstimate();
      updateModeUI();
    });
    row.append(remove);
    body.append(row);
  });

  listBox.classList.toggle('hidden', files.length < 2);
  $('fileListTitle').textContent = reorderable ? 'PDF 顺序（可拖拽调整）' : '已选文件';
  $('fileListTotal').textContent = files.length + ' 个 · ' + humanSize(total);

  const hasPdf = files.some((f) => ext(f.name) === 'pdf');
  const quick = $('pdfQuickActions');
  quick.classList.toggle('hidden', !hasPdf);
  $('quickPdfEdit').classList.toggle('hidden', !(files.length === 1 && ext(files[0].name) === 'pdf'));
  $('quickPdfMerge').classList.toggle('hidden', !(files.filter((f) => ext(f.name) === 'pdf').length >= 2));
  $('quickTranslate').classList.toggle('hidden', !(files.length === 1 && TRANSLATE_EXTS.includes(ext(files[0].name))));
}

/* ---------- 拖拽 / 选择 ---------- */
dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
  dropzone.classList.add('dragover');
});
dropzone.addEventListener('dragleave', (e) => {
  if (!dropzone.contains(e.relatedTarget)) dropzone.classList.remove('dragover');
});
dropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const items = [...(e.dataTransfer.items || [])];
  const hasDir = items.some((item) => item.webkitGetAsEntry?.()?.isDirectory);
  if (hasDir) {
    showActionRunning('正在读取文件夹内容…');
    const files = await readDroppedItems(items);
    acceptFiles(files);
  } else if (e.dataTransfer.files.length) {
    acceptFiles([...e.dataTransfer.files]);
  }
});
$('pickFiles').addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
$('pickFolder').addEventListener('click', (e) => { e.stopPropagation(); folderInput.click(); });
fileInput.addEventListener('change', () => {
  if (fileInput.files.length) acceptFiles([...fileInput.files]);
});
folderInput.addEventListener('change', () => {
  if (folderInput.files.length) acceptFiles([...folderInput.files]);
});
$('clearFiles').addEventListener('click', clearFiles);

function readDirectory(entry) {
  return new Promise((resolve) => {
    const reader = entry.createReader();
    const result = [];
    const readBatch = () => reader.readEntries((entries) => {
      if (!entries.length) return resolve(result);
      Promise.all(entries.map(readEntry)).then((files) => {
        result.push(...files.flat());
        readBatch();
      });
    }, () => resolve(result));
    readBatch();
  });
}

function readEntry(entry) {
  if (entry.isFile) {
    return new Promise((resolve) => entry.file((file) => resolve([file]), () => resolve([])));
  }
  return entry.isDirectory ? readDirectory(entry) : Promise.resolve([]);
}

async function readDroppedItems(items) {
  const entries = items.map((item) => item.webkitGetAsEntry?.()).filter(Boolean);
  const files = (await Promise.all(entries.map(readEntry))).flat();
  const unique = new Map();
  files.filter((file) => file.name.includes('.')).forEach((file) => {
    unique.set(`${file.name}:${file.size}:${file.lastModified}`, file);
  });
  return [...unique.values()];
}

/* ============================================================
 * 4. 各功能模块
 * ============================================================ */

/* ---------- 4.1 格式转换 ---------- */
function availableTargets(files) {
  if (!files.length) return [];
  let targets = state.formats[ext(files[0].name)] || [];
  for (const file of files.slice(1)) {
    const other = state.formats[ext(file.name)] || [];
    targets = targets.filter((t) => other.includes(t));
  }
  return targets;
}

function renderConvertPanel() {
  const files = state.files;
  const select = $('targetSelect');
  if (!files.length) {
    select.replaceChildren();
    select.disabled = true;
    $('convertSummary').textContent = '选择一个或多个文件，再选目标格式。批量时只会列出所有文件都支持的目标格式。';
    return '';
  }
  const targets = availableTargets(files);
  const previous = select.value;
  if (!targets.length) {
    select.replaceChildren();
    select.disabled = true;
    const bad = [...new Set(files.map((f) => '.' + ext(f.name)))].join(' ');
    $('convertSummary').textContent = `当前选择的 ${bad} 没有共通的可转换目标格式。可以改用「文件压缩」，或分批选择文件。`;
    return '所选文件没有共通的转换目标格式';
  }
  select.disabled = false;
  select.replaceChildren(...targets.map((t) => el('option', { value: t, text: targetLabel(t) })));
  if (targets.includes(previous)) select.value = previous;
  const names = [...new Set(files.map((f) => '.' + ext(f.name)))].join(' + ');
  $('convertSummary').textContent = files.length === 1
    ? `将 ${names} 转换为所选格式。`
    : `${files.length} 个文件（${names}）将批量转换为同一目标格式，结果自动打包为 ZIP。`;
  return '';
}

/* ---------- 4.2 文件压缩 ---------- */
function estKey() {
  const file = state.files.length === 1 ? state.files[0] : null;
  return file ? `${file.name}:${file.size}:${file.lastModified}` : '';
}

function renderEstimate() {
  const box = $('estimateBox');
  const grid = $('estimateGrid');
  const foot = $('estimateFoot');
  const hint = $('estimateHint');
  const btn = $('estimateBtn');
  const single = state.files.length === 1;

  btn.disabled = !single;
  grid.classList.add('hidden');
  foot.classList.add('hidden');

  if (!single) {
    hint.textContent = state.files.length ? '批量压缩不做预估，直接点下方按钮即可。' : '选择单个文件后可以预估四档压缩后的大小。';
    $('estimateBox').classList.remove('hidden');
    return;
  }

  if (state.est?.loading) {
    hint.textContent = '正在后台试压四档，请稍候…';
    btn.disabled = true;
    return;
  }
  if (state.est?.error) {
    hint.textContent = '预估失败：' + state.est.error + '（可直接压缩，不影响使用）';
    return;
  }
  if (!state.est?.results) {
    hint.textContent = '预估会在后台真正压缩一次，不额外改变最终结果；也可以跳过，直接压缩。';
    return;
  }

  const results = state.est.results;
  const selected = $('levelSelect').value;
  hint.textContent = '预估结果已就绪，点击某一档即可选用该级别。';
  grid.replaceChildren(...COMPRESS_LEVELS.map((level) => {
    const data = results[level];
    const saved = Number(data?.saved_percent ?? 0);
    const card = el('button', {
      class: 'estimate' + (level === selected ? ' active' : ''),
      type: 'button',
      title: `点击选用 ${level}% 档`,
    }, [
      el('span', { class: 'est-lv', text: level + '%' }),
      el('span', { class: 'est-size', text: data ? humanSize(data.output_size) : '—' }),
      el('span', { class: 'est-save' + (saved < 0 ? ' grow' : ''), text: data ? (saved > 0 ? '↓ ' + saved + '%' : '↑ ' + Math.abs(saved) + '%') : '' }),
    ]);
    card.addEventListener('click', () => {
      $('levelSelect').value = level;
      renderEstimate();
    });
    return card;
  }));
  grid.classList.remove('hidden');

  foot.replaceChildren(
    el('span', { text: `原文件 ${humanSize(state.files[0].size)}；选用 ${selected}% 档，压缩后约 ${humanSize(results[selected]?.output_size)}` }),
    el('button', {
      class: 'btn sm', type: 'button', text: '⬇ 直接使用该档结果',
      onclick: () => {
        const task = state.est?.results?.[$('levelSelect').value];
        if (task) {
          clearResults();
          showResults([{ ok: true, label: '压缩', data: task, name: state.files[0].name }], { title: '已使用预估结果（压缩只执行了一次）' });
        }
      },
    }),
  );
  foot.classList.remove('hidden');
}

async function estimateCompress() {
  const file = state.files.length === 1 ? state.files[0] : null;
  if (!file || state.mode !== 'compress') return;
  const key = estKey();
  if (state.est && state.est.key === key && state.est.results) return;

  state.est = { key, loading: true, seq: (state.est?.seq || 0) + 1 };
  const seq = state.est.seq;
  renderEstimate();
  try {
    const results = await Promise.all(COMPRESS_LEVELS.map(async (level) => {
      const fd = new FormData();
      fd.append('file', file);
      fd.append('level', level);
      const data = await request('/api/compress', { body: fd });
      return [level, data];
    }));
    if (state.est?.seq !== seq || estKey() !== key) return;   // 结果已过期
    state.est = { key, seq, results: Object.fromEntries(results) };
  } catch (err) {
    if (state.est?.seq !== seq) return;
    state.est = { key, seq, error: err.message };
  }
  renderEstimate();
}

$('estimateBtn').addEventListener('click', estimateCompress);
$('levelSelect').addEventListener('change', renderEstimate);

/* ---------- 4.3 PDF 转 JPG ---------- */
function renderPdfJpgPanel() {
  const pdfs = state.files.filter((f) => ext(f.name) === 'pdf');
  const summary = $('pdfjpgSummary');
  if (!state.files.length) {
    summary.textContent = '尚未选择文件。可点击上方「选择整个文件夹」一次性导入。';
    return '请先选择包含 PDF 的文件夹';
  }
  if (!pdfs.length) {
    summary.textContent = '所选文件中没有 PDF。这个功能只处理 PDF。';
    return '所选文件中没有 PDF';
  }
  const totalBytes = pdfs.reduce((sum, f) => sum + f.size, 0);
  const skipped = state.files.length - pdfs.length;
  summary.textContent = `已识别 ${pdfs.length} 个 PDF（共 ${humanSize(totalBytes)}）` +
    (skipped ? `，其余 ${skipped} 个非 PDF 文件会被自动忽略。` : '。') +
    ' 全部页面将导出为不超过 300KB 的 JPG 并打包成 ZIP。';
  return '';
}

/* ---------- 4.4 扫描照片 ---------- */
function renderScanPanel() {
  if (!state.files.length) return '请先选择一张照片';
  const file = state.files[0];
  if (state.files.length !== 1) return '扫描照片一次只能处理一张图片';
  if (!SCAN_EXTS.includes(ext(file.name))) return `不支持 .${ext(file.name)} 格式，请上传 JPG / PNG / WEBP 等常见图片`;
  return '';
}

/* ---------- 4.5 PDF 合并 ---------- */
function renderMergePanel() {
  const pdfs = state.files.filter((f) => ext(f.name) === 'pdf');
  const summary = $('mergeSummary');
  if (!state.files.length) {
    summary.textContent = '尚未选择文件。至少需要两个 PDF。';
    return '请选择两个或更多 PDF';
  }
  if (pdfs.length < 2) {
    summary.textContent = `只识别到 ${pdfs.length} 个 PDF，合并至少需要两个。`;
    return 'PDF 合并至少需要两个 PDF';
  }
  const ignored = state.files.length - pdfs.length;
  const modeText = $('mergeMode').selectedOptions[0]?.textContent || '';
  summary.textContent = `已选择 ${pdfs.length} 个 PDF，方式：${modeText}。` +
    (ignored ? ` 另有 ${ignored} 个非 PDF 文件会被忽略。` : '');
  return '';
}

$('mergeMode').addEventListener('change', () => {
  $('insertIndexRow').classList.toggle('hidden', $('mergeMode').value !== 'insert');
  updateModeUI();
});

/* ---------- 4.6 文献翻译 ---------- */
let translateOutput = 'bilingual';

$$('#translateOutputCards label').forEach((label) => {
  label.addEventListener('click', () => {
    $$('#translateOutputCards label').forEach((l) => l.classList.remove('on'));
    label.classList.add('on');
    translateOutput = label.querySelector('input').value;
    updateModeUI();
  });
});

function renderTranslatePanel() {
  if (!state.translateReady) {
    return '后端尚未配置翻译密钥。请在项目根目录创建 .env，填写 TRANSLATE_API_KEY 后重启服务。';
  }
  if (!state.files.length) return '请选择一个英文文献文件（PDF / DOCX / TXT）';
  if (state.files.length !== 1) return '文献翻译一次只能处理一个文件';
  const kind = ext(state.files[0].name);
  if (!TRANSLATE_EXTS.includes(kind)) return `不支持 .${kind} 格式，请上传 PDF / DOCX / TXT`;
  if (translateOutput === 'layout_pdf' && kind !== 'pdf') {
    return '「原版式 PDF」只支持 PDF 输入；DOCX / TXT 请选择 Word 输出形式';
  }
  return '';
}

/* ============================================================
 * 5. 模式与操作分发
 * ============================================================ */
const PANEL_IDS = MODES.map((m) => 'panel-' + m);
const ACTION_LABELS = {
  convert: ['开始转换', '正在转换…'],
  compress: ['开始压缩', '正在压缩…'],
  pdfjpg: ['一键转 JPG（≤300KB/张）', '正在转换…'],
  scan: ['生成扫描件', '正在处理…'],
  pdfmerge: ['开始合并 PDF', '正在合并…'],
  pdfedit: ['保存 PDF 编辑结果', '正在保存…'],
  translate: ['开始翻译文献', '正在翻译…'],
};

function setMode(next, { silent = false } = {}) {
  if (!MODES.includes(next)) return;
  const changed = state.mode !== next;
  state.mode = next;

  $$('.tab').forEach((tab) => {
    const on = tab.dataset.mode === next;
    tab.classList.toggle('active', on);
    tab.setAttribute('aria-selected', String(on));
  });
  PANEL_IDS.forEach((id) => $(id).classList.toggle('hidden', id !== 'panel-' + next));

  if (next === 'pdfedit' && state.files.length === 1 && ext(state.files[0].name) === 'pdf') loadEditor();
  if (next === 'pdfedit') renderEditor();

  // 切换功能不再清空已选文件；只有当当前文件组合对新功能无效时才提示
  updateModeUI();
  renderFileBar();

  if (!silent && changed) { clearResults(); hideActionError(); }
}

function updateModeUI() {
  let problem = '';
  switch (state.mode) {
    case 'convert': problem = renderConvertPanel(); break;
    case 'compress': problem = state.files.length ? '' : '请先选择要压缩的文件'; renderEstimate(); break;
    case 'pdfjpg': problem = renderPdfJpgPanel(); break;
    case 'scan': problem = renderScanPanel(); break;
    case 'pdfmerge': problem = renderMergePanel(); break;
    case 'pdfedit': problem = renderEditPanel(); break;
    case 'translate': problem = renderTranslatePanel(); break;
    default: problem = '';
  }
  if (problem) showActionError(problem);
  else hideActionError();

  const [idle] = ACTION_LABELS[state.mode];
  $('actionBtn').textContent = idle;
  $('actionBtn').disabled = state.busy || !!problem;
  syncGlobalSaveState();
}

function syncGlobalSaveState() {
  const badge = $('globalSaveState');
  if (state.mode === 'pdfedit' && state.editorLoaded && editor.dirty) {
    badge.hidden = false;
    badge.textContent = 'PDF 编辑未保存';
    badge.classList.add('dirty');
  } else {
    badge.hidden = true;
    badge.classList.remove('dirty');
  }
}

async function runAction() {
  if (state.busy) return;
  const files = state.files;
  if (!files.length) return;

  switch (state.mode) {
    case 'pdfjpg': return runPdfFolderToJpg();
    case 'scan': return runScanPhoto();
    case 'pdfmerge': return runPdfMerge();
    case 'pdfedit': return runPdfEdit();
    case 'translate': return runTranslate();
    default: break;
  }

  if (files.length > 1) return runBatch();
  return runSingle();
}

async function runBatch() {
  const mode = state.mode;
  const total = state.files.length;
  beginAction(`正在批量${mode === 'convert' ? '转换' : '压缩'} ${total} 个文件，请耐心等待…`);
  try {
    const fd = new FormData();
    state.files.forEach((f) => fd.append('files', f));
    fd.append('mode', mode);
    if (mode === 'convert') fd.append('target', $('targetSelect').value);
    else fd.append('level', $('levelSelect').value);
    const data = await request('/api/batch', { body: fd });

    const failed = data.failed_files || [];
    const okCount = data.success_files || 0;
    showResults([{ ok: true, label: mode === 'convert' ? '批量转换' : '批量压缩', data, name: data.filename }], {
      kind: failed.length ? 'warn' : 'ok',
      title: failed.length
        ? `${okCount}/${data.total_files} 个文件成功，${failed.length} 个失败`
        : `${okCount} 个文件全部处理完成`,
      details: failed,
    });
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runSingle() {
  const mode = state.mode;
  const file = state.files[0];

  // 压缩：如果已经预估过，直接把那次结果交给用户，避免重复压缩
  if (mode === 'compress' && state.est?.results) {
    const level = $('levelSelect').value;
    const task = state.est.results[level];
    if (task) {
      clearResults();
      showResults([{ ok: true, label: '压缩', data: task, name: file.name }], {
        kind: 'ok',
        title: `压缩完成（${level}% 档，使用已有预估结果，未重复执行）`,
      });
      return;
    }
  }

  beginAction(`正在${mode === 'convert' ? '转换' : '压缩'}，请稍候…`);
  try {
    const fd = new FormData();
    fd.append('file', file);
    const url = mode === 'convert' ? '/api/convert' : '/api/compress';
    if (mode === 'convert') fd.append('target', $('targetSelect').value);
    else fd.append('level', $('levelSelect').value);
    const data = await request(url, { body: fd });
    clearResults();
    showResults([{ ok: true, label: mode === 'convert' ? '转换' : '压缩', data, name: data.filename }], {
      kind: 'ok',
      title: (mode === 'convert' ? '转换' : '压缩') + '完成',
    });
    if (mode === 'compress') state.est = null;   // 源文件已被重新处理，预估结果不再对应
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runPdfFolderToJpg() {
  const pdfs = state.files.filter((f) => ext(f.name) === 'pdf');
  beginAction(`正在处理 ${pdfs.length} 个 PDF，逐页压缩为不超过 300KB 的 JPG…`);
  try {
    const fd = new FormData();
    pdfs.forEach((f) => fd.append('files', f));
    const data = await request('/api/pdf-folder-to-jpg', { body: fd });
    clearResults();
    showResults([{ ok: true, label: 'PDF 转 JPG', data, name: data.filename }], {
      kind: (data.failed_files || []).length ? 'warn' : 'ok',
      title: `已处理 ${data.pdf_files} 个 PDF，生成 ${data.image_files} 张 JPG（每张 ≤300KB）`,
      details: data.failed_files || [],
    });
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runScanPhoto() {
  beginAction('正在识别纸张边缘、校正透视并增强…');
  try {
    const fd = new FormData();
    fd.append('file', state.files[0]);
    const data = await request('/api/scan-photo', { body: fd });
    clearResults();
    showResults([{ ok: true, label: '扫描件', data, name: data.filename }], {
      kind: 'ok',
      title: data.document_detected
        ? '已识别纸张边缘，完成透视矫正与扫描增强'
        : '未找到清晰的纸张边缘，已保留全图并完成扫描增强',
    });
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runPdfMerge() {
  const pdfs = state.files.filter((f) => ext(f.name) === 'pdf');
  const mode = $('mergeMode').value;
  beginAction(`正在合并 ${pdfs.length} 个 PDF…`);
  try {
    const fd = new FormData();
    pdfs.forEach((f) => fd.append('files', f));
    fd.append('mode', mode);
    if (mode === 'insert') {
      fd.append('insert_index', String(INSERT_LABEL_TO_API(Number($('insertIndex').value) || 1)));
    }
    const data = await request('/api/pdf-merge', { body: fd });
    clearResults();
    showResults([{ ok: true, label: 'PDF 合并', data, name: data.filename }], { kind: 'ok', title: 'PDF 合并完成' });
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runPdfEdit() {
  if (!state.editorLoaded) {
    failAction(new Error('预览尚未加载完成，请稍候再试'));
    return;
  }
  beginAction('正在写入编辑结果…');
  try {
    const data = await submitEditor();
    clearResults();
    showResults([{ ok: true, label: 'PDF 编辑', data, name: data.filename }], { kind: 'ok', title: 'PDF 编辑结果已生成' });
    editor.dirty = false;
    renderEditorSaveState();
    endAction();
  } catch (err) {
    failAction(err);
  }
}

async function runTranslate() {
  beginAction('正在翻译，长文献可能需要几分钟，请保持页面打开…');
  $('translateProgress').classList.remove('hidden');
  $('translateStatus').replaceChildren();
  setProgress(0, true, '正在上传文献…');
  try {
    const fd = new FormData();
    fd.append('file', state.files[0]);
    fd.append('output', translateOutput);
    fd.append('domain', $('translateDomain').value);
    fd.append('glossary', $('translateGlossary').value);
    fd.append('translate_names', $('translateNames').value);
    fd.append('skip_references', $('translateSkipRefs').checked ? '1' : '0');

    const start = await request('/api/translate', { body: fd });
    for (;;) {
      await new Promise((r) => setTimeout(r, 1500));
      const info = await request('/api/translate/progress/' + encodeURIComponent(start.task_id), { method: 'GET' });
      if (info.total > 0) {
        setProgress(Math.round((info.done / info.total) * 100), false, info.message || '');
      } else {
        setProgress(100, true, info.message || '正在处理…');
      }
      if (info.state === 'done') {
        setProgress(100, false, '翻译完成');
        const r = info.result || {};
        clearResults();
        showResults([{ ok: true, label: '翻译', data: r, name: r.filename }], {
          kind: 'ok',
          title: `翻译完成，共 ${r.translated || 0} 段` + (r.skipped_references ? '（已跳过参考文献）' : ''),
        });
        break;
      }
      if (info.state === 'error') throw new Error(info.message || '翻译失败');
    }
    endAction();
  } catch (err) {
    $('translateProgress').classList.add('hidden');
    failAction(err);
  }
}

function setProgress(percent, indeterminate, message) {
  const fill = $('translateBarFill');
  fill.classList.toggle('indeterminate', !!indeterminate);
  if (!indeterminate) fill.style.width = Math.max(0, Math.min(100, percent)) + '%';
  $('translateMsg').textContent = message || '';
  $('translatePct').textContent = indeterminate ? '' : percent + '%';
}

async function convertPdfToWord() {
  if (!state.files.length) return;
  beginAction('正在将 PDF 转为可编辑 Word（OCR）…');
  try {
    const fd = new FormData();
    fd.append('file', state.files[0]);
    fd.append('target', 'docx');
    const data = await request('/api/convert', { body: fd });
    clearResults();
    showResults([{ ok: true, label: 'PDF 转 Word', data, name: data.filename }], {
      kind: 'ok',
      title: '已生成可编辑 Word，请下载后在 Word 中修改',
    });
    endAction();
  } catch (err) {
    failAction(err);
  }
}

/* ---------- 忙碌状态与提示 ---------- */
function beginAction(message) {
  state.busy = true;
  const btn = $('actionBtn');
  btn.disabled = true;
  btn.replaceChildren(el('span', { class: 'spinner' }), document.createTextNode(ACTION_LABELS[state.mode][1]));
  showActionRunning(message);
  clearResults();
}

function endAction() {
  state.busy = false;
  updateModeUI();
}

function failAction(err) {
  state.busy = false;
  state.lastError = err?.message || String(err);
  updateModeUI();
  showActionError(state.lastError);
  toast('处理失败：' + state.lastError, 'error', 9000);
}

/* ============================================================
 * 6. 结果与状态展示
 * ============================================================ */
const statusBar = $('statusBar');
const resultGrid = $('resultGrid');

function statusLine({ kind = 'info', title, body, details, actions }) {
  const icons = { info: 'ℹ️', running: '⏳', ok: '✅', warn: '⚠️', error: '❌' };
  const bodyNode = el('div', { class: 'st-body' });
  if (title) bodyNode.append(el('div', { class: 'st-title', text: title }));
  if (body) bodyNode.append(el('div', { text: body }));
  if (details?.length) {
    bodyNode.append(el('ul', { class: 'detail-list' }, details.map((d) => el('li', { text: d }))));
  }
  if (actions?.length) {
    bodyNode.append(el('div', { class: 'btn-row' }, actions));
  }
  return el('div', { class: 'status-line ' + kind }, [
    el('span', { class: 'st-icon', text: icons[kind] || icons.info }),
    bodyNode,
  ]);
}

function showActionRunning(message) {
  statusBar.replaceChildren(statusLine({ kind: 'running', body: message }));
}
function showActionError(message) {
  statusBar.replaceChildren(statusLine({ kind: 'error', title: '无法继续', body: message }));
}
function showActionInfo(message) {
  statusBar.replaceChildren(statusLine({ kind: 'info', body: message }));
}
function hideActionError() {
  const current = statusBar.firstElementChild;
  if (current?.classList.contains('error') || current?.classList.contains('running')) statusBar.replaceChildren();
}

function clearResults() {
  resultGrid.replaceChildren();
}

function resultCard({ ok, label, data, name }) {
  const card = el('div', { class: 'result-card ' + (ok ? 'ok' : 'fail') });
  card.append(el('div', { class: 'rc-head' }, [
    el('span', { text: ok ? '✅' : '❌' }),
    el('span', { text: ok ? `${label}完成` : `${label}结果` }),
  ]));
  card.append(el('p', { class: 'rc-file', text: name || data.filename || '' }));

  if (ok && data) {
    const original = Number(data.original_size || 0);
    const output = Number(data.output_size || 0);
    const saved = Number(data.saved_percent ?? 0);
    const stats = [];
    if (original) {
      stats.push(el('div', { class: 'rc-stat' }, [el('span', { text: '原始大小' }), el('b', { text: humanSize(original) })]));
      stats.push(el('span', { class: 'rc-arrow', text: '→' }));
    }
    stats.push(el('div', { class: 'rc-stat after' }, [el('span', { text: '输出大小' }), el('b', { text: humanSize(output) })]));
    if (original && saved !== 0) {
      stats.push(el('div', { class: 'rc-stat saved' }, [
        el('span', { text: saved > 0 ? '节省' : '增大' }),
        el('b', { text: (saved > 0 ? '↓ ' : '↑ ') + Math.abs(saved) + '%' }),
      ]));
    }
    card.append(el('div', { class: 'rc-stats' }, stats));

    const meta = [];
    if (data.output_format) meta.push('格式 .' + data.output_format);
    if (data.total_files) meta.push(`共 ${data.total_files} 个文件，成功 ${data.success_files} 个`);
    if (data.pdf_files) meta.push(`${data.pdf_files} 个 PDF → ${data.image_files} 张 JPG`);
    if (data.paragraphs) meta.push(`共 ${data.paragraphs} 段，翻译 ${data.translated} 段`);
    if (data.dropped) meta.push(`因版面冲突保留原文 ${data.dropped} 段`);
    if (data.truncated) meta.push('预览只包含前 50 页');
    if (meta.length) card.append(el('div', { class: 'rc-meta', text: meta.join(' · ') }));

    if (data.note) card.append(el('div', { class: 'rc-note', text: 'ⓘ ' + data.note }));

    const actions = [];
    if (data.download_url) {
      actions.push(el('a', {
        class: 'download-btn', href: data.download_url, download: data.filename || '',
        text: '⬇ 下载 ' + (data.filename || '结果文件'),
      }));
    }
    if (data.failed_files?.length) {
      actions.push(el('button', {
        class: 'btn sm', type: 'button', text: '📋 复制失败清单',
        onclick: async () => {
          const okCopy = await copyText(data.failed_files.join('\n'));
          toast(okCopy ? '失败清单已复制到剪贴板' : '复制失败，请手动选择文本', okCopy ? 'ok' : 'warn');
        },
      }));
    }
    if (data.task_id) {
      actions.push(el('button', {
        class: 'btn sm', type: 'button', text: '🔗 复制下载链接',
        onclick: async () => {
          const url = new URL(data.download_url, location.origin).href;
          const okCopy = await copyText(url);
          toast(okCopy ? '下载链接已复制' : '复制失败', okCopy ? 'ok' : 'warn');
        },
      }));
    }
    card.append(el('div', { class: 'rc-actions' }, actions));
  }
  return card;
}

/**
 * @param {Array<{ok:boolean,label:string,data:any,name?:string}>} cards
 * @param {{kind?:string,title?:string,body?:string,details?:string[]}} summary
 */
function showResults(cards, summary = {}) {
  const actions = [];
  const downloadables = cards.filter((c) => c.ok && c.data?.download_url);
  if (downloadables.length > 1) {
    actions.push(el('button', {
      class: 'btn sm', type: 'button', text: '⬇ 依次下载全部',
      onclick: () => downloadables.forEach((c, i) => setTimeout(() => download(c.data.download_url, c.data.filename), i * 400)),
    }));
  }
  if (state.lastError) {
    actions.push(el('button', {
      class: 'btn sm', type: 'button', text: '📋 复制错误信息',
      onclick: async () => {
        const okCopy = await copyText(state.lastError);
        toast(okCopy ? '错误信息已复制' : '复制失败', okCopy ? 'ok' : 'warn');
      },
    }));
  }
  if (summary.title || summary.body || summary.details?.length) {
    statusBar.replaceChildren(statusLine({ ...summary, kind: summary.kind || 'ok', actions }));
  }
  resultGrid.replaceChildren(...cards.map(resultCard));
  resultGrid.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/* ============================================================
 * 7. PDF 可视化编辑器
 * ============================================================ */
const editor = {
  pages: [],            // { id, original, rotation, deleted, width, height, raw, image, texts }
  annotations: [],
  tool: 'select',
  current: 0,
  zoom: 1,
  dirty: false,
  history: [],
  historyIndex: -1,
  drawing: null,
  editing: null,
  nextId: 1,
  objectUrl: null,
};

function resetEditor() {
  editor.pages = [];
  editor.annotations = [];
  editor.history = [];
  editor.historyIndex = -1;
  editor.current = 0;
  editor.tool = 'select';
  editor.zoom = 1;
  editor.dirty = false;
  editor.drawing = null;
  editor.editing = null;
  editor.objectUrl = null;
  editor.selectedId = null;
  editor.deletedOriginal = new Set();
  editor.manualRotations = new Map();
  editor.totalPages = 0;
  editor.loading = false;
  state.editorLoaded = false;
  $('deletePages').value = '';
  $('rotatePage').value = '';
  $('pageOrder').value = '';
  $('pdfStage').replaceChildren();
  $('thumbList').replaceChildren();
  renderEditor();
}

function renderEditorSaveState() {
  const text = editor.dirty ? '未保存' : (state.editorLoaded ? '已同步' : '未保存');
  $('editorSaveState').textContent = text;
  $('editorSaveState').className = 'chip' + (editor.dirty ? ' warn' : '');
  syncGlobalSaveState();
}

function renderEditPanel() {
  if (!state.files.length) {
    $('editor').classList.add('hidden');
    $('editorPlaceholder').classList.remove('hidden');
    $('editorPlaceholder').textContent = '选择一个 PDF 后，这里会出现可视化编辑器。';
    return '请选择一个 PDF 文件';
  }
  if (state.files.length !== 1) {
    $('editor').classList.add('hidden');
    $('editorPlaceholder').classList.remove('hidden');
    return 'PDF 编辑一次只能处理一个文件';
  }
  if (ext(state.files[0].name) !== 'pdf') {
    $('editor').classList.add('hidden');
    $('editorPlaceholder').classList.remove('hidden');
    return 'PDF 编辑只支持 PDF 文件';
  }
  if (!state.editorLoaded) {
    $('editor').classList.add('hidden');
    $('editorPlaceholder').classList.remove('hidden');
    return '正在加载预览…';
  }
  $('editorPlaceholder').classList.add('hidden');
  $('editor').classList.remove('hidden');
  return '';
}

async function loadEditor() {
  const file = state.files[0];
  if (!file || ext(file.name) !== 'pdf') return;
  if (state.editorLoaded || editor.loading) return;
  editor.loading = true;
  renderEditor();
  showActionRunning('正在生成页面预览（最多 50 页）…');
  try {
    const fd = new FormData();
    fd.append('file', file);
    const data = await request('/api/pdf-preview', { body: fd });
    editor.totalPages = data.page_count || (data.pages || []).length;
    editor.pages = (data.pages || []).map((page, index) => ({
      id: 'page-' + index,
      original: index + 1,        // 在后端 edit_pdf 中使用的原始 1-based 页码
      rotation: 0,
      deleted: false,
      width: page.width,
      height: page.height,
      raw: { w: page.width, h: page.height },   // 未旋转的 PDF 尺寸，旋转后仍按它换算坐标
      image: page.image,
      rotated: null,
      texts: (page.texts || []).map((t) => ({ text: t.text, rect: t.rect, size: t.size })),
    }));
    editor.annotations = [];
    editor.current = 0;
    editor.dirty = false;
    editor.nextId = 1;
    state.editorLoaded = true;
    pushHistory(true);
    hideActionError();
    renderEditor();
    if (editor.totalPages > editor.pages.length) {
      toast(`文档共 ${editor.totalPages} 页，仅前 ${editor.pages.length} 页可预览编辑；保存时其余页面会原样保留在结果文件末尾。`, 'warn', 9000);
    }
  } catch (err) {
    state.editorLoaded = false;
    showActionError('预览加载失败：' + err.message);
    toast('PDF 预览失败：' + err.message, 'error', 8000);
  } finally {
    editor.loading = false;
    updateModeUI();
  }
}

/* ---------- 历史记录（整篇可撤销） ---------- */
function snapshot() {
  return {
    annotations: JSON.parse(JSON.stringify(editor.annotations)),
    current: editor.current,
    pages: editor.pages.map((p) => ({ id: p.id, original: p.original, rotation: p.rotation, deleted: !!p.deleted })),
    // 手填的页码输入框也要一起进历史，否则撤销后输入框和画面会对不上
    deletedOriginal: [...editor.deletedOriginal],
    manualRotations: [...editor.manualRotations.entries()],
  };
}

function pushHistory(reset = false) {
  if (reset) {
    editor.history = [];
    editor.historyIndex = -1;
  }
  const snap = JSON.stringify(snapshot());
  if (editor.history[editor.historyIndex] === snap) return;
  editor.history = editor.history.slice(0, editor.historyIndex + 1);
  editor.history.push(snap);
  if (editor.history.length > 100) editor.history.shift();
  editor.historyIndex = editor.history.length - 1;
}

function commit({ rebuild = false } = {}) {
  editor.dirty = true;
  pushHistory();
  syncPageEditFields();
  renderEditorSaveState();
  if (rebuild) renderEditor();
  else redrawAllCanvases();
}

function restoreHistory(delta) {
  const next = editor.historyIndex + delta;
  if (next < 0 || next >= editor.history.length) return;
  editor.historyIndex = next;
  const snap = JSON.parse(editor.history[next]);
  editor.annotations = snap.annotations;
  editor.current = Math.min(snap.current, snap.pages.length - 1);
  editor.deletedOriginal = new Set(snap.deletedOriginal || []);
  editor.manualRotations = new Map(snap.manualRotations || []);
  // 页面结构变化时需要重建 DOM，否则只重绘画布即可
  const structureChanged = snap.pages.length !== editor.pages.length ||
    snap.pages.some((p, i) => p.id !== editor.pages[i]?.id || p.rotation !== editor.pages[i]?.rotation);
  snap.pages.forEach((p, i) => {
    if (!editor.pages[i]) return;
    editor.pages[i].id = p.id;
    editor.pages[i].original = p.original;
    editor.pages[i].rotation = p.rotation;
    editor.pages[i].deleted = !!p.deleted;
  });
  editor.dirty = true;
  if (structureChanged) renderEditor();
  else redrawAllCanvases();
  renderPropertyPanel();
  renderEditorSaveState();
}

/* ---------- 渲染 ---------- */
function redrawAllCanvases() {
  $$('#pdfStage .pdf-page').forEach((box) => {
    const canvas = box.querySelector('canvas');
    if (canvas) drawCanvas(Number(box.dataset.page), canvas);
  });
}

/** 旋转后画布位图的尺寸（预览数据以未旋转的 PDF 为准，这里自行按角度换算）。 */
const rotatedDims = displayDims;

/** 页面旋转 90/270 时文本层要整体旋转映射，这里直接重新计算每个文字块的位置。 */
function rotateTextRects(page) {
  return (page.texts || []).map((item) => ({ ...item, rect: rectToDisplay(page, item.rect) }));
}

/** 生成（并缓存）旋转后的图片；旋转通过画布完成，避免 CSS transform 与坐标换算冲突。 */
function rotatedSrc(page) {
  if (!page.rotation) return page.image;
  const cached = page.rotated?.[page.rotation];
  if (cached && cached !== 'pending') return cached;
  const source = page.image;
  const quarter = ((page.rotation % 360) + 360) % 360;
  const raw = page.raw || { w: page.width, h: page.height };
  const build = (img) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (quarter === 180) { canvas.width = raw.w; canvas.height = raw.h; ctx.translate(raw.w, raw.h); ctx.rotate(Math.PI); }
    else if (quarter === 90) { canvas.width = raw.h; canvas.height = raw.w; ctx.translate(raw.h, 0); ctx.rotate(Math.PI / 2); }
    else if (quarter === 270) { canvas.width = raw.h; canvas.height = raw.w; ctx.translate(0, raw.w); ctx.rotate(-Math.PI / 2); }
    else return source;
    ctx.drawImage(img, 0, 0, raw.w, raw.h);
    page.rotated = page.rotated || {};
    page.rotated[page.rotation] = canvas.toDataURL('image/jpeg', 0.9);
    return page.rotated[page.rotation];
  };
  const ready = page.loadedImage;
  if (ready?.naturalWidth) return build(ready);
  page.rotated = page.rotated || {};
  page.rotated[page.rotation] = 'pending';
  const img = new Image();
  img.onload = () => {
    build(img);
    if (state.editorLoaded) renderEditor();   // 旋转图就绪后重绘
  };
  img.src = source;
  return source;   // 首次渲染先显示未旋转图，加载完成后自动替换
}

function renderEditor() {
  if (!state.editorLoaded) {
    renderEditorSaveState();
    return;
  }
  const stage = $('pdfStage');
  const oldScroll = $('stageScroll').scrollTop;
  stage.replaceChildren();

  editor.pages.forEach((page, index) => {
    const dims = rotatedDims(page);
    const baseImage = rotatedSrc(page);
    const box = el('div', {
      class: 'pdf-page' + (editor.tool === 'select' ? ' tool-select' : ''),
      dataset: { page: String(index) },
      style: `width:${Math.round(Math.min(dims.w * 1.25, 760) * editor.zoom)}px`,
    });
    const img = el('img', { src: baseImage, alt: `第 ${index + 1} 页`, draggable: 'false' });
    page.loadedImage = img;
    const canvas = el('canvas', { width: '1000', height: '1400' });
    bindCanvas(canvas, index);
    const layer = el('div', { class: 'text-layer' });
    rotateTextRects(page).forEach((item, textIndex) => {
      const node = el('div', {
        class: 'text-item',
        text: item.text,
        dataset: { textIndex: String(textIndex) },
        style: [
          `left:${(item.rect[0] / dims.w) * 100}%`,
          `top:${(item.rect[1] / dims.h) * 100}%`,
          `min-height:${Math.max(1, ((item.rect[3] - item.rect[1]) / dims.h) * 100)}%`,
          `font-size:${(item.size / (page.raw?.h || page.height)) * 100}%`,
          'line-height:1.15',
        ].join(';'),
      });
      node.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        startTextEdit(node, page.texts[textIndex], index);
      });
      layer.append(node);
    });
    box.append(img, canvas, layer);
    stage.append(box);

    const syncCanvas = () => {
      canvas.width = dims.w;
      canvas.height = dims.h;
      drawCanvas(index, canvas);
    };
    if (img.complete && img.naturalWidth) syncCanvas();
    else img.addEventListener('load', syncCanvas, { once: true });
  });

  renderThumbs();
  $('pageCountLabel').textContent = `${editor.pages.length} 页`;
  $('pageJump').value = String(editor.current + 1);
  $('zoomLabel').textContent = Math.round(editor.zoom * 100) + '%';
  renderPropertyPanel();
  renderEditorSaveState();
  // 没有文字层（扫描件）才需要提示转为可编辑 Word
  $('convertPdfWord').classList.toggle('hidden', editor.pages.some((p) => (p.texts || []).length));
  $('stageScroll').scrollTop = oldScroll;
  syncPageEditFields();
}

function renderThumbs() {
  const list = $('thumbList');
  list.replaceChildren(...editor.pages.map((page, index) => {
    const thumb = el('div', {
      class: 'thumb' + (index === editor.current ? ' active' : ''),
      draggable: 'true',
      dataset: { page: String(index) },
      title: `第 ${index + 1} 页（原第 ${page.original} 页）`,
    }, [
      el('img', { src: rotatedSrc(page), alt: '', draggable: 'false' }),
      el('small', { text: `第 ${index + 1} 页` }),
    ]);
    if (page.rotation) thumb.append(el('span', { class: 'th-badge', text: page.rotation + '°' }));
    thumb.addEventListener('click', () => goToPage(index));
    thumb.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', String(index));
      e.dataTransfer.effectAllowed = 'move';
      thumb.classList.add('dragging');
    });
    thumb.addEventListener('dragend', () => thumb.classList.remove('dragging'));
    thumb.addEventListener('dragover', (e) => { e.preventDefault(); thumb.classList.add('drop-target'); });
    thumb.addEventListener('dragleave', () => thumb.classList.remove('drop-target'));
    thumb.addEventListener('drop', (e) => {
      e.preventDefault();
      thumb.classList.remove('drop-target');
      const from = Number(e.dataTransfer.getData('text/plain'));
      if (!Number.isFinite(from) || from === index) return;
      const [moved] = editor.pages.splice(from, 1);
      editor.pages.splice(index, 0, moved);
      remapAnnotationsAfterReorder(from, index);
      commit({ rebuild: true });
      goToPage(index);
    });
    return thumb;
  }));
}

/** 页面顺序变化时，把标注的 page 索引一起搬过去，避免标注跟着错位。 */
function remapAnnotationsAfterReorder(from, to) {
  const map = (p) => {
    if (p === from) return to;
    if (from < to && p > from && p <= to) return p - 1;
    if (from > to && p >= to && p < from) return p + 1;
    return p;
  };
  editor.annotations.forEach((a) => { a.page = map(a.page); });
}

function goToPage(index) {
  editor.current = Math.max(0, Math.min(index, editor.pages.length - 1));
  $('pageJump').value = String(editor.current + 1);
  const target = $(`#pdfStage .pdf-page[data-page="${editor.current}"]`);
  target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  $$('#thumbList .thumb').forEach((t, i) => t.classList.toggle('active', i === editor.current));
}

/* ---------- 画布绘制 ----------
 * 画布位图与「旋转后的显示页面」一一对应，所有绘制先把 PDF 坐标换算到显示坐标。
 */
function drawCanvas(pageIndex, canvas) {
  const page = editor.pages[pageIndex];
  if (!page) return;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const dims = displayDims(page);
  const sx = canvas.width / dims.w;
  const sy = canvas.height / dims.h;

  const stroke = (color, width) => {
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = Math.max(1, (width || 2) * ((sx + sy) / 2));
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
  };
  const toHex = (rgb, alpha = 1) => {
    const [r, g, b] = (rgb || [0, 0, 0]).map((v) => Math.max(0, Math.min(255, Math.round(Number(v) <= 1 ? Number(v) * 255 : Number(v)))));
    return `rgba(${r},${g},${b},${alpha})`;
  };
  const font = `${getComputedStyle(document.body).fontFamily}`;
  const disp = (x, y) => { const [dx, dy] = pointToDisplay(page, x, y); return [dx * sx, dy * sy]; };

  editor.annotations.filter((a) => a.page === pageIndex).forEach((a) => {
    const rect = a.rect ? rectToDisplay(page, a.rect) : null;
    const box = rect ? [rect[0] * sx, rect[1] * sy, (rect[2] - rect[0]) * sx, (rect[3] - rect[1]) * sy] : null;
    switch (a.type) {
      case 'replace_text':
        ctx.fillStyle = '#fff';
        ctx.fillRect(box[0], box[1], box[2], box[3]);
        ctx.fillStyle = toHex(a.color || [0, 0, 0]);
        ctx.font = `${Math.max(6, a.size) * sy}px ${font}`;
        ctx.textBaseline = 'top';
        ctx.fillText(a.text, box[0], box[1]);
        break;
      case 'text':
        ctx.fillStyle = toHex(a.color || [0, 0, 0]);
        ctx.font = `${Math.max(6, a.size) * sy}px ${font}`;
        ctx.textBaseline = 'top';
        ctx.fillText(a.text, box[0], box[1]);
        break;
      case 'rect':
        stroke(toHex(a.color || [0.1, 0.35, 0.85], a.opacity ?? 1), a.width);
        if (a.fill) {
          ctx.fillStyle = toHex(a.fill, a.opacity ?? 1);
          ctx.fillRect(box[0], box[1], box[2], box[3]);
        }
        ctx.strokeRect(box[0], box[1], box[2], box[3]);
        break;
      case 'ellipse':
        stroke(toHex(a.color || [0.1, 0.35, 0.85], a.opacity ?? 1), a.width);
        ctx.beginPath();
        ctx.ellipse(box[0] + box[2] / 2, box[1] + box[3] / 2, box[2] / 2, box[3] / 2, 0, 0, Math.PI * 2);
        if (a.fill) { ctx.fillStyle = toHex(a.fill, a.opacity ?? 1); ctx.fill(); }
        ctx.stroke();
        break;
      case 'highlight':
        ctx.fillStyle = toHex(a.color || [1, 0.75, 0.1], a.opacity ?? 0.35);
        ctx.fillRect(box[0], box[1], box[2], box[3]);
        break;
      case 'underline':
      case 'strike': {
        stroke(toHex(a.color || [0.9, 0.6, 0]), a.width);
        const y = a.type === 'underline' ? box[1] + box[3] - 2 : box[1] + box[3] / 2;
        ctx.beginPath();
        ctx.moveTo(box[0], y);
        ctx.lineTo(box[0] + box[2], y);
        ctx.stroke();
        break;
      }
      case 'line':
      case 'arrow': {
        const [x0, y0] = disp(a.start[0], a.start[1]);
        const [x1, y1] = disp(a.end[0], a.end[1]);
        stroke(toHex(a.color || [0.05, 0.15, 0.65]), a.width);
        ctx.beginPath();
        ctx.moveTo(x0, y0);
        ctx.lineTo(x1, y1);
        ctx.stroke();
        if (a.type === 'arrow') {
          const angle = Math.atan2(y1 - y0, x1 - x0);
          const head = 11 * ((sx + sy) / 2);
          ctx.beginPath();
          ctx.moveTo(x1, y1);
          ctx.lineTo(x1 - head * Math.cos(angle - 0.4), y1 - head * Math.sin(angle - 0.4));
          ctx.lineTo(x1 - head * Math.cos(angle + 0.4), y1 - head * Math.sin(angle + 0.4));
          ctx.closePath();
          ctx.fill();
        }
        break;
      }
      case 'signature':
        stroke('rgba(13,38,166,1)', 1.6);
        (a.strokes || []).forEach((points) => {
          ctx.beginPath();
          points.forEach(([x, y], i) => {
            const [dx, dy] = disp(x, y);
            return i ? ctx.lineTo(dx, dy) : ctx.moveTo(dx, dy);
          });
          ctx.stroke();
        });
        break;
      default:
        break;
    }
  });
}

/* ---------- 交互：新增标注 ---------- */
const RECT_TOOLS = ['text', 'cover', 'highlight', 'underline', 'strike', 'rect', 'ellipse'];
const DRAG_TOOLS = ['line', 'arrow', 'sign'];

function canvasPoint(canvas, event) {
  const box = canvas.getBoundingClientRect();
  const page = editor.pages[Number(canvas.parentElement.dataset.page)];
  const dims = displayDims(page);
  // 画布显示的是旋转后的页面，先换算到显示坐标，再转回 PDF 坐标存储
  const dx = ((event.clientX - box.left) / box.width) * dims.w;
  const dy = ((event.clientY - box.top) / box.height) * dims.h;
  const [x, y] = pointToPdf(page, dx, dy);
  return [Math.round(x * 100) / 100, Math.round(y * 100) / 100];
}

/** 在「PDF 坐标」下按旋转后的可视范围生成默认矩形。 */
function defaultRect(page, tool, x, y) {
  const dims = displayDims(page);
  const isTextLine = ['highlight', 'underline', 'strike'].includes(tool);
  const w = isTextLine ? 150 : 120;
  const h = isTextLine ? Math.max(11, Number($('editSize').value) || 14) : 30;
  const dw = Math.max(1, w - 2);
  const dh = (tool === 'underline' || tool === 'strike') ? Math.max(2, h - 2) : Math.max(3, h);
  // 在显示坐标里画一个框，再整体转回 PDF 坐标（旋转时包围盒会随之互换宽高）
  const [px0, py0] = pointToPdf(page, x, y);
  const [px1, py1] = pointToPdf(page, Math.min(x + dw, dims.w), Math.min(y + dh, dims.h));
  return [Math.min(px0, px1), Math.min(py0, py1), Math.max(px0, px1), Math.max(py0, py1)];
}

function createId() {
  return editor.nextId++;
}

function bindCanvas(canvas, pageIndex) {
  const page = editor.pages[pageIndex];

  canvas.addEventListener('click', (event) => {
    if (!RECT_TOOLS.includes(editor.tool)) return;
    const [x, y] = canvasPoint(canvas, event);
    if (editor.tool === 'text') {
      const text = $('editText').value.trim();
      if (!text) {
        toast('请先在工具栏填写「要添加的文字内容」', 'warn');
        return;
      }
      const size = Math.max(6, Number($('editSize').value) || 14);
      const dims = displayDims(page);
      // 文字框宽高在显示坐标下估算，再换算回 PDF 坐标，保证旋转页面上位置一致
      const base = defaultRect(page, 'text', x, y);
      const shown = rectToDisplay(page, base);
      const boxW = Math.max(shown[2] - shown[0], size * text.length * 0.5);
      const boxH = Math.max(shown[3] - shown[1], size * 1.3);
      const corner = pointToPdf(page, Math.min(shown[0] + boxW, dims.w), Math.min(shown[1] + boxH, dims.h));
      editor.annotations.push({
        id: createId(), page: pageIndex, type: 'text', text, size,
        rect: [Math.min(base[0], corner[0]), Math.min(base[1], corner[1]),
               Math.max(base[0], corner[0]), Math.max(base[1], corner[1])],
        color: [0, 0, 0], opacity: 1, width: 1.5,
      });
    } else {
      const type = editor.tool === 'cover' ? 'rect' : editor.tool;
      const item = {
        id: createId(), page: pageIndex, type,
        rect: defaultRect(page, type, x, y),
        color: type === 'highlight' ? [1, 0.75, 0.1] : [0.1, 0.35, 0.85],
        opacity: type === 'highlight' ? 0.35 : 1,
        width: 1.5,
      };
      if (editor.tool === 'cover') { item.color = [1, 1, 1]; item.fill = [1, 1, 1]; }
      editor.annotations.push(item);
    }
    selectAnnotation(editor.annotations.at(-1)?.id);
    commit();
  });

  canvas.addEventListener('pointerdown', (event) => {
    if (!DRAG_TOOLS.includes(editor.tool)) return;
    event.preventDefault();
    canvas.setPointerCapture(event.pointerId);
    editor.drawing = { page: pageIndex, canvas, points: [canvasPoint(canvas, event)] };
  });

  canvas.addEventListener('pointermove', (event) => {
    if (!editor.drawing || editor.drawing.canvas !== canvas) return;
    editor.drawing.points.push(canvasPoint(canvas, event));
    drawCanvas(pageIndex, canvas);
    const ctx = canvas.getContext('2d');
    ctx.strokeStyle = '#102a9b';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.beginPath();
    editor.drawing.points.forEach(([x, y], i) => {
      const [dx, dy] = pointToDisplay(page, x, y);
      return i ? ctx.lineTo(dx, dy) : ctx.moveTo(dx, dy);
    });
    ctx.stroke();
  });

  const finish = (event) => {
    if (!editor.drawing || editor.drawing.canvas !== canvas) return;
    const points = editor.drawing.points;
    editor.drawing = null;
    if (points.length < 2) { drawCanvas(pageIndex, canvas); return; }
    const type = editor.tool === 'sign' ? 'signature' : editor.tool;
    const item = { id: createId(), page: pageIndex, type, color: [0.05, 0.15, 0.65], width: 2.2 };
    if (type === 'signature') item.strokes = [points];
    else { item.start = points[0]; item.end = points.at(-1); }
    editor.annotations.push(item);
    selectAnnotation(item.id);
    commit();
  };
  canvas.addEventListener('pointerup', finish);
  canvas.addEventListener('pointercancel', () => { if (editor.drawing?.canvas === canvas) { editor.drawing = null; drawCanvas(pageIndex, canvas); } });
  canvas.addEventListener('pointerleave', (event) => { if (event.buttons === 0) return; finish(event); });
}

/* ---------- 选中、属性面板、对象操作 ---------- */
function selectAnnotation(id) {
  editor.selectedId = id;
  renderPropertyPanel();
}

function selectedAnnotation() {
  return editor.annotations.find((a) => a.id === editor.selectedId) || null;
}

function annotationBounds(item) {
  if (item.rect) return item.rect.slice();
  const xs = [item.start?.[0] ?? 0, item.end?.[0] ?? 0];
  const ys = [item.start?.[1] ?? 0, item.end?.[1] ?? 0];
  (item.strokes || []).flat().forEach(([x, y]) => { xs.push(x); ys.push(y); });
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

function renderPropertyPanel() {
  const panel = $('propertyPanel');
  const item = selectedAnnotation();
  if (!item) {
    panel.className = 'empty-props';
    panel.replaceChildren(
      el('div', { html: `<b>当前文件</b><br>${state.files[0]?.name || '-'}<br><b>页数</b> ${editor.pages.length} 页<br><b>标注</b> ${editor.annotations.length} 个` }),
      el('div', { style: 'margin-top:10px', text: editor.pages.some((p) => (p.texts || []).length)
        ? '双击页面上的原文字可以直接改写；用下方工具添加标注。'
        : '这份 PDF 没有文字层（可能是扫描件），请先转为可编辑 Word。' }),
    );
    return;
  }
  const names = { replace_text: '文字替换', text: '文字', rect: '矩形 / 遮盖', ellipse: '圆形',
    highlight: '高亮', underline: '下划线', strike: '删除线', line: '线条', arrow: '箭头', signature: '签名' };
  const bounds = annotationBounds(item);
  panel.className = '';
  panel.replaceChildren();

  panel.append(el('div', { class: 'chip fmt', style: 'margin-bottom:9px;display:inline-block', text: names[item.type] || item.type }));
  panel.append(el('div', { class: 'hint', style: 'margin:0 0 9px', text: `第 ${item.page + 1} 页` }));

  const grid = el('div', { class: 'prop-grid' });
  const num = (label, value, onChange, step = 1) => {
    const input = el('input', { type: 'number', value: String(Math.round(value)), step: String(step) });
    input.addEventListener('change', () => { onChange(Number(input.value) || 0); commit(); });
    return el('label', {}, [document.createTextNode(label), input]);
  };

  // 位置（矩形类对象可移动，线/签名只显示）
  const movable = ['replace_text', 'text', 'rect', 'ellipse', 'highlight', 'underline', 'strike'].includes(item.type);
  if (movable) {
    const w = bounds[2] - bounds[0];
    const h = bounds[3] - bounds[1];
    grid.append(num('X', bounds[0], (v) => { item.rect = [v, item.rect[1], v + w, item.rect[3]]; }));
    grid.append(num('Y', bounds[1], (v) => { item.rect = [item.rect[0], v, item.rect[2], v + h]; }));
    grid.append(num('宽', w, (v) => { item.rect = [item.rect[0], item.rect[1], item.rect[0] + Math.max(2, v), item.rect[3]]; }));
    grid.append(num('高', h, (v) => { item.rect = [item.rect[0], item.rect[1], item.rect[2], item.rect[1] + Math.max(2, v)]; }));
  } else {
    grid.append(el('div', { class: 'full hint', style: 'margin:0', text: `起点 (${Math.round(bounds[0])}, ${Math.round(bounds[1])}) → 终点 (${Math.round(bounds[2])}, ${Math.round(bounds[3])})` }));
  }

  if (item.type === 'text' || item.type === 'replace_text') {
    const text = el('input', { type: 'text', value: item.text, class: 'full' });
    text.addEventListener('change', () => { item.text = text.value; commit(); });
    grid.append(el('label', { class: 'full' }, [document.createTextNode('文字内容'), text]));
    grid.append(num('字号', item.size || 14, (v) => { item.size = Math.max(5, v); }, 0.5));
  }

  if (['rect', 'ellipse'].includes(item.type)) {
    const opacity = el('input', { type: 'range', min: '0.1', max: '1', step: '0.05', value: String(item.opacity ?? 1) });
    opacity.addEventListener('input', () => { item.opacity = Number(opacity.value); commit(); });
    grid.append(el('label', { class: 'full' }, [document.createTextNode('不透明度'), opacity]));
  }
  if (['line', 'arrow', 'rect', 'ellipse', 'underline', 'strike'].includes(item.type)) {
    grid.append(num('线宽', item.width ?? 2, (v) => { item.width = Math.max(0.5, v); }, 0.5));
  }
  panel.append(grid);

  const actions = el('div', { class: 'prop-actions' });
  actions.append(el('button', {
    class: 'btn sm', type: 'button', text: '复制到下一页',
    onclick: () => {
      if (item.page + 1 >= editor.pages.length) { toast('已经是最后一页', 'warn'); return; }
      const copy = JSON.parse(JSON.stringify(item));
      copy.id = createId();
      copy.page = item.page + 1;
      editor.annotations.push(copy);
      editor.current = copy.page;
      selectAnnotation(copy.id);
      commit({ rebuild: true });
      goToPage(copy.page);
    },
  }));
  if (item.type === 'signature' && item.page + 1 < editor.pages.length) {
    // 多页签名：一次性把同一签名盖到后面所有页
    actions.append(el('button', {
      class: 'btn sm', type: 'button', text: '盖到后面所有页',
      onclick: () => {
        for (let p = item.page + 1; p < editor.pages.length; p++) {
          const copy = JSON.parse(JSON.stringify(item));
          copy.id = createId();
          copy.page = p;
          editor.annotations.push(copy);
        }
        commit({ rebuild: true });
        toast('签名已复制到后面所有页', 'ok');
      },
    }));
  }
  actions.append(el('button', {
    class: 'btn sm', type: 'button', text: '删除该对象',
    onclick: () => {
      editor.annotations = editor.annotations.filter((a) => a !== item);
      editor.selectedId = null;
      commit({ rebuild: true });
    },
  }));
  panel.append(actions);
}

/* ---------- 原文字改写 ---------- */
function startTextEdit(node, item, pageIndex) {
  if (editor.editing) finishTextEdit(editor.editing.node, editor.editing.item, editor.editing.page);
  editor.editing = { node, item, page: pageIndex };
  node.contentEditable = 'true';
  node.classList.add('editing');
  node.focus();
  const range = document.createRange();
  range.selectNodeContents(node);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  const onKey = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); node.blur(); }
    if (e.key === 'Escape') { node.textContent = item.text; node.blur(); }
  };
  node.addEventListener('keydown', onKey);
  node.addEventListener('blur', () => {
    node.removeEventListener('keydown', onKey);
    finishTextEdit(node, item, pageIndex);
  }, { once: true });
}

function finishTextEdit(node, item, pageIndex) {
  node.contentEditable = 'false';
  node.classList.remove('editing');
  const text = node.textContent.replace(/\s*\n\s*/g, ' ').trim();
  editor.editing = null;
  if (!text || text === item.text) { node.textContent = item.text; return; }

  const page = editor.pages[pageIndex];
  const rect = item.rect;
  // 预览图是 PDF 坐标的 1.2 倍，字号要按同一比例缩小后才能与原文对齐
  const size = Math.max(5, Math.round(item.size * 1.2) / 1.2);
  const height = Math.max((rect[3] - rect[1]) * 1.15, size * 1.25);
  editor.annotations.push({
    id: createId(), page: pageIndex, type: 'replace_text', text, size,
    rect: [rect[0], rect[1], rect[2], Math.min(rect[1] + height, page.raw?.h || page.height)],
    color: [0.07, 0.09, 0.13], opacity: 1, width: 1,
  });
  item.text = text;
  node.textContent = text;
  selectAnnotation(editor.annotations.at(-1).id);
  commit();
}

/* ---------- 页面级操作与双向同步 ---------- */
/**
 * 输入框是「原始 1-based 页码」的唯一真源：
 * 后端 edit_pdf 的 delete_pages / rotations / page_order 都按原始页码解释。
 */
function syncPageEditFields() {
  const originals = editor.pages.map((p) => p.original);
  const visible = editor.pages.filter((p) => !p.deleted).map((p) => p.original);
  const deleted = new Set([...editor.deletedOriginal]);
  originals.forEach((n, i) => { if (editor.pages[i].deleted) deleted.add(n); });
  $('deletePages').value = [...deleted].sort((a, b) => a - b).join(',');

  const rotations = new Map(editor.manualRotations);
  editor.pages.forEach((p) => { if (p.rotation) rotations.set(p.original, p.rotation); });
  $('rotatePage').value = [...rotations.entries()].sort((a, b) => a[0] - b[0]).map(([p, a]) => `${p}:${a}`).join(',');

  // 只有顺序真的变了才写 page_order：空值表示「保持原顺序」（后端同样按此解释）
  const unchanged = visible.length === originals.length && visible.every((v, i) => v === originals[i]);
  $('pageOrder').value = unchanged ? '' : visible.join(',');
}

function deleteCurrentPage() {
  if (editor.pages.length <= 1) { toast('至少要保留一页', 'warn'); return; }
  const [removed] = editor.pages.splice(editor.current, 1);
  editor.deletedOriginal.add(removed.original);
  // 标注页码要跟着往前挪，否则会落到错误的页上
  editor.annotations = editor.annotations
    .filter((a) => a.page !== editor.current)
    .map((a) => (a.page > editor.current ? { ...a, page: a.page - 1 } : a));
  editor.current = Math.min(editor.current, editor.pages.length - 1);
  editor.selectedId = null;
  toast(`已删除预览中的第 ${editor.current + 1} 页（原第 ${removed.original} 页）`, 'info');
  commit({ rebuild: true });
}

/* ---------- 提交到后端 ----------
 * 关键约定（对应 converter.edit_pdf）：
 *   delete_pages / rotations 的页码 = 原始 PDF 的 1-based 页码
 *   page_order 是「原始终页码」按目标顺序排列的列表（可重复，用来复制页）；
 *              它同时充当白名单——没列进去的页会被丢弃
 *   标注的 page 是 0-based，且在后端先于重排写入，所以要用
 *              「该页在 page_order 中的位置」而不是原始页码
 */
function currentOrder() {
  const order = editor.pages.filter((p) => !p.deleted).map((p) => p.original);
  // 超过 50 页的文档只加载了前 50 页预览，其余页面必须原样追加，
  // 否则保存会把这些页面丢掉。
  for (let n = editor.pages.length + 1; n <= editor.totalPages; n++) order.push(n);
  return order;
}

async function submitEditor() {
  const body = new FormData();
  body.append('file', state.files[0]);

  // 删除页：可视化删除 + 手填页码，都按原始 1-based 页码提交
  const manualDeleted = $('deletePages').value.split(',').map((v) => v.trim()).filter(Boolean).map(Number).filter((n) => n > 0);
  const visualDeleted = editor.pages.filter((p) => p.deleted).map((p) => p.original);
  body.append('delete_pages', [...new Set([...manualDeleted, ...visualDeleted])].join(','));

  // 旋转：合并手填与可视化，键为原始页码
  const rotations = {};
  $('rotatePage').value.split(',').forEach((chunk) => {
    const [page, angle] = chunk.split(':').map((v) => v.trim());
    if (page && angle && Number(page) > 0) rotations[Number(page)] = Number(angle) || 0;
  });
  editor.pages.filter((p) => p.rotation).forEach((p) => { rotations[p.original] = p.rotation; });
  body.append('rotations', JSON.stringify(rotations));

  const order = currentOrder();
  body.append('page_order', order.join(','));

  const marks = [];
  editor.annotations.forEach((a) => {
    const page = editor.pages[a.page];
    if (!page || page.deleted) return;
    // 该页在最终顺序中的位置（重复页取其第一次出现的位置）
    const position = order.indexOf(page.original);
    if (position < 0) return;
    const mark = { ...a, page: position };
    delete mark.id;
    marks.push(mark);
  });
  // 后端只在 annotations 非空时才走「先擦除再写入」的通道，
  // 空数组同样会触发全页 apply_redactions，因此没有标注时干脆不传。
  if (marks.length) body.append('annotations', JSON.stringify(marks));

  return request('/api/pdf-edit', { body });
}

/* ---------- 编辑器事件绑定 ---------- */
$$('.tab').forEach((tab) => {
  tab.addEventListener('click', () => setMode(tab.dataset.mode));
});

$$('#toolGrid .tool-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    $$('#toolGrid .tool-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    editor.tool = btn.dataset.tool;
    $$('#pdfStage .pdf-page').forEach((box) => box.classList.toggle('tool-select', editor.tool === 'select'));
    if (editor.tool === 'text' && !$('editText').value.trim()) $('editText').focus();
  });
});

$('undoBtn').addEventListener('click', () => restoreHistory(-1));
$('redoBtn').addEventListener('click', () => restoreHistory(1));
$('zoomOut').addEventListener('click', () => { editor.zoom = Math.max(0.4, Math.round((editor.zoom - 0.1) * 10) / 10); renderEditor(); });
$('zoomIn').addEventListener('click', () => { editor.zoom = Math.min(2, Math.round((editor.zoom + 0.1) * 10) / 10); renderEditor(); });
$('fitWidth').addEventListener('click', () => { editor.zoom = 1; renderEditor(); });
$('pageJump').addEventListener('change', (e) => goToPage(Number(e.target.value) - 1));
$('fullscreenBtn').addEventListener('click', () => {
  const node = $('editor');
  if (document.fullscreenElement) document.exitFullscreen?.();
  else node.requestFullscreen?.().catch(() => toast('当前浏览器不支持全屏', 'warn'));
});
$('duplicatePage').addEventListener('click', () => {
  const page = editor.pages[editor.current];
  if (!page) return;
  const copy = { ...page, id: 'page-dup-' + Date.now(), duplicated: true };
  editor.pages.splice(editor.current + 1, 0, copy);
  editor.annotations.forEach((a) => { if (a.page > editor.current) a.page += 1; });
  editor.current += 1;
  commit({ rebuild: true });
});
$('rotateCurrentPage').addEventListener('click', () => {
  const page = editor.pages[editor.current];
  if (!page) return;
  page.rotation = ((page.rotation || 0) + 90) % 360;
  editor.manualRotations.set(page.original, page.rotation);
  commit({ rebuild: true });
});
$('deleteCurrentPage').addEventListener('click', deleteCurrentPage);
$('clearMarks').addEventListener('click', () => {
  if (!editor.annotations.length) return;
  editor.annotations = [];
  editor.selectedId = null;
  commit({ rebuild: true });
  toast('已清空所有标注', 'ok');
});
$('convertPdfWord').addEventListener('click', convertPdfToWord);
$('editText').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.preventDefault(); });

document.addEventListener('keydown', (e) => {
  if (state.mode !== 'pdfedit' || !state.editorLoaded) return;
  const typing = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName) || editor.editing;
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 'z') {
    e.preventDefault();
    restoreHistory(e.shiftKey ? 1 : -1);
  } else if (mod && e.key.toLowerCase() === 's') {
    e.preventDefault();
    if (!state.busy && !$('actionBtn').disabled) runPdfEdit();
  } else if (mod && e.key.toLowerCase() === 'y') {
    e.preventDefault();
    restoreHistory(1);
  } else if ((e.key === 'Delete' || e.key === 'Backspace') && !typing && selectedAnnotation()) {
    e.preventDefault();
    editor.annotations = editor.annotations.filter((a) => a !== selectedAnnotation());
    editor.selectedId = null;
    commit({ rebuild: true });
  } else if (e.key === 'PageDown' && !typing) {
    e.preventDefault(); goToPage(editor.current + 1);
  } else if (e.key === 'PageUp' && !typing) {
    e.preventDefault(); goToPage(editor.current - 1);
  }
});

/** 手填「页面顺序」→ 重排可视化页面（用原始页码指代，可重复以复制页）。 */
function applyManualOrder() {
  const requested = $('pageOrder').value.split(',').map((v) => Number(v.trim())).filter((n) => n > 0);
  if (!requested.length) { renderEditor(); return; }
  const pool = new Map();
  editor.pages.forEach((p) => {
    if (!pool.has(p.original)) pool.set(p.original, []);
    pool.get(p.original).push(p);
  });
  const rebuilt = [];
  requested.forEach((original) => {
    const candidates = pool.get(original);
    if (!candidates?.length) return;
    const page = candidates.shift();
    page.deleted = false;
    rebuilt.push(page);
  });
  // 没被列出的页视为删除（与后端 page_order 的白名单语义一致）
  const listed = new Set(rebuilt.map((p) => p.original));
  editor.deletedOriginal = new Set(
    [...editor.deletedOriginal].filter((n) => !listed.has(n))
      .concat(editor.pages.filter((p) => !listed.has(p.original)).map((p) => p.original)),
  );
  if (!rebuilt.length) {
    toast('页面顺序无效，至少需要保留一页', 'warn');
    renderEditor();
    return;
  }
  editor.pages = rebuilt;
  editor.current = Math.min(editor.current, editor.pages.length - 1);
  editor.selectedId = null;
  commit({ rebuild: true });
}

// 手填页码 → 同步到可视化状态（只影响预览中仍存在的页面）
$('deletePages').addEventListener('change', () => {
  const pages = new Set($('deletePages').value.split(',').map((v) => Number(v.trim())).filter((n) => n > 0));
  editor.pages.forEach((p) => { p.deleted = pages.has(p.original); });
  if (editor.pages.every((p) => p.deleted)) {
    toast('不能删除全部页面，已忽略该输入', 'warn');
    editor.pages.forEach((p) => { p.deleted = false; });
  } else if (editor.pages[editor.current]?.deleted) {
    const next = editor.pages.findIndex((p) => !p.deleted);
    if (next >= 0) editor.current = next;
  }
  commit({ rebuild: true });
});

// 手填旋转 → 同步到可视化状态（键为原始页码）
$('rotatePage').addEventListener('change', () => {
  editor.manualRotations = new Map();
  $('rotatePage').value.split(',').forEach((chunk) => {
    const [page, angle] = chunk.split(':').map((v) => v.trim());
    const pageNo = Number(page);
    const deg = Number(angle);
    if (pageNo > 0 && Number.isFinite(deg)) editor.manualRotations.set(pageNo, ((deg % 360) + 360) % 360);
  });
  editor.pages.forEach((p) => { p.rotation = editor.manualRotations.get(p.original) || 0; });
  commit({ rebuild: true });
});

$('pageOrder').addEventListener('change', applyManualOrder);

/* ============================================================
 * 8. 启动
 * ============================================================ */
$('actionBtn').addEventListener('click', runAction);
$('quickPdfEdit').addEventListener('click', () => setMode('pdfedit'));
$('quickPdfMerge').addEventListener('click', () => setMode('pdfmerge'));
$('quickTranslate').addEventListener('click', () => setMode('translate'));

// 选项卡右侧渐隐提示
function updateTabsFade() {
  const wrap = $('tabsWrap');
  const tabs = $('tabs');
  wrap.classList.toggle('at-end', tabs.scrollLeft + tabs.clientWidth >= tabs.scrollWidth - 4);
}
$('tabs').addEventListener('scroll', updateTabsFade);
window.addEventListener('resize', updateTabsFade);

function setServerPill(kind, text) {
  const pill = $('serverPill');
  pill.className = 'pill ' + kind;
  $('serverPillText').textContent = text;
}

async function boot() {
  initTheme();
  resetEditor();
  renderFileBar();
  setMode('convert', { silent: true });
  updateTabsFade();

  try {
    const [formats, config] = await Promise.all([
      fetch('/api/formats').then((r) => r.json()),
      fetch('/api/config').then((r) => r.json()),
    ]);
    state.formats = formats || {};
    state.limits = {
      maxFileMb: config.max_file_size_mb ?? 200,
      maxBatchMb: config.max_batch_size_mb ?? 500,
    };
    state.translateReady = !!config.translate_configured;
    state.translateModel = config.translate_model || '';
    const limitText = `单文件最大 ${state.limits.maxFileMb}MB，批量总计 ${state.limits.maxBatchMb}MB`;
    $('uploadLimitHint').textContent = limitText;
    $('footerLimits').textContent = `单文件 ${state.limits.maxFileMb}MB，批量 ${state.limits.maxBatchMb}MB`;
    const glossary = $('translateGlossary');
    if (!glossary.value && config.default_glossary) glossary.value = config.default_glossary;
    setServerPill('ok', state.translateReady ? `服务正常 · 翻译已配置` : '服务正常 · 翻译未配置');
  } catch {
    setServerPill('bad', '无法连接后端服务');
    showActionInfo('无法读取后端配置，请确认服务已启动（python main.py 或 start.bat）。');
  }

  updateModeUI();
}

boot();
