const fileInput = document.querySelector('#file');
const chooseButton = document.querySelector('#choose');
const selection = document.querySelector('#selection');
const errorBox = document.querySelector('#error');
const work = document.querySelector('#work');
const rows = document.querySelector('#rows');
const counter = document.querySelector('#counter');
const totalProgress = document.querySelector('#total-progress');
const result = document.querySelector('#result');
const createUrl = document.body.dataset.createUrl;

chooseButton.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) upload(fileInput.files[0]);
});

function showError(message) {
  errorBox.textContent = message;
  errorBox.hidden = false;
}

async function upload(file) {
  errorBox.hidden = true;
  result.hidden = true;
  selection.textContent = `${file.name} · проверяем файл…`;
  chooseButton.disabled = true;
  const body = new FormData();
  body.append('file', file);
  try {
    const response = await fetch(createUrl, {method: 'POST', body});
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Ошибка загрузки файла');
    selection.textContent = `${file.name} · ${payload.total} позиций`;
    work.hidden = false;
    render(payload);
    poll(payload.status_url);
  } catch (error) {
    selection.textContent = file.name;
    showError(error.message);
    chooseButton.disabled = false;
  }
}

function statusCell(row) {
  if (row.status === 'processing') return '<span class="status active"><i></i>Обработка</span>';
  if (row.status === 'done') return `<span class="status done"><i>✓</i>Готово · <a href="${row.file_url}">Карточка</a> · <a href="${row.models_url}">Модели</a> · <a href="${row.json_url}">JSON</a></span>`;
  if (row.status === 'error') return `<span class="status failed" title="${escapeHtml(row.error || '')}"><i>!</i>Ошибка</span>`;
  return '<span class="status waiting"><i></i>В очереди</span>';
}

function escapeHtml(value) {
  const node = document.createElement('div');
  node.textContent = value;
  return node.innerHTML;
}

function render(job) {
  rows.innerHTML = job.rows.map(row => `<tr>
    <td class="article">${escapeHtml(row.article)}</td>
    <td>${escapeHtml(row.name)}</td>
    <td>${statusCell(row)}</td>
  </tr>`).join('');
  counter.textContent = `${job.complete} из ${job.total}`;
  totalProgress.style.width = `${job.total ? job.complete / job.total * 100 : 0}%`;
  if (job.status === 'done') {
    document.querySelector('#folder-link').href = job.folder_url;
    document.querySelector('#download-link').href = job.download_url;
    document.querySelector('#local-path').textContent = job.local_path;
    result.hidden = false;
    chooseButton.disabled = false;
  }
}

async function poll(statusUrl) {
  try {
    const response = await fetch(statusUrl, {cache: 'no-store'});
    if (!response.ok) throw new Error('Не удалось получить состояние задачи');
    const job = await response.json();
    render(job);
    if (job.status !== 'done') setTimeout(() => poll(job.status_url), 900);
  } catch (error) {
    showError(error.message);
    chooseButton.disabled = false;
  }
}
