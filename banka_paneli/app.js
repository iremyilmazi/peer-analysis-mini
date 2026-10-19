'use strict';
const $ = selector => document.querySelector(selector);
let mode = 'dry', cursor = 0, lastState = 'IDLE', context = {ready: false};
let polling = false, starting = false, contextSequence = 0;
const progressMap = {'1-KAP':18,'2-EXCEL':43,'3-PDF':68,'3B-BIRLESTIR':82,'4-KONTROL':94,'Tamamlandi':100,'Kayit':98,'Dry-run':98};

async function api(url, options) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'İşlem tamamlanamadı.');
  return data;
}
function post(url, data = {}) {
  return api(url, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
}
function fmt(seconds) {
  seconds = Math.max(0, Math.floor(seconds || 0));
  return `${String(Math.floor(seconds/60)).padStart(2,'0')}:${String(seconds%60).padStart(2,'0')}`;
}
function syncControls() {
  const running = lastState === 'RUNNING';
  $('#bank').disabled = running || starting;
  $('#refreshContext').disabled = running || starting;
  document.querySelectorAll('.mode button').forEach(b => b.disabled = running || starting);
  $('#runButton').disabled = starting || (!running && !context.ready);
}
async function refreshContext() {
  const sequence = ++contextSequence, bank = $('#bank').value;
  context = {ready:false};
  syncControls();
  const data = await api('/api/context?bank=' + encodeURIComponent(bank));
  if (sequence !== contextSequence || bank !== $('#bank').value) return;
  context = data;
  $('#latestPeriod').textContent = context.latest_period || 'Dosya/dönem bulunamadı';
  $('#nextPeriod').textContent = context.next_period || '—';
  syncControls();
  if (lastState === 'IDLE') $('#heroDetail').textContent = context.ready ? 'Yalnız '+context.next_period+' dönemi aranacak.' : context.error;
}
document.querySelectorAll('.mode button').forEach(button => button.onclick = () => {
  if (lastState === 'RUNNING' || starting) return;
  mode = button.dataset.mode;
  document.querySelectorAll('.mode button').forEach(b => b.classList.toggle('active', b === button));
  $('#safetyText').textContent = mode === 'dry' ? 'Ana Excel değiştirilmez. Sonuç ayrı bir kontrol kopyasına yazılır.' : 'Ana Excel önce yedeklenir; tüm kontroller geçerse atomik olarak güncellenir.';
});
$('#bank').onchange = () => refreshContext().catch(showError);
$('#refreshContext').onclick = () => refreshContext().catch(showError);
function showError(error) { $('#heroDetail').textContent = error.message; }
$('#runButton').onclick = async () => {
  if (starting) return;
  try {
    if (lastState === 'RUNNING') { await post('/api/stop'); return; }
    if (mode === 'live' && !window.confirm('Çalışma Excel’iniz yedeklenip güncellenecek. Devam edilsin mi?')) return;
    starting = true;
    syncControls();
    await refreshContext();
    if (!context.ready) throw new Error(context.error || 'Excel hazır değil.');
    cursor = 0;
    $('#log').textContent = '';
    render(await post('/api/run', {bank:$('#bank').value, dry_run:mode === 'dry'}));
  } catch (error) { showError(error); }
  finally { starting = false; syncControls(); }
};
$('#clearLogs').onclick = () => { $('#log').textContent = ''; };
function addLogs(lines) {
  const log = $('#log');
  for (const line of lines) {
    const span = document.createElement('span');
    span.className = line.includes('HATA') || line.includes('ERROR') ? 'error' : line.includes('Tamamlandi') || line.includes('BASARILI') ? 'ok' : 'time';
    span.textContent = line + '\n';
    log.appendChild(span);
  }
  if (lines.length) log.scrollTop = log.scrollHeight;
}
function render(state) {
  const wasRunning = lastState === 'RUNNING';
  lastState = state.state;
  addLogs(state.logs || []);
  cursor = state.cursor ?? cursor;
  const running = lastState === 'RUNNING', ok = lastState === 'OK', error = lastState === 'ERROR', noNew = lastState === 'NO_NEW_PERIOD';
  syncControls();
  $('#runButton').textContent = running ? '■ Çalışmayı Durdur' : '↓ Güncellemeyi Başlat';
  $('#runButton').classList.toggle('stop', running);
  $('#statusLabel').textContent = `● DURUM: ${running?'ÇALIŞIYOR':ok?'TAMAMLANDI':noNew?'YENİ DÖNEM YOK':error?'HATA':'HAZIR'}`;
  $('#statusLabel').style.color = error ? 'var(--red)' : running || noNew ? 'var(--amber)' : 'var(--green)';
  $('#heroTitle').textContent = running ? (state.bank_name || 'Banka')+' güncelleniyor' : ok ? 'İşlem Tamamlandı' : noNew ? 'Beklenen Dönem Bulunamadı' : error ? 'İşlem Durduruldu' : 'Güncellemeye Hazır';
  $('#heroDetail').textContent = running ? state.step : ok ? 'Kontroller geçti. Çıktı dosyalarını aşağıdan açabilirsiniz.' : noNew ? 'Excel değişmedi. Daha sonraki bir döneme atlanmadı.' : error ? state.error || 'Konsol ayrıntılarını kontrol edin.' : context.ready ? 'Yalnız '+context.next_period+' dönemi aranacak.' : context.error || 'Excel kontrol ediliyor.';
  const pct = running ? progressMap[state.step] || state.progress || 5 : ok ? 100 : 0;
  $('#progressBar').style.width = pct+'%';
  $('#progressText').textContent = pct+'%';
  $('#stepLabel').textContent = state.step || 'Bekleniyor';
  $('#duration').textContent = $('#totalTime').textContent = fmt(state.elapsed);
  $('#responseTime').textContent = state.last_response_seconds == null ? '—' : fmt(state.last_response_seconds);
  $('#systemState').textContent = running ? 'İŞLEM DEVAM EDİYOR' : error || !context.ready ? 'KONTROL GEREKİYOR' : 'SİSTEM HAZIR';
  renderFiles(state.files || []);
  if (wasRunning && !running) refreshContext().catch(showError);
}
function renderFiles(files) {
  $('#fileCount').textContent = `${files.length} dosya`;
  const list = $('#fileList');
  list.replaceChildren();
  list.className = files.length ? '' : 'empty';
  if (!files.length) { list.textContent = 'Henüz bir çalışma çıktısı oluşmadı.'; return; }
  for (const file of files) {
    const row = document.createElement('div'); row.className = 'file-row';
    const info = document.createElement('div'); info.className = 'file-info';
    const label = document.createElement('strong'); label.textContent = file.label;
    const path = document.createElement('span'); path.textContent = file.path;
    info.append(label, path);
    const button = document.createElement('button'); button.className = 'ghost'; button.textContent = 'Klasörü Aç';
    button.onclick = () => post('/api/open', {path:file.path}).catch(showError);
    row.append(info, button); list.appendChild(row);
  }
}
async function poll() {
  if (polling) return;
  polling = true;
  try { render(await api('/api/status?cursor='+cursor)); }
  catch (error) { $('#systemState').textContent = 'BAĞLANTI HATASI'; }
  finally { polling = false; }
}
async function init() {
  try {
    const data = await api('/api/banks');
    $('#bank').replaceChildren(...data.banks.map(bank => new Option(bank.name, bank.key)));
    await refreshContext();
    await poll();
    setInterval(poll, 750);
  } catch (error) { $('#systemState').textContent = 'BAĞLANTI HATASI'; showError(error); }
}
init();
