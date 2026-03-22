import os
import io
import tempfile
from flask import Flask, request, send_file, jsonify
import pandas as pd
from btg_consolidador import consolidar

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

ALLOWED_EXT = {'xlsx', 'xls'}

def allowed(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Apex Partners · Consolidador BTG</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --apex-blue:  #0077D1;
  --apex-dark:  #004F8C;
  --apex-light: #E8F4FF;
  --bg:         #F4F6F9;
  --white:      #FFFFFF;
  --border:     #DDE3EC;
  --text:       #1A2332;
  --muted:      #6B7A92;
  --green:      #0C8A4E;
  --green-bg:   #EAF7F0;
  --red:        #C0392B;
  --red-bg:     #FDECEA;
  --shadow-sm:  0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md:  0 4px 16px rgba(0,0,0,0.08), 0 2px 6px rgba(0,0,0,0.04);
  --radius:     10px;
}

html, body { min-height: 100vh; background: var(--bg); color: var(--text); font-family: 'Inter', sans-serif; font-size: 14px; line-height: 1.6; }

/* topbar */
.topbar {
  background: var(--white);
  border-bottom: 1px solid var(--border);
  padding: 0 40px;
  height: 60px;
  display: flex; align-items: center; justify-content: space-between;
  position: sticky; top: 0; z-index: 100;
  box-shadow: var(--shadow-sm);
}
.topbar-logo { display: flex; align-items: center; gap: 14px; }
.topbar-logo img { height: 28px; width: auto; display: block; }
.topbar-divider { width: 1px; height: 22px; background: var(--border); }
.topbar-title { font-size: 13px; font-weight: 500; color: var(--muted); }
.topbar-badge {
  font-family: 'IBM Plex Mono', monospace; font-size: 10px;
  color: var(--apex-blue); background: var(--apex-light);
  border: 1px solid rgba(0,119,209,0.2); padding: 3px 10px;
  border-radius: 99px; font-weight: 500;
}

/* hero */
.hero {
  background: linear-gradient(135deg, var(--apex-dark) 0%, var(--apex-blue) 100%);
  padding: 56px 40px 60px; position: relative; overflow: hidden;
}
.hero::before {
  content: ''; position: absolute; top: -60px; right: -80px;
  width: 400px; height: 400px; border-radius: 50%;
  background: rgba(255,255,255,0.04); pointer-events: none;
}
.hero::after {
  content: ''; position: absolute; bottom: -100px; left: -40px;
  width: 300px; height: 300px; border-radius: 50%;
  background: rgba(255,255,255,0.03); pointer-events: none;
}
.hero-inner { max-width: 820px; margin: 0 auto; position: relative; z-index: 1; }
.hero h1 { font-size: clamp(22px, 3.5vw, 32px); font-weight: 700; color: #fff; letter-spacing: -0.025em; line-height: 1.2; margin-bottom: 10px; }
.hero p { font-size: 14px; color: rgba(255,255,255,0.65); max-width: 480px; line-height: 1.7; }
.hero-steps { display: flex; gap: 8px; margin-top: 28px; flex-wrap: wrap; }
.hero-step { display: flex; align-items: center; gap: 8px; background: rgba(255,255,255,0.10); border: 1px solid rgba(255,255,255,0.15); border-radius: 99px; padding: 6px 14px 6px 10px; font-size: 11px; color: rgba(255,255,255,0.85); font-weight: 500; }
.step-num { width: 18px; height: 18px; background: rgba(255,255,255,0.2); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 9px; font-weight: 700; color: #fff; flex-shrink: 0; }

/* main */
.main { max-width: 820px; margin: 0 auto; padding: 40px 40px 80px; }

/* section head */
.section-head { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; }
.section-num { width: 24px; height: 24px; background: var(--apex-blue); color: #fff; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; flex-shrink: 0; }
.section-head h2 { font-size: 13px; font-weight: 600; color: var(--text); }

/* portfolio input */
.portfolio-field {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px 24px;
  margin-bottom: 32px;
  box-shadow: var(--shadow-sm);
  display: flex;
  align-items: center;
  gap: 20px;
  flex-wrap: wrap;
}
.portfolio-field label {
  font-size: 13px;
  font-weight: 600;
  color: var(--text);
  white-space: nowrap;
  flex-shrink: 0;
}
.portfolio-field label span {
  display: block;
  font-size: 11px;
  font-weight: 400;
  color: var(--muted);
  margin-top: 2px;
}
.portfolio-input-wrap { flex: 1; min-width: 200px; position: relative; }
.portfolio-input-wrap input {
  width: 100%;
  padding: 10px 14px;
  border: 1.5px solid var(--border);
  border-radius: 8px;
  font-family: 'IBM Plex Mono', monospace;
  font-size: 13px;
  font-weight: 500;
  color: var(--text);
  background: var(--bg);
  transition: border-color 0.2s, box-shadow 0.2s;
  outline: none;
}
.portfolio-input-wrap input:focus {
  border-color: var(--apex-blue);
  box-shadow: 0 0 0 3px rgba(0,119,209,0.12);
  background: var(--white);
}
.portfolio-input-wrap input.filled {
  border-color: var(--green);
  background: var(--white);
}
.portfolio-input-wrap::before {
  
  position: absolute; left: 12px; top: 50%; transform: translateY(-50%);
  font-size: 14px; pointer-events: none;
}
.portfolio-preview {
  font-size: 11px;
  color: var(--muted);
  margin-top: 6px;
  font-family: 'IBM Plex Mono', monospace;
  min-height: 16px;
  transition: color 0.2s;
}
.portfolio-preview.active { color: var(--apex-blue); }

/* upload grid */
.upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 32px; }
@media (max-width: 560px) { .upload-grid { grid-template-columns: 1fr; } }

.drop-card {
  background: var(--white); border: 2px dashed var(--border);
  border-radius: var(--radius); padding: 32px 20px 24px;
  text-align: center; cursor: pointer; position: relative;
  transition: border-color 0.2s, background 0.2s, box-shadow 0.2s, transform 0.2s;
  user-select: none;
}
.drop-card:hover, .drop-card.drag-over { border-color: var(--apex-blue); border-style: solid; background: var(--apex-light); transform: translateY(-2px); box-shadow: var(--shadow-md); }
.drop-card.has-file { border-style: solid; border-color: var(--green); background: var(--green-bg); box-shadow: var(--shadow-sm); }
.drop-card input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
.drop-icon { font-size: 28px; margin-bottom: 12px; display: block; transition: transform 0.2s; }
.drop-card:hover .drop-icon, .drop-card.drag-over .drop-icon { transform: translateY(-3px); }
.drop-card.has-file .drop-icon { opacity: 0.3; }
.drop-label { font-size: 13px; font-weight: 600; color: var(--text); margin-bottom: 5px; display: block; }
.drop-sub { font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: var(--muted); line-height: 1.8; }
.file-name { display: none; margin-top: 12px; font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: var(--green); font-weight: 500; word-break: break-all; line-height: 1.5; padding: 6px 10px; border-radius: 6px; background: rgba(12,138,78,0.08); }
.has-file .file-name { display: block; }
.has-file .drop-sub  { display: none; }
.check-badge { display: none; position: absolute; top: 10px; right: 12px; background: var(--green); color: #fff; font-size: 9px; font-weight: 700; padding: 2px 8px; border-radius: 99px; letter-spacing: 0.05em; }
.has-file .check-badge { display: block; }

/* status */
#status { display: none; padding: 11px 15px; border-radius: 8px; font-size: 12px; margin-bottom: 12px; animation: fadein 0.25s ease; font-family: 'IBM Plex Mono', monospace; border-left: 3px solid; }
#status.info  { background: var(--apex-light); border-color: var(--apex-blue); color: var(--apex-dark); }
#status.ok    { background: var(--green-bg);   border-color: var(--green);     color: var(--green); }
#status.error { background: var(--red-bg);     border-color: var(--red);       color: var(--red); }
.dots span { display: inline-block; animation: blink 1.2s infinite; }
.dots span:nth-child(2) { animation-delay: 0.2s; }
.dots span:nth-child(3) { animation-delay: 0.4s; }

/* run button */
#run {
  width: 100%; padding: 15px 24px;
  background: var(--apex-blue); color: #fff;
  border: none; border-radius: var(--radius); cursor: pointer;
  font-family: 'Inter', sans-serif; font-size: 13px; font-weight: 600;
  transition: background 0.2s, transform 0.15s, box-shadow 0.2s;
  box-shadow: 0 2px 8px rgba(0,119,209,0.3);
}
#run:hover:not(:disabled) { background: var(--apex-dark); transform: translateY(-1px); box-shadow: 0 6px 20px rgba(0,119,209,0.35); }
#run:active:not(:disabled) { transform: translateY(0); }
#run:disabled { opacity: 0.4; cursor: not-allowed; box-shadow: none; }

/* log */
#log { display: none; background: #1A2332; border-radius: var(--radius); padding: 16px 18px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; line-height: 2; max-height: 160px; overflow-y: auto; margin-bottom: 24px; animation: fadein 0.25s ease; }
#log::-webkit-scrollbar { width: 4px; }
#log::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.15); border-radius: 2px; }
.ln-ok     { color: #4ADE80; }
.ln-err    { color: #F87171; }
.ln-dim    { color: rgba(255,255,255,0.35); }
.ln-prompt { color: #60A5FA; }

/* download */
#dl-wrap { display: none; animation: slideup 0.35s cubic-bezier(.22,.68,0,1.15); }
.dl-card { background: var(--white); border: 1px solid var(--border); border-radius: var(--radius); padding: 24px 28px; display: flex; align-items: center; justify-content: space-between; gap: 16px; box-shadow: var(--shadow-md); border-left: 4px solid var(--green); }
.dl-info strong { display: block; font-size: 14px; font-weight: 600; margin-bottom: 3px; }
.dl-info span { font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: var(--muted); }
#dl-btn { display: inline-flex; align-items: center; gap: 7px; padding: 11px 20px; background: var(--green); color: #fff; border: none; border-radius: 8px; font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 600; cursor: pointer; white-space: nowrap; transition: background 0.2s, transform 0.15s, box-shadow 0.2s; box-shadow: 0 2px 8px rgba(12,138,78,0.25); }
#dl-btn:hover { background: #0a7a44; transform: translateY(-1px); box-shadow: 0 6px 18px rgba(12,138,78,0.3); }

/* info cards */
.info-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 48px; }
@media (max-width: 560px) { .info-row { grid-template-columns: 1fr; } }
.info-card { background: var(--white); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px 20px; box-shadow: var(--shadow-sm); }
.info-card-icon { font-size: 20px; margin-bottom: 10px; display: block; }
.info-card dt { font-size: 10px; font-weight: 600; letter-spacing: 0.12em; text-transform: uppercase; color: var(--apex-blue); margin-bottom: 7px; }
.info-card dd { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--muted); line-height: 1.9; }

/* footer */
footer { max-width: 820px; margin: 0 auto; padding: 24px 40px 40px; display: flex; justify-content: space-between; align-items: center; border-top: 1px solid var(--border); font-size: 10px; color: var(--muted); font-family: 'IBM Plex Mono', monospace; flex-wrap: wrap; gap: 8px; }

@media (max-width: 600px) { .topbar { padding: 0 20px; } .hero { padding: 40px 20px 44px; } .main { padding: 28px 20px 60px; } footer { padding: 20px; } }

@keyframes fadein  { from{opacity:0;}to{opacity:1;} }
@keyframes slideup { from{opacity:0;transform:translateY(12px);}to{opacity:1;transform:translateY(0);} }
@keyframes blink   { 0%,80%,100%{opacity:0;}40%{opacity:1;} }
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-logo">
    <img src="data:image/png;base64,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" alt="Apex Partners">
    <div class="topbar-divider"></div>
    <span class="topbar-title">Consolidador de Posições</span>
  </div>
  <span class="topbar-badge">BTG → ComDinheiro</span>
</div>

<div class="hero">
  <div class="hero-inner">
    <h1>Consolidação de Posições BTG</h1>
    <p>Informe o portfólio, faça upload do extrato BTG e do template. A planilha ComDinheiro é gerada automaticamente.</p>
    <div class="hero-steps">
      <div class="hero-step"><span class="step-num">1</span> Informe o portfólio</div>
      <div class="hero-step"><span class="step-num">2</span> Upload dos arquivos</div>
      <div class="hero-step"><span class="step-num">3</span> Gerar e baixar</div>
    </div>
  </div>
</div>

<div class="main">

  <!-- STEP 1: PORTFÓLIO -->
  <div class="section-head">
    <span class="section-num">1</span>
    <h2>Nome do portfólio</h2>
  </div>

  <div class="portfolio-field">
    <label>
      Portfólio
      <span>Será usado como identificador na planilha</span>
    </label>
    <div class="portfolio-input-wrap">
      <input type="text" id="portfolio-input" placeholder="Ex: MNS_Onshore, JBL_Offshore…" autocomplete="new-password" autocorrect="off" autocapitalize="off" spellcheck="false" data-form-type="other" data-lpignore="true">
      <div class="portfolio-preview" id="portfolio-preview">Digite o nome do portfólio acima</div>
    </div>
  </div>

  <!-- STEP 2: ARQUIVOS -->
  <div class="section-head">
    <span class="section-num">2</span>
    <h2>Arquivos de entrada</h2>
  </div>

  <div class="upload-grid">
    <div class="drop-card" id="drop-extrato">
      <input type="file" id="file-extrato" accept=".xlsx,.xls">
      <span class="check-badge">✓ Selecionado</span>
      <span class="drop-icon">📊</span>
      <span class="drop-label">Extrato BTG Pactual</span>
      <span class="drop-sub">exemplo_extrato_btg_*.xlsx<br>Clique ou arraste o arquivo</span>
      <span class="file-name" id="name-extrato"></span>
    </div>
    <div class="drop-card" id="drop-template">
      <input type="file" id="file-template" accept=".xlsx,.xls">
      <span class="check-badge">✓ Selecionado</span>
      <span class="drop-icon">📋</span>
      <span class="drop-label">Planilha Template</span>
      <span class="drop-sub">planilha_cd_vazia.xlsx<br>Clique ou arraste o arquivo</span>
      <span class="file-name" id="name-template"></span>
    </div>
  </div>

  <!-- STEP 3: PROCESSAR -->
  <div class="section-head">
    <span class="section-num">3</span>
    <h2>Processar</h2>
  </div>

  <div id="status"></div>
  <button id="run" disabled>Gerar Planilha ComDinheiro</button>
  <div id="log"></div>

  <!-- DOWNLOAD -->
  <div id="dl-wrap">
    <div class="section-head" style="margin-top:28px;">
      <span class="section-num" style="background:var(--green);">↓</span>
      <h2>Download</h2>
    </div>
    <div class="dl-card">
      <div class="dl-info">
        <strong>comdinheiro_preenchida.xlsx</strong>
        <span id="dl-meta">Planilha gerada com sucesso</span>
      </div>
      <button id="dl-btn">↓ &nbsp;Baixar arquivo</button>
    </div>
  </div>

  <dl class="info-row">
    <div class="info-card">
      <span class="info-card-icon">📁</span>
      <dt>Ativos suportados</dt>
      <dd>Fundos de investimento<br>CRA / CRI<br>CDB · NTN-B<br>Carteira Administrada</dd>
    </div>
    <div class="info-card">
      <span class="info-card-icon">🔍</span>
      <dt>Fonte dos dados</dt>
      <dd>Aba Fundos (BTG)<br>Aba Renda Fixa<br>Posições Detalhadas<br>Carteiras Administradas</dd>
    </div>
    <div class="info-card">
      <span class="info-card-icon">⚙️</span>
      <dt>Configuração</dt>
      <dd>Instituição: BTG Pactual<br>Formato: ComDinheiro<br>Portfólio: dinâmico<br>Versão: 1.1</dd>
    </div>
  </dl>

</div>

<footer>
  <span>© 2025 Apex Partners · Gestão de Portfólios</span>
  <span>Ferramenta interna · uso exclusivo</span>
</footer>

<script>
const portfolioInput   = document.getElementById('portfolio-input');
const portfolioPreview = document.getElementById('portfolio-preview');
const dropExtrato  = document.getElementById('drop-extrato');
const dropTemplate = document.getElementById('drop-template');
const fileExtrato  = document.getElementById('file-extrato');
const fileTemplate = document.getElementById('file-template');
const nameExtrato  = document.getElementById('name-extrato');
const nameTemplate = document.getElementById('name-template');
const runBtn   = document.getElementById('run');
const statusEl = document.getElementById('status');
const logEl    = document.getElementById('log');
const dlWrap   = document.getElementById('dl-wrap');
const dlBtn    = document.getElementById('dl-btn');
const dlMeta   = document.getElementById('dl-meta');

let blobUrl = null;

// portfolio input feedback
portfolioInput.addEventListener('input', () => {
  const val = portfolioInput.value.trim();
  if (val) {
    portfolioPreview.textContent = 'nome_portfolio = "' + val + '"';
    portfolioPreview.className = 'portfolio-preview active';
    portfolioInput.classList.add('filled');
  } else {
    portfolioPreview.textContent = 'Digite o nome do portfólio acima';
    portfolioPreview.className = 'portfolio-preview';
    portfolioInput.classList.remove('filled');
  }
  checkReady();
});

function setFile(drop, nameEl, input, file) {
  if (!file) return;
  drop.classList.add('has-file');
  nameEl.textContent = file.name;
  input._file = file;
  checkReady();
}

function checkReady() {
  const hasPortfolio = portfolioInput.value.trim().length > 0;
  const hasExtrato   = !!fileExtrato._file;
  const hasTemplate  = !!fileTemplate._file;
  runBtn.disabled = !(hasPortfolio && hasExtrato && hasTemplate);

  if (!hasPortfolio && (hasExtrato || hasTemplate)) {
    runBtn.title = 'Informe o nome do portfólio primeiro';
  } else {
    runBtn.title = '';
  }
}

function setupDrop(drop, input, nameEl) {
  input.addEventListener('change', e => {
    if (e.target.files[0]) setFile(drop, nameEl, input, e.target.files[0]);
  });
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) setFile(drop, nameEl, input, f);
  });
}

setupDrop(dropExtrato, fileExtrato, nameExtrato);
setupDrop(dropTemplate, fileTemplate, nameTemplate);

runBtn.addEventListener('click', async () => {
  const extrato   = fileExtrato._file;
  const template  = fileTemplate._file;
  const portfolio = portfolioInput.value.trim();
  if (!extrato || !template || !portfolio) return;

  dlWrap.style.display = 'none';
  logEl.style.display = 'none';
  logEl.innerHTML = '';
  if (blobUrl) { URL.revokeObjectURL(blobUrl); blobUrl = null; }

  showStatus('info', '<span class="dots">Processando<span>.</span><span>.</span><span>.</span></span>');
  runBtn.disabled = true;
  logEl.style.display = 'block';

  addLog('prompt', '$ python btg_consolidador.py');
  addLog('dim',    '  portfólio: ' + portfolio);
  addLog('dim',    '  extrato:   ' + extrato.name);
  addLog('dim',    '  template:  ' + template.name);

  const fd = new FormData();
  fd.append('extrato',        extrato);
  fd.append('template',       template);
  fd.append('nome_portfolio', portfolio);

  try {
    const t0 = Date.now();
    const res = await fetch('/processar', { method: 'POST', body: fd });
    if (!res.ok) {
      const d = await res.json().catch(() => ({ erro: res.statusText }));
      throw new Error(d.erro || 'Erro desconhecido');
    }
    const blob = await res.blob();
    blobUrl = URL.createObjectURL(blob);
    const ms = Date.now() - t0;

    addLog('ok', '  ✓ Renda Fixa — posições extraídas');
    addLog('ok', '  ✓ Fundos — posições extraídas');
    addLog('ok', '  ✓ Planilha gerada em ' + (ms/1000).toFixed(1) + 's');
    showStatus('ok', '✓ Planilha gerada com sucesso — pronta para download');

    const fname = portfolio + '_comdinheiro.xlsx';
    dlMeta.textContent = portfolio + ' · ' + (ms/1000).toFixed(1) + 's · ' + (blob.size/1024).toFixed(0) + ' KB';
    dlWrap.style.display = 'block';
    dlBtn.onclick = () => {
      const a = document.createElement('a');
      a.href = blobUrl; a.download = fname; a.click();
    };
  } catch (err) {
    addLog('err', '  ✗ Erro: ' + err.message);
    showStatus('error', '✗ ' + err.message);
  } finally {
    runBtn.disabled = false;
  }
});

function showStatus(type, html) {
  statusEl.className = type; statusEl.innerHTML = html; statusEl.style.display = 'block';
}
function addLog(type, text) {
  const d = document.createElement('div');
  d.className = 'ln-' + type; d.textContent = text;
  logEl.appendChild(d); logEl.scrollTop = logEl.scrollHeight;
}
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return HTML, 200, {'Content-Type': 'text/html; charset=utf-8'}

@app.route('/processar', methods=['POST'])
def processar():
    try:
        extrato        = request.files.get('extrato')
        template       = request.files.get('template')
        nome_portfolio = request.form.get('nome_portfolio', 'JBL_Onshore').strip() or 'JBL_Onshore'

        if not extrato or not template:
            return jsonify({'erro': 'Envie os dois arquivos: extrato e planilha template.'}), 400
        if not allowed(extrato.filename) or not allowed(template.filename):
            return jsonify({'erro': 'Apenas arquivos .xlsx são suportados.'}), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            ext_path = os.path.join(tmpdir, 'extrato.xlsx')
            tpl_path = os.path.join(tmpdir, 'template.xlsx')
            out_path = os.path.join(tmpdir, 'saida.xlsx')

            extrato.save(ext_path)
            template.save(tpl_path)

            consolidar(ext_path, tpl_path, out_path, nome_portfolio=nome_portfolio)

            with open(out_path, 'rb') as f:
                data = f.read()

        return send_file(
            io.BytesIO(data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='comdinheiro_preenchida.xlsx'
        )

    except Exception as e:
        return jsonify({'erro': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7860, debug=False)
