<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Canvas PDI — Refinamento de Projetos</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js"></script>
<style>
:root{
  --navy:#1B3358;--navy-light:#243f6a;--blue:#2D6A9F;--blue-light:#E6F1FB;--blue-mid:#4A8DC0;
  --teal:#0C6B52;--teal-light:#E1F5EE;--teal-mid:#1A9970;
  --amber:#B57218;--amber-light:#FAF0D7;--amber-mid:#D4890A;
  --red:#9E2B2B;--red-light:#FCEBEB;--red-mid:#C43A3A;
  --green:#3A6B10;--green-light:#EAF3DE;--green-mid:#5A9E28;
  --orange:#8F3818;--orange-light:#FAECE7;--orange-mid:#C04E22;
  --purple:#4E45AE;--purple-light:#ECEAF9;--purple-mid:#6B62C8;
  --gray:#5A5857;--gray-light:#F5F3ED;--gray-mid:#888580;
  --border:#DDD9D3;--border-light:#EEEAE4;
  --text:#1C1B1A;--text-muted:#6B6965;--text-hint:#A09C97;
  --white:#FFFFFF;--surface:#FAFAF8;--bg:#F5F3ED;
  --sidebar-w:260px;
  --font:'DM Sans',system-ui,sans-serif;
  --font-mono:'DM Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;font-family:var(--font);font-size:14px;color:var(--text);background:var(--bg)}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:10px}

/* ── LAYOUT ───────────────────────────── */
.app{display:flex;flex-direction:column;height:100vh;overflow:hidden}
.topbar{height:52px;background:var(--navy);display:flex;align-items:center;justify-content:space-between;padding:0 1.25rem;flex-shrink:0;z-index:10}
.brand{display:flex;align-items:center;gap:10px}
.brand-dot{width:8px;height:8px;border-radius:50%;background:var(--teal-mid)}
.brand-name{font-size:15px;font-weight:600;color:#fff;letter-spacing:-.01em}
.brand-tag{font-size:11px;color:#7FA8CC;margin-left:2px}
.top-actions{display:flex;align-items:center;gap:8px}
.save-indicator{font-size:11px;padding:3px 10px;border-radius:10px;color:#7FA8CC;font-family:var(--font-mono)}
.save-indicator.saving{color:#FAC775}
.save-indicator.saved{color:#9FE1CB}

.body{display:flex;flex:1;overflow:hidden}

/* ── SIDEBAR ─────────────────────────── */
.sidebar{width:var(--sidebar-w);background:var(--white);border-right:1px solid var(--border-light);display:flex;flex-direction:column;flex-shrink:0}
.sidebar-top{padding:1rem;border-bottom:1px solid var(--border-light)}
.sidebar-top h3{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--text-hint);margin-bottom:.75rem}
.new-btn{width:100%;padding:8px;border-radius:8px;background:var(--navy);color:#fff;border:none;cursor:pointer;font-family:var(--font);font-size:13px;font-weight:500;display:flex;align-items:center;justify-content:center;gap:6px;transition:.15s}
.new-btn:hover{background:var(--navy-light)}
.project-list{flex:1;overflow-y:auto;padding:.5rem}
.project-card{padding:10px 12px;border-radius:8px;cursor:pointer;margin-bottom:3px;border:1px solid transparent;transition:.15s}
.project-card:hover{background:var(--gray-light)}
.project-card.active{background:#EAF1F9;border-color:#B5D4F4}
.pc-name{font-size:13px;font-weight:500;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.project-card.active .pc-name{color:var(--blue)}
.pc-meta{display:flex;align-items:center;gap:5px;margin-top:3px}
.status-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.pc-date{font-size:11px;color:var(--text-hint)}
.sidebar-footer{padding:.75rem 1rem;border-top:1px solid var(--border-light)}
.sidebar-stat{font-size:11px;color:var(--text-hint);font-family:var(--font-mono)}

/* ── MAIN ───────────────────────────── */
.main{flex:1;overflow-y:auto;background:var(--bg)}
.empty-wrap{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;text-align:center;padding:2rem}
.empty-icon{width:64px;height:64px;border-radius:16px;background:var(--white);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;margin:0 auto 1rem;font-size:28px}
.empty-wrap h2{font-size:18px;font-weight:500;margin-bottom:.5rem}
.empty-wrap p{font-size:13px;color:var(--text-muted);margin-bottom:1.5rem;max-width:320px}

.canvas-wrap{max-width:860px;margin:0 auto;padding:1.5rem 1.5rem 3rem}
.canvas-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:1.25rem;gap:1rem}
.canvas-header-left{flex:1}
.canvas-title{font-size:22px;font-weight:600;letter-spacing:-.02em;color:var(--text)}
.canvas-subtitle{font-size:12px;color:var(--text-muted);margin-top:3px;font-family:var(--font-mono)}
.canvas-header-actions{display:flex;gap:8px;flex-shrink:0}

/* ── BUTTONS ─────────────────────────── */
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border-radius:7px;font-family:var(--font);font-size:12px;font-weight:500;cursor:pointer;border:1px solid;transition:.15s;white-space:nowrap}
.btn:active{transform:scale(.98)}
.btn-navy{background:var(--navy);border-color:var(--navy);color:#fff}
.btn-navy:hover{background:var(--navy-light)}
.btn-teal{background:var(--teal);border-color:var(--teal);color:#fff}
.btn-teal:hover{background:#0a5c45}
.btn-outline{background:transparent;border-color:var(--border);color:var(--text-muted)}
.btn-outline:hover{background:var(--white);border-color:var(--border);color:var(--text)}
.btn-danger{background:transparent;border-color:#F09595;color:var(--red)}
.btn-danger:hover{background:var(--red-light)}
.btn-sm{padding:5px 11px;font-size:11px}
.btn-ghost-white{background:transparent;border-color:rgba(255,255,255,.25);color:#fff}
.btn-ghost-white:hover{background:rgba(255,255,255,.1)}

/* ── SECTION CARDS ───────────────────── */
.section-card{background:var(--white);border:1px solid var(--border-light);border-radius:12px;margin-bottom:1rem;overflow:hidden}
.section-head{display:flex;align-items:center;gap:10px;padding:.75rem 1.25rem;border-bottom:1px solid var(--border-light)}
.sec-num{width:22px;height:22px;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;flex-shrink:0;background:rgba(255,255,255,.2);color:#fff}
.sec-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:#fff;flex:1}
.sec-body{padding:1.25rem;display:flex;flex-direction:column;gap:.875rem}

/* ── FORM FIELDS ─────────────────────── */
.fg{display:grid;gap:.75rem}
.g2{grid-template-columns:1fr 1fr}
.g3{grid-template-columns:1fr 1fr 1fr}
.g4{grid-template-columns:1fr 1fr 1fr 1fr}
.field label{display:block;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--text-hint);margin-bottom:5px}
.field input,.field textarea,.field select{
  width:100%;padding:8px 11px;border-radius:7px;border:1px solid var(--border);
  background:var(--surface);color:var(--text);font-size:13px;font-family:var(--font);
  transition:.15s;resize:vertical
}
.field input:focus,.field textarea:focus,.field select:focus{
  outline:none;border-color:var(--blue-mid);background:var(--white);
  box-shadow:0 0 0 3px rgba(55,138,221,.12)
}
.field textarea{min-height:76px;line-height:1.5}
.field-danger input,.field-danger textarea{border-color:#F09595;background:#FFF8F8}
.field-danger label{color:var(--red-mid)}

/* ── TRL ─────────────────────────────── */
.trl-row{display:flex;gap:5px}
.trl-btn{flex:1;height:32px;border-radius:6px;border:1px solid var(--border);background:var(--surface);cursor:pointer;font-family:var(--font-mono);font-size:12px;font-weight:500;color:var(--text-hint);transition:.15s}
.trl-btn:hover{border-color:var(--blue-mid);color:var(--blue)}
.trl-btn.active{background:var(--blue-light);border-color:var(--blue-mid);color:var(--blue)}
.trl-labels{display:flex;justify-content:space-between;margin-top:4px}
.trl-label{font-size:10px;color:var(--text-hint)}

/* ── STATUS PILLS ────────────────────── */
.status-pills{display:flex;flex-wrap:wrap;gap:7px}
.spill{display:flex;align-items:center;gap:7px;padding:7px 13px;border-radius:20px;border:1px solid var(--border);cursor:pointer;font-size:12px;color:var(--text-muted);background:var(--surface);transition:.15s;user-select:none}
.spill:hover{border-color:var(--gray-mid)}
.spill-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.spill.s-blue{background:var(--blue-light);border-color:var(--blue-mid);color:#0C447C;font-weight:500}
.spill.s-amber{background:var(--amber-light);border-color:var(--amber-mid);color:#412402;font-weight:500}
.spill.s-orange{background:var(--orange-light);border-color:var(--orange-mid);color:#4A1B0C;font-weight:500}
.spill.s-green{background:var(--green-light);border-color:var(--green-mid);color:#173404;font-weight:500}
.spill.s-red{background:var(--red-light);border-color:var(--red-mid);color:#501313;font-weight:500}
.progress-bar{height:5px;border-radius:3px;background:var(--border-light);overflow:hidden;margin-top:10px}
.progress-fill{height:100%;border-radius:3px;background:var(--blue-mid);transition:width .4s ease}
.progress-lbl{font-size:10px;color:var(--text-hint);font-family:var(--font-mono);margin-top:5px}

/* ── STARS ───────────────────────────── */
.star-table{width:100%;border-collapse:collapse}
.star-table tr{border-bottom:1px solid var(--border-light)}
.star-table tr:last-child{border-bottom:none}
.star-table td{padding:7px 4px;vertical-align:middle}
.star-label{font-size:12px;color:var(--text-muted)}
.stars{display:flex;gap:4px}
.star{width:24px;height:24px;border-radius:4px;border:1px solid var(--border);background:var(--surface);cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;transition:.15s;color:transparent}
.star:hover{border-color:var(--amber-mid)}
.star.on{background:var(--amber-light);border-color:var(--amber-mid);color:var(--amber)}

/* ── PARTNER BADGES ──────────────────── */
.p-badges{display:flex;gap:8px;margin-top:.5rem}
.pbadge{padding:5px 16px;border-radius:20px;font-size:12px;font-weight:500;border:1px solid var(--border);cursor:pointer;background:var(--surface);color:var(--text-muted);transition:.15s}
.pbadge.pg.sel{background:var(--green-light);border-color:var(--green-mid);color:var(--green)}
.pbadge.pa.sel{background:var(--amber-light);border-color:var(--amber-mid);color:var(--amber)}
.pbadge.pr.sel{background:var(--red-light);border-color:var(--red-mid);color:var(--red)}

/* ── CHECKS ──────────────────────────── */
.check-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.check-item{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:6px;cursor:pointer;transition:.15s}
.check-item:hover{background:var(--gray-light)}
.check-item input[type=checkbox]{width:15px;height:15px;cursor:pointer;accent-color:var(--teal)}
.check-item span{font-size:12px;color:var(--text)}

/* ── REC OPTIONS ─────────────────────── */
.rec-opts{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:.875rem}
.rec-opt{display:flex;align-items:center;gap:8px;padding:10px 16px;border-radius:8px;border:1px solid var(--border);cursor:pointer;font-size:13px;color:var(--text-muted);background:var(--surface);transition:.15s}
.rec-radio{width:14px;height:14px;border-radius:50%;border:1.5px solid var(--border);flex-shrink:0;transition:.15s}
.rec-opt.rt{background:var(--teal-light);border-color:var(--teal-mid);color:var(--teal);font-weight:500}
.rec-opt.rt .rec-radio{background:var(--teal);border-color:var(--teal)}
.rec-opt.rg{background:var(--green-light);border-color:var(--green-mid);color:var(--green);font-weight:500}
.rec-opt.rg .rec-radio{background:var(--green);border-color:var(--green)}
.rec-opt.ra{background:var(--amber-light);border-color:var(--amber-mid);color:var(--amber);font-weight:500}
.rec-opt.ra .rec-radio{background:var(--amber);border-color:var(--amber)}
.rec-opt.rr{background:var(--red-light);border-color:var(--red-mid);color:var(--red);font-weight:500}
.rec-opt.rr .rec-radio{background:var(--red);border-color:var(--red)}

/* ── FOOTER BAR ──────────────────────── */
.form-footer{display:flex;justify-content:space-between;align-items:center;padding:1rem 0 2rem}

/* ── PRINT ───────────────────────────── */
@media print{
  .sidebar,.topbar,.canvas-header-actions,.form-footer,#toast{display:none!important}
  .main{overflow:visible}
  .app,.body,.main{height:auto;overflow:visible}
  .section-card{break-inside:avoid;page-break-inside:avoid}
  .canvas-wrap{padding:.5rem}
}

/* ── TOAST ───────────────────────────── */
#toast{position:fixed;bottom:1.5rem;right:1.5rem;background:var(--navy);color:#fff;padding:10px 18px;border-radius:8px;font-size:12px;font-weight:500;z-index:9999;opacity:0;transform:translateY(8px);transition:.25s;pointer-events:none}
#toast.show{opacity:1;transform:translateY(0)}

/* ── MODAL ───────────────────────────── */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:1000;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:var(--white);border-radius:12px;padding:1.5rem;width:360px;max-width:90vw}
.modal h3{font-size:16px;font-weight:600;margin-bottom:.5rem}
.modal p{font-size:13px;color:var(--text-muted);margin-bottom:1.25rem}
.modal-actions{display:flex;justify-content:flex-end;gap:8px}

@media(max-width:700px){
  .sidebar{display:none}
  .g2,.g3,.g4{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="app">

<!-- TOPBAR -->
<div class="topbar">
  <div class="brand">
    <div class="brand-dot"></div>
    <span class="brand-name">Canvas PDI</span>
    <span class="brand-tag">Refinamento de Projetos</span>
  </div>
  <div class="top-actions">
    <span class="save-indicator" id="save-badge">● salvo</span>
    <button class="btn btn-ghost-white btn-sm" onclick="exportPDF()">&#8659; PDF</button>
    <button class="btn btn-ghost-white btn-sm" onclick="exportPPTX()">&#8659; PPTX</button>
    <button class="btn btn-teal btn-sm" onclick="newProject()">+ Novo projeto</button>
  </div>
</div>

<div class="body">

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sidebar-top">
    <h3>Projetos</h3>
    <button class="new-btn" onclick="newProject()">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 1v12M1 7h12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
      Novo projeto
    </button>
  </div>
  <div class="project-list" id="project-list"></div>
  <div class="sidebar-footer">
    <div class="sidebar-stat" id="sidebar-stat">0 projetos</div>
  </div>
</div>

<!-- MAIN -->
<div class="main" id="main">
  <div class="empty-wrap" id="empty-state">
    <div class="empty-icon">&#128196;</div>
    <h2>Nenhum projeto selecionado</h2>
    <p>Crie um novo projeto para começar a preencher o canvas de refinamento de PDI.</p>
    <button class="btn btn-navy" onclick="newProject()">Criar primeiro projeto</button>
  </div>

  <div id="canvas-form" style="display:none">
  <div class="canvas-wrap">

    <div class="canvas-header">
      <div class="canvas-header-left">
        <div class="canvas-title" id="canvas-title">Novo projeto</div>
        <div class="canvas-subtitle" id="canvas-subtitle">Canvas de Refinamento · PDI</div>
      </div>
      <div class="canvas-header-actions">
        <button class="btn btn-outline btn-sm" onclick="exportPDF()">&#8659; PDF</button>
        <button class="btn btn-outline btn-sm" onclick="exportPPTX()">&#8659; PPTX</button>
      </div>
    </div>

    <!-- BLOCO 1 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--navy)">
        <div class="sec-num">1</div>
        <span class="sec-title">Visão geral do projeto</span>
      </div>
      <div class="sec-body">
        <div class="field"><label>Nome do projeto</label><input id="f-nome" placeholder="Ex.: Monitoramento inteligente de redes de distribuição" oninput="onInput()"></div>
        <div class="fg g3">
          <div class="field"><label>Parceiro(s)</label><input id="f-parceiro" placeholder="Ex.: Universidade X, Empresa Y" oninput="onInput()"></div>
          <div class="field"><label>Responsável interno</label><input id="f-responsavel" placeholder="Nome do gestor" oninput="onInput()"></div>
          <div class="field"><label>Área demandante</label><input id="f-area" placeholder="Ex.: Distribuição / Operações" oninput="onInput()"></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Linha estratégica / programa</label><input id="f-linha" placeholder="Ex.: P&D ANEEL — Eficiência energética" oninput="onInput()"></div>
          <div class="field"><label>Orçamento estimado (R$)</label><input id="f-orcamento" placeholder="Ex.: R$ 1.200.000" oninput="onInput()"></div>
        </div>
        <div class="field">
          <label>TRL atual — clique para selecionar</label>
          <div class="trl-row" id="trl-row"></div>
          <div class="trl-labels">
            <span class="trl-label">1–3 Pesquisa básica</span>
            <span class="trl-label">4–6 Desenvolvimento</span>
            <span class="trl-label">7–9 Implantação</span>
          </div>
        </div>
      </div>
    </div>

    <!-- BLOCO 2 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--blue)">
        <div class="sec-num">2</div>
        <span class="sec-title">Status do refinamento</span>
      </div>
      <div class="sec-body">
        <div class="field">
          <label>Fase atual</label>
          <div class="status-pills">
            <div class="spill" data-val="business" data-cls="s-blue" onclick="pickStatus(this)"><span class="spill-dot" style="background:#378ADD"></span>Aprofundamento de Business Case</div>
            <div class="spill" data-val="cronograma" data-cls="s-amber" onclick="pickStatus(this)"><span class="spill-dot" style="background:#D4890A"></span>Alinhamento de Cronograma</div>
            <div class="spill" data-val="escopo" data-cls="s-orange" onclick="pickStatus(this)"><span class="spill-dot" style="background:#C04E22"></span>Fechamento de Escopo Técnico</div>
            <div class="spill" data-val="pronto" data-cls="s-green" onclick="pickStatus(this)"><span class="spill-dot" style="background:#5A9E28"></span>Pronto para Submissão</div>
            <div class="spill" data-val="descontinuado" data-cls="s-red" onclick="pickStatus(this)"><span class="spill-dot" style="background:#C43A3A"></span>Descontinuado / Desclassificado</div>
          </div>
          <div class="progress-bar"><div class="progress-fill" id="progress" style="width:0%"></div></div>
          <div class="progress-lbl" id="progress-lbl"></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Data da última interação</label><input type="date" id="f-ultima" oninput="onInput()"></div>
          <div class="field"><label>Próxima ação / reunião</label><input type="date" id="f-proxima" oninput="onInput()"></div>
        </div>
      </div>
    </div>

    <!-- BLOCO 3 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--teal)">
        <div class="sec-num">3</div>
        <span class="sec-title">Escopo em refinamento</span>
      </div>
      <div class="sec-body">
        <div class="field"><label>Problema que resolve</label><textarea id="f-problema" placeholder="Descreva o problema central que o projeto busca resolver..."></textarea></div>
        <div class="fg g2">
          <div class="field"><label>Escopo técnico preliminar</label><textarea id="f-escopo-tec" placeholder="Tecnologias, metodologias e abordagens no escopo..."></textarea></div>
          <div class="field"><label>Principais entregáveis</label><textarea id="f-entregaveis" placeholder="Relatórios, protótipos, publicações..."></textarea></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Premissas consideradas</label><textarea id="f-premissas" placeholder="Quais premissas foram assumidas para viabilizar o escopo?"></textarea></div>
          <div class="field field-danger">
            <label>Fora do escopo — o que NÃO está incluído</label>
            <textarea id="f-fora" placeholder="Delimitar claramente evita ruído e retrabalho..."></textarea>
          </div>
        </div>
      </div>
    </div>

    <!-- BLOCO 4 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--purple)">
        <div class="sec-num">4</div>
        <span class="sec-title">Comportamento do parceiro</span>
      </div>
      <div class="sec-body">
        <table class="star-table" id="star-table"></table>
        <div>
          <label style="display:block;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--text-hint);margin-bottom:8px">Classificação geral</label>
          <div class="p-badges">
            <button class="pbadge pg" id="pb-forte" onclick="setPC('forte')">Forte</button>
            <button class="pbadge pa" id="pb-moderado" onclick="setPC('moderado')">Moderado</button>
            <button class="pbadge pr" id="pb-critico" onclick="setPC('critico')">Crítico</button>
          </div>
        </div>
        <div class="field"><label>Observações / evidências</label><textarea id="f-obs-parceiro" placeholder="Comportamentos relevantes, situações observadas..."></textarea></div>
      </div>
    </div>

    <!-- BLOCO 5 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--orange)">
        <div class="sec-num">5</div>
        <span class="sec-title">Desafios no refinamento</span>
      </div>
      <div class="sec-body">
        <div class="fg g2">
          <div class="field"><label>Gargalos técnicos</label><textarea id="f-gargalos"></textarea></div>
          <div class="field"><label>Problemas de alinhamento</label><textarea id="f-alinhamento"></textarea></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Riscos identificados</label><textarea id="f-riscos"></textarea></div>
          <div class="field"><label>Dependências externas</label><textarea id="f-dependencias"></textarea></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Pontos de atenção regulatórios (ANEEL)</label><textarea id="f-regulatorio"></textarea></div>
          <div class="field"><label>Plano de mitigação</label><textarea id="f-mitigacao"></textarea></div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Impacto geral dos desafios</label>
            <select id="f-impacto" onchange="onInput()">
              <option value="">Selecione...</option>
              <option>Alto</option><option>Médio</option><option>Baixo</option>
            </select>
          </div>
          <div class="field"><label>Próximos passos / responsáveis</label><input id="f-proximos-passos" oninput="onInput()"></div>
        </div>
      </div>
    </div>

    <!-- BLOCO 6 -->
    <div class="section-card">
      <div class="section-head" style="background:var(--red)">
        <div class="sec-num">6</div>
        <span class="sec-title">Desclassificação (se aplicável)</span>
      </div>
      <div class="sec-body">
        <div class="field">
          <label>Motivo principal</label>
          <div class="check-grid">
            <label class="check-item"><input type="checkbox" id="ck-maturidade" onchange="onInput()"><span>Baixa maturidade técnica</span></label>
            <label class="check-item"><input type="checkbox" id="ck-parceiro" onchange="onInput()"><span>Problemas com parceiro</span></label>
            <label class="check-item"><input type="checkbox" id="ck-impacto" onchange="onInput()"><span>Baixo impacto potencial</span></label>
            <label class="check-item"><input type="checkbox" id="ck-escopo" onchange="onInput()"><span>Escopo inconsistente</span></label>
            <label class="check-item"><input type="checkbox" id="ck-estrategia" onchange="onInput()"><span>Falta de aderência estratégica</span></label>
            <label class="check-item"><input type="checkbox" id="ck-outros" onchange="onInput()"><span>Outros</span></label>
          </div>
        </div>
        <div class="fg g2">
          <div class="field"><label>Descrição objetiva</label><textarea id="f-desc-desc"></textarea></div>
          <div>
            <div class="field" style="margin-bottom:.75rem"><label>Em qual etapa caiu</label><input id="f-etapa-caiu" oninput="onInput()"></div>
            <div class="fg g2">
              <div class="field"><label>Pode retornar?</label>
                <select id="f-retorno" onchange="onInput()"><option value="">—</option><option>Sim</option><option>Não</option></select>
              </div>
              <div class="field"><label>Condição</label><input id="f-condicao" oninput="onInput()"></div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- BLOCO 7 -->
    <div class="section-card">
      <div class="section-head" style="background:#0A7A60">
        <div class="sec-num">7</div>
        <span class="sec-title">Recomendação final</span>
      </div>
      <div class="sec-body">
        <div class="field">
          <label>Decisão</label>
          <div class="rec-opts">
            <div class="rec-opt" data-val="avancar" data-cls="rt" onclick="pickRec(this)"><span class="rec-radio"></span>Avançar para estruturação formal</div>
            <div class="rec-opt" data-val="continuar" data-cls="rg" onclick="pickRec(this)"><span class="rec-radio"></span>Continuar refinamento</div>
            <div class="rec-opt" data-val="pausar" data-cls="ra" onclick="pickRec(this)"><span class="rec-radio"></span>Pausar</div>
            <div class="rec-opt" data-val="descontinuar" data-cls="rr" onclick="pickRec(this)"><span class="rec-radio"></span>Descontinuar</div>
          </div>
        </div>
        <div class="field"><label>Justificativa executiva (3–5 linhas)</label><textarea id="f-justificativa" rows="4" placeholder="Explique a recomendação e seu embasamento..."></textarea></div>
        <div class="fg g3">
          <div class="field"><label>Revisado por</label><input id="f-revisor" oninput="onInput()"></div>
          <div class="field"><label>Data da revisão</label><input type="date" id="f-data-rev" oninput="onInput()"></div>
          <div class="field"><label>Comentários / ressalvas</label><input id="f-comentarios" oninput="onInput()"></div>
        </div>
      </div>
    </div>

    <div class="form-footer">
      <button class="btn btn-danger btn-sm" onclick="confirmDelete()">Excluir projeto</button>
      <div style="display:flex;gap:8px">
        <button class="btn btn-outline btn-sm" onclick="exportPDF()">&#8659; Imprimir / PDF</button>
        <button class="btn btn-outline btn-sm" onclick="exportPPTX()">&#8659; Exportar PPTX</button>
        <button class="btn btn-navy btn-sm" onclick="forceSave()">Salvar</button>
      </div>
    </div>

  </div>
  </div>
</div>
</div><!-- body -->
</div><!-- app -->

<!-- DELETE MODAL -->
<div class="modal-bg" id="del-modal">
  <div class="modal">
    <h3>Excluir projeto?</h3>
    <p>Esta ação não pode ser desfeita. O projeto e todos os dados serão removidos permanentemente.</p>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeModal()">Cancelar</button>
      <button class="btn btn-danger" onclick="deleteProject()">Excluir</button>
    </div>
  </div>
</div>

<div id="toast"></div>

<script>
// ── STATE ───────────────────────────────────────────────────────
let DB = {};
let CID = null;
let saveTimer = null;

const STAR_CRITERIA = [
  'Dedicação / Disponibilidade',
  'Qualidade técnica das entregas',
  'Velocidade de resposta',
  'Proatividade',
  'Aderência ao desafio',
];
const STATUS_MAP = {
  business:      {pct:25,  lbl:'Fase 1/4 — Aprofundamento de Business Case', dot:'#378ADD'},
  cronograma:    {pct:50,  lbl:'Fase 2/4 — Alinhamento de Cronograma',       dot:'#D4890A'},
  escopo:        {pct:75,  lbl:'Fase 3/4 — Fechamento de Escopo Técnico',    dot:'#C04E22'},
  pronto:        {pct:100, lbl:'Fase 4/4 — Pronto para Submissão',           dot:'#5A9E28'},
  descontinuado: {pct:0,   lbl:'Descontinuado / Desclassificado',            dot:'#C43A3A'},
};

// ── INIT ──────────────────────────────────────────────────────────
function init() {
  buildTRL();
  buildStars();
  bindTextareas();
  loadDB();
  renderSidebar();
  const ids = Object.keys(DB);
  if (ids.length) openProject(ids[0]);
}

function buildTRL() {
  const row = document.getElementById('trl-row');
  for (let i=1;i<=9;i++) {
    const b = document.createElement('button');
    b.className = 'trl-btn';
    b.textContent = i;
    b.dataset.v = i;
    b.onclick = () => { setTRL(i); schedSave(); };
    row.appendChild(b);
  }
}

function buildStars() {
  const t = document.getElementById('star-table');
  STAR_CRITERIA.forEach((label,ci) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="star-label" style="width:60%">${label}</td><td><div class="stars" id="s${ci}"></div></td>`;
    t.appendChild(tr);
    const wrap = tr.querySelector(`#s${ci}`);
    for (let i=1;i<=5;i++) {
      const s = document.createElement('div');
      s.className = 'star';
      s.textContent = '★';
      s.dataset.v = i;
      s.onclick = () => { setStars(ci,i); schedSave(); };
      wrap.appendChild(s);
    }
  });
}

function bindTextareas() {
  document.querySelectorAll('textarea').forEach(ta => ta.addEventListener('input', schedSave));
}

// ── PROJECTS ──────────────────────────────────────────────────────
function newProject() {
  const id = 'p'+Date.now();
  DB[id] = {id, nome:'Novo projeto', created: new Date().toLocaleDateString('pt-BR'), stars:{}, checks:{}, status:null, rec:null, trl:0, pc:null};
  saveDB();
  renderSidebar();
  openProject(id);
  setTimeout(()=>{document.getElementById('f-nome').focus();document.getElementById('f-nome').select();},80);
}

function openProject(id) {
  CID = id;
  document.getElementById('empty-state').style.display = 'none';
  document.getElementById('canvas-form').style.display = 'block';
  loadForm();
  renderSidebar();
}

function confirmDelete() { document.getElementById('del-modal').classList.add('open'); }
function closeModal() { document.getElementById('del-modal').classList.remove('open'); }

function deleteProject() {
  closeModal();
  if (!CID) return;
  delete DB[CID];
  CID = null;
  saveDB();
  renderSidebar();
  const ids = Object.keys(DB);
  if (ids.length) openProject(ids[0]);
  else {
    document.getElementById('empty-state').style.display = 'flex';
    document.getElementById('canvas-form').style.display = 'none';
  }
}

// ── FORM LOAD / SAVE ──────────────────────────────────────────────
const FIELD_MAP = {
  'f-nome':'nome','f-parceiro':'parceiro','f-responsavel':'responsavel','f-area':'area',
  'f-linha':'linha','f-orcamento':'orcamento','f-ultima':'ultima','f-proxima':'proxima',
  'f-problema':'problema','f-escopo-tec':'escopo_tec','f-entregaveis':'entregaveis',
  'f-premissas':'premissas','f-fora':'fora','f-obs-parceiro':'obs_parceiro',
  'f-gargalos':'gargalos','f-alinhamento':'alinhamento','f-riscos':'riscos',
  'f-dependencias':'dependencias','f-regulatorio':'regulatorio','f-mitigacao':'mitigacao',
  'f-impacto':'impacto','f-proximos-passos':'proximos_passos',
  'f-desc-desc':'desc_desc','f-etapa-caiu':'etapa_caiu','f-retorno':'retorno','f-condicao':'condicao',
  'f-justificativa':'justificativa','f-revisor':'revisor','f-data-rev':'data_rev','f-comentarios':'comentarios'
};
const CHECKS = ['ck-maturidade','ck-parceiro','ck-impacto','ck-escopo','ck-estrategia','ck-outros'];

function loadForm() {
  const p = DB[CID]||{};
  Object.entries(FIELD_MAP).forEach(([fid,key])=>{ const el=document.getElementById(fid); if(el) el.value=p[key]||''; });
  setTRL(p.trl||0, true);
  STAR_CRITERIA.forEach((_,ci) => setStars(ci, (p.stars||{})[ci]||0, true));
  setPC(p.pc||null, true);
  pickStatus(null, p.status||null, true);
  pickRec(null, p.rec||null, true);
  CHECKS.forEach(id=>{ const el=document.getElementById(id); if(el) el.checked=!!(p.checks&&p.checks[id]); });
  document.getElementById('canvas-title').textContent = p.nome||'Novo projeto';
  document.getElementById('canvas-subtitle').textContent = `Canvas de Refinamento · PDI · Criado em ${p.created||'—'}`;
}

function collectForm() {
  const p = {};
  Object.entries(FIELD_MAP).forEach(([fid,key])=>{ const el=document.getElementById(fid); if(el) p[key]=el.value; });
  const checks = {};
  CHECKS.forEach(id=>{ const el=document.getElementById(id); if(el) checks[id]=el.checked; });
  p.checks = checks;
  return p;
}

function onInput() { schedSave(); }

function schedSave() {
  const badge = document.getElementById('save-badge');
  badge.textContent = '○ salvando...';
  badge.className = 'save-indicator saving';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(()=>{
    if (!CID) return;
    Object.assign(DB[CID], collectForm());
    const nome = document.getElementById('f-nome').value||'Novo projeto';
    DB[CID].nome = nome;
    document.getElementById('canvas-title').textContent = nome;
    saveDB();
    renderSidebar();
    badge.textContent = '● salvo';
    badge.className = 'save-indicator saved';
    setTimeout(()=>{ badge.textContent='● salvo'; badge.className='save-indicator'; },2000);
  }, 700);
}

function forceSave() {
  clearTimeout(saveTimer);
  if (!CID) return;
  Object.assign(DB[CID], collectForm());
  DB[CID].nome = document.getElementById('f-nome').value||'Novo projeto';
  saveDB();
  renderSidebar();
  toast('Projeto salvo com sucesso!');
}

// ── CONTROLS ──────────────────────────────────────────────────────
function setTRL(val, silent) {
  if (!silent && CID) DB[CID].trl = val;
  document.querySelectorAll('.trl-btn').forEach(b=>b.classList.toggle('active', parseInt(b.dataset.v)<=val));
}

function setStars(ci, val, silent) {
  if (!silent && CID) { if(!DB[CID].stars)DB[CID].stars={}; DB[CID].stars[ci]=val; }
  document.querySelectorAll(`#s${ci} .star`).forEach(s=>s.classList.toggle('on', parseInt(s.dataset.v)<=val));
}

function setPC(val, silent) {
  if (!silent && CID) DB[CID].pc=val;
  ['forte','moderado','critico'].forEach(k=>{
    const el=document.getElementById('pb-'+k);
    if(el){ el.className=`pbadge ${k==='forte'?'pg':k==='moderado'?'pa':'pr'}${val===k?' sel':''}`; }
  });
  if (!silent) schedSave();
}

function pickStatus(el, val, silent) {
  const v = val||(el&&el.dataset.val);
  document.querySelectorAll('.spill').forEach(p=>{ p.className='spill'; if(p.dataset.val===v) p.classList.add(p.dataset.cls); });
  const m = STATUS_MAP[v]||{pct:0,lbl:''};
  document.getElementById('progress').style.width = m.pct+'%';
  document.getElementById('progress-lbl').textContent = m.lbl;
  if (!silent && CID && el) { DB[CID].status=v; schedSave(); }
}

function pickRec(el, val, silent) {
  const v = val||(el&&el.dataset.val);
  document.querySelectorAll('.rec-opt').forEach(o=>{ o.className='rec-opt'; if(o.dataset.val===v) o.classList.add(o.dataset.cls); });
  if (!silent && CID && el) { DB[CID].rec=v; schedSave(); }
}

// ── SIDEBAR ───────────────────────────────────────────────────────
function renderSidebar() {
  const list = document.getElementById('project-list');
  list.innerHTML='';
  const ids = Object.keys(DB);
  ids.forEach(id=>{
    const p = DB[id];
    const div = document.createElement('div');
    div.className = 'project-card'+(id===CID?' active':'');
    const dot = p.status&&STATUS_MAP[p.status] ? `<span class="status-dot" style="background:${STATUS_MAP[p.status].dot}"></span>` : '';
    div.innerHTML=`<div class="pc-name">${p.nome||'Sem título'}</div><div class="pc-meta">${dot}<span class="pc-date">${p.created||''}</span></div>`;
    div.onclick = ()=>openProject(id);
    list.appendChild(div);
  });
  document.getElementById('sidebar-stat').textContent = `${ids.length} projeto${ids.length!==1?'s':''}`;
}

// ── PERSISTENCE ───────────────────────────────────────────────────
function saveDB() { try { localStorage.setItem('pdi_canvas_v1', JSON.stringify(DB)); } catch(e){} }

function loadDB() {
  try {
    const raw = localStorage.getItem('pdi_canvas_v1');
    if (raw) { DB = JSON.parse(raw); return; }
  } catch(e){}
  const demo = 'p_demo';
  DB[demo] = {id:demo, nome:'Projeto demonstração', created:new Date().toLocaleDateString('pt-BR'), stars:{}, checks:{}, status:'business', rec:null, trl:3, pc:'moderado'};
}

// ── EXPORT PDF ────────────────────────────────────────────────────
function exportPDF() {
  toast('Abrindo impressão... Selecione "Salvar como PDF" no diálogo.');
  setTimeout(()=>window.print(), 300);
}

// ── EXPORT PPTX ───────────────────────────────────────────────────
function exportPPTX() {
  if (!CID) { toast('Selecione um projeto primeiro.'); return; }
  toast('Gerando PPTX...');
  const p = DB[CID];

  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';

  const C = {
    navy:'1B3358',blue:'2D6A9F',blueL:'D6E8F7',teal:'0C6B52',tealL:'D2EDE5',tealM:'1A9970',
    amber:'B57218',amberL:'FAF0D7',red:'9E2B2B',redL:'FAE8E8',redM:'C43A3A',
    green:'3A6B10',greenL:'EAF3DE',orange:'8F3818',orangeL:'FAECE7',orangeM:'C04E22',
    purple:'4E45AE',purpleL:'ECEAF9',white:'FFFFFF',border:'CCCCCC',
    text:'1C1B1A',muted:'6B6965',bg:'F5F3ED',
  };

  function hdr(sl, x,y,w,h, num,label, fill) {
    sl.addShape(pres.ShapeType.rect, {x,y,w,h,fill:{color:fill},line:{color:fill}});
    sl.addShape(pres.ShapeType.rect, {x,y,w:0.24,h,fill:{color:'FFFFFF',transparency:82},line:{color:fill}});
    sl.addText(num,  {x:x+0.01,y,w:0.24,h, fontSize:9,bold:true,color:C.white,align:'center',valign:'middle',margin:0});
    sl.addText(label,{x:x+0.28,y,w:w-0.32,h, fontSize:8,bold:true,color:C.white,valign:'middle',charSpacing:1,margin:0});
  }
  function fld(sl, x,y,w,h, label, opts={}) {
    const lh=0.16;
    sl.addShape(pres.ShapeType.rect,{x,y,w,h:lh,fill:{color:opts.lbg||'E2DFD8'},line:{color:C.border,pt:0.4}});
    sl.addText(label,{x:x+0.05,y,w:w-0.1,h:lh,fontSize:5.8,bold:true,color:opts.lc||C.muted,valign:'middle',margin:0});
    sl.addShape(pres.ShapeType.rect,{x,y:y+lh,w,h:h-lh,fill:{color:opts.vbg||C.white},line:{color:opts.vbc||C.border,pt:0.4}});
    if(opts.val) sl.addText(String(opts.val),{x:x+0.08,y:y+lh+0.02,w:w-0.16,h:h-lh-0.04,fontSize:7,color:C.text,margin:0});
    if(opts.accent) sl.addShape(pres.ShapeType.rect,{x,y:y+lh,w:0.05,h:h-lh,fill:{color:opts.accent},line:{color:opts.accent}});
  }
  function topBar(sl, pg, total) {
    sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:13.3,h:0.56,fill:{color:C.navy},line:{color:C.navy}});
    sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.07,h:0.56,fill:{color:C.tealM},line:{color:C.tealM}});
    sl.addText(p.nome||'Projeto PDI',{x:0.2,y:0,w:9,h:0.56,fontSize:14,bold:true,color:C.white,valign:'middle',margin:0});
    sl.addText(`Canvas de Refinamento · PDI`,{x:9.5,y:0,w:3.2,h:0.56,fontSize:8,color:'7FA8CC',align:'right',valign:'middle',margin:0});
    sl.addShape(pres.ShapeType.rect,{x:12.8,y:0.15,w:0.4,h:0.26,fill:{color:'243F6A'},line:{color:'4A8DC0',pt:0.5}});
    sl.addText(`${pg}/${total}`,{x:12.8,y:0.15,w:0.4,h:0.26,fontSize:8,bold:true,color:'D6E8F7',align:'center',valign:'middle',margin:0});
    sl.addShape(pres.ShapeType.rect,{x:0,y:7.38,w:13.3,h:0.12,fill:{color:C.navy},line:{color:C.navy}});
    sl.addText('PDI — Canvas de Refinamento de Projetos · Instrumento de Governança de Inovação',{x:0.2,y:7.38,w:13.1,h:0.12,fontSize:5.5,color:'7FA8CC',valign:'middle',margin:0});
  }

  function v(key) { return p[key]||''; }

  // ── SLIDE 1 ─────────────────────────────────────────────────────
  const s1 = pres.addSlide();
  s1.background = {color:C.bg};
  topBar(s1,1,2);
  const TOP=0.68, M=0.18;

  // Bloco 1
  hdr(s1, M,TOP,4.5,0.26, '1','VISÃO GERAL DO PROJETO',C.navy);
  fld(s1, M,TOP+0.26,4.5,0.42, 'NOME DO PROJETO',{val:v('nome')});
  fld(s1, M,TOP+0.70,4.5*0.55,0.38, 'PARCEIRO(S)',{val:v('parceiro')});
  fld(s1, M+4.5*0.56,TOP+0.70,4.5*0.44,0.38, 'RESPONSÁVEL INTERNO',{val:v('responsavel')});
  fld(s1, M,TOP+1.10,4.5*0.42,0.36, 'ÁREA DEMANDANTE',{val:v('area')});
  fld(s1, M+4.5*0.43,TOP+1.10,4.5*0.34,0.36, 'LINHA ESTRATÉGICA',{val:v('linha')});
  fld(s1, M+4.5*0.78,TOP+1.10,4.5*0.22,0.36, 'ORÇAMENTO (R$)',{val:v('orcamento')});
  // TRL
  const trlY=TOP+1.48;
  fld(s1, M,trlY,4.5,0.46,'TRL ATUAL');
  for(let i=1;i<=9;i++) {
    const on=(p.trl||0)>=i;
    s1.addShape(pres.ShapeType.rect,{x:M+0.05+(i-1)*0.47,y:trlY+0.18,w:0.42,h:0.14,fill:{color:on?'4A8DC0':'D6D4CE'},line:{color:on?'2D6A9F':C.border,pt:0.4}});
    s1.addText(String(i),{x:M+0.05+(i-1)*0.47,y:trlY+0.18,w:0.42,h:0.14,fontSize:6,color:on?C.white:'888580',align:'center',valign:'middle',bold:on,margin:0});
  }
  s1.addText('1–3 Pesquisa    4–6 Desenvolvimento    7–9 Implantação',{x:M+0.05,y:trlY+0.34,w:4.4,h:0.1,fontSize:5.2,color:C.muted,margin:0});

  // Bloco 2
  const b2y=trlY+0.56;
  hdr(s1, M,b2y,4.5,0.26,'2','STATUS DO REFINAMENTO',C.blue);
  const statLabels={business:'Aprofundamento de Business Case',cronograma:'Alinhamento de Cronograma',escopo:'Fechamento de Escopo Técnico',pronto:'Pronto para Submissão',descontinuado:'Descontinuado / Desclassificado'};
  const statDots={business:'378ADD',cronograma:'D4890A',escopo:'C04E22',pronto:'5A9E28',descontinuado:'C43A3A'};
  const statBgs={business:'D6E8F7',cronograma:'FAF0D7',escopo:'FAECE7',pronto:'EAF3DE',descontinuado:'FCEBEB'};
  const pillH=0.28, pillGap=0.05;
  Object.keys(statLabels).forEach((k,i)=>{
    const active=p.status===k;
    const bg=active?statBgs[k]:'EDEBE6';
    const bc=active?statDots[k]:C.border;
    s1.addShape(pres.ShapeType.rect,{x:M,y:b2y+0.26+i*(pillH+pillGap),w:4.5,h:pillH,fill:{color:bg},line:{color:bc,pt:active?1.2:0.4}});
    s1.addShape(pres.ShapeType.oval,{x:M+0.1,y:b2y+0.26+i*(pillH+pillGap)+(pillH-0.12)/2,w:0.12,h:0.12,fill:{color:statDots[k]},line:{color:statDots[k]}});
    s1.addText(statLabels[k],{x:M+0.28,y:b2y+0.26+i*(pillH+pillGap),w:4.2,h:pillH,fontSize:7,bold:active,color:active?C.text:C.muted,valign:'middle',margin:0});
  });
  const pbY=b2y+0.26+5*(pillH+pillGap)+0.04;
  const pct=(STATUS_MAP[p.status]||{pct:0}).pct;
  s1.addShape(pres.ShapeType.rect,{x:M,y:pbY,w:4.5,h:0.08,fill:{color:'D6D4CE'},line:{color:C.border,pt:0.3}});
  if(pct>0) s1.addShape(pres.ShapeType.rect,{x:M,y:pbY,w:4.5*(pct/100),h:0.08,fill:{color:'4A8DC0'},line:{color:'4A8DC0'}});
  fld(s1, M,pbY+0.14,4.5*0.50,0.36,'ÚLTIMA INTERAÇÃO',{val:v('ultima')});
  fld(s1, M+4.5*0.51,pbY+0.14,4.5*0.49,0.36,'PRÓXIMA AÇÃO',{val:v('proxima')});

  // Bloco 3
  const b3x=4.86, b3w=8.26;
  hdr(s1, b3x,TOP,b3w,0.26,'3','ESCOPO EM REFINAMENTO',C.teal);
  fld(s1, b3x,TOP+0.26,b3w,0.44,'PROBLEMA QUE RESOLVE',{val:v('problema')});
  fld(s1, b3x,TOP+0.72,b3w*0.5,0.44,'ESCOPO TÉCNICO PRELIMINAR',{val:v('escopo_tec')});
  fld(s1, b3x+b3w*0.51,TOP+0.72,b3w*0.49,0.44,'PRINCIPAIS ENTREGÁVEIS',{val:v('entregaveis')});
  fld(s1, b3x,TOP+1.18,b3w*0.5,0.44,'PREMISSAS CONSIDERADAS',{val:v('premissas')});
  // Fora do escopo
  const foeX=b3x+b3w*0.51, foeY=TOP+1.18, foeW=b3w*0.49;
  s1.addShape(pres.ShapeType.rect,{x:foeX,y:foeY,w:foeW,h:0.16,fill:{color:'F5D8D8'},line:{color:'D49090',pt:0.4}});
  s1.addText('FORA DO ESCOPO',{x:foeX+0.05,y:foeY,w:foeW-0.1,h:0.16,fontSize:5.8,bold:true,color:C.red,valign:'middle',margin:0});
  s1.addShape(pres.ShapeType.rect,{x:foeX,y:foeY+0.16,w:foeW,h:0.28,fill:{color:'FEF6F6'},line:{color:'D49090',pt:0.4}});
  s1.addShape(pres.ShapeType.rect,{x:foeX,y:foeY+0.16,w:0.05,h:0.28,fill:{color:C.redM},line:{color:C.redM}});
  if(v('fora')) s1.addText(v('fora'),{x:foeX+0.09,y:foeY+0.18,w:foeW-0.15,h:0.25,fontSize:6.5,color:C.text,margin:0});

  fld(s1, b3x,TOP+1.64,b3w,0.42,'ENTREGÁVEIS-CHAVE / MARCOS',{val:v('entregaveis')});

  // ── SLIDE 2 ─────────────────────────────────────────────────────
  const s2 = pres.addSlide();
  s2.background = {color:C.bg};
  topBar(s2,2,2);
  const T2=0.68;
  const cA=M, wA=4.14, cB=4.42, wB=4.14, cC=8.68, wC=4.44;

  // Bloco 4
  hdr(s2, cA,T2,wA,0.26,'4','COMPORTAMENTO DO PARCEIRO',C.purple);
  STAR_CRITERIA.forEach((lbl,ci)=>{
    const ry=T2+0.28+0.05+ci*0.24;
    if(ci%2===0) s2.addShape(pres.ShapeType.rect,{x:cA,y:ry,w:wA,h:0.23,fill:{color:'EEEAE4'},line:{color:'EEEAE4'}});
    s2.addText(lbl,{x:cA+0.06,y:ry+0.04,w:wA*0.55,h:0.16,fontSize:6.5,color:C.muted,margin:0});
    const sv=(p.stars||{})[ci]||0;
    for(let i=1;i<=5;i++) {
      s2.addShape(pres.ShapeType.rect,{x:cA+wA*0.56+(i-1)*0.17,y:ry+0.05,w:0.14,h:0.14,fill:{color:i<=sv?'FAF0D7':'E0DDD6'},line:{color:i<=sv?'D4890A':C.border,pt:0.4}});
      if(i<=sv) s2.addText('★',{x:cA+wA*0.56+(i-1)*0.17,y:ry+0.04,w:0.14,h:0.16,fontSize:7,color:C.amber,align:'center',valign:'middle',margin:0});
    }
  });
  const divY4=T2+0.28+0.05+STAR_CRITERIA.length*0.24+0.06;
  s2.addShape(pres.ShapeType.line,{x:cA+0.1,y:divY4,w:wA-0.2,h:0,line:{color:C.border,pt:0.5}});
  s2.addText('Classificação:',{x:cA+0.06,y:divY4+0.06,w:1.4,h:0.2,fontSize:6.5,color:C.muted,valign:'middle',margin:0});
  const pcMap={forte:{bg:'EAF3DE',tc:C.green,bc:'5A9E28'},moderado:{bg:'FAF0D7',tc:C.amber,bc:'D4890A'},critico:{bg:'FCEBEB',tc:C.red,bc:'C43A3A'}};
  Object.entries(pcMap).forEach(([k,cfg],i)=>{
    const active=p.pc===k;
    s2.addShape(pres.ShapeType.rect,{x:cA+1.5+i*0.88,y:divY4+0.06,w:0.82,h:0.22,fill:{color:active?cfg.bg:'EDEBE6'},line:{color:active?cfg.bc:C.border,pt:active?1.2:0.4}});
    s2.addText(k.charAt(0).toUpperCase()+k.slice(1),{x:cA+1.5+i*0.88,y:divY4+0.06,w:0.82,h:0.22,fontSize:7,bold:active,color:active?cfg.tc:C.muted,align:'center',valign:'middle',margin:0});
  });
  fld(s2, cA,divY4+0.32,wA,0.50,'OBSERVAÇÕES / EVIDÊNCIAS',{val:v('obs_parceiro')});

  // Bloco 5
  hdr(s2, cB,T2,wB,0.26,'5','DESAFIOS NO REFINAMENTO',C.orange);
  const riskItems=[
    {lbl:'Gargalos técnicos',val:v('gargalos'),impL:'ALTO',ibg:'FCEBEB',itc:C.redM},
    {lbl:'Problemas de alinhamento',val:v('alinhamento'),impL:'MÉDIO',ibg:'FAF0D7',itc:C.amber},
    {lbl:'Riscos identificados',val:v('riscos'),impL:'ALTO',ibg:'FCEBEB',itc:C.redM},
    {lbl:'Dependências externas',val:v('dependencias'),impL:'MÉDIO',ibg:'FAF0D7',itc:C.amber},
    {lbl:'Atenção reg. ANEEL',val:v('regulatorio'),impL:'BAIXO',ibg:'EAF3DE',itc:C.green},
  ];
  riskItems.forEach((r,i)=>{
    const ry=T2+0.28+0.06+i*0.30;
    s2.addShape(pres.ShapeType.rect,{x:cB,y:ry,w:wB-0.72,h:0.26,fill:{color:C.white},line:{color:C.border,pt:0.4}});
    s2.addText(r.val||r.lbl,{x:cB+0.06,y:ry,w:wB-0.82,h:0.26,fontSize:6.5,color:r.val?C.text:'AAAAAA',valign:'middle',italic:!r.val,margin:0});
    s2.addShape(pres.ShapeType.rect,{x:cB+wB-0.68,y:ry+0.04,w:0.62,h:0.18,fill:{color:r.ibg},line:{color:r.ibg}});
    s2.addText(r.impL,{x:cB+wB-0.68,y:ry+0.04,w:0.62,h:0.18,fontSize:6,bold:true,color:r.itc,align:'center',valign:'middle',margin:0});
  });
  const mitY=T2+0.28+0.06+riskItems.length*0.30+0.06;
  fld(s2, cB,mitY,wB,0.38,'PLANO DE MITIGAÇÃO',{val:v('mitigacao')});
  fld(s2, cB,mitY+0.40,wB,0.30,'PRÓXIMOS PASSOS',{val:v('proximos_passos')});

  // Bloco 6
  hdr(s2, cC,T2,wC,0.26,'6','DESCLASSIFICAÇÃO (SE APLICÁVEL)',C.red);
  const motivos=[['Baixa maturidade técnica','ck-maturidade'],['Baixo impacto potencial','ck-impacto'],['Falta de aderência estratégica','ck-estrategia'],['Problemas com parceiro','ck-parceiro'],['Escopo inconsistente','ck-escopo'],['Outros','ck-outros']];
  const mw6=wC/2-0.06;
  motivos.forEach(([lbl,ck],i)=>{
    const x=cC+(i<3?0:mw6+0.12), y=T2+0.28+0.08+(i%3)*0.22;
    const checked=!!(p.checks&&p.checks[ck]);
    s2.addShape(pres.ShapeType.rect,{x:x+0.05,y:y+0.04,w:0.12,h:0.12,fill:{color:checked?'D2EDE5':C.white},line:{color:checked?C.tealM:C.border,pt:0.5}});
    if(checked) s2.addText('✓',{x:x+0.05,y:y+0.03,w:0.12,h:0.14,fontSize:7,bold:true,color:C.teal,align:'center',valign:'middle',margin:0});
    s2.addText(lbl,{x:x+0.21,y,w:mw6-0.22,h:0.20,fontSize:6.5,color:C.text,valign:'middle',margin:0});
  });
  const div6y=T2+0.28+0.08+3*0.22+0.06;
  s2.addShape(pres.ShapeType.line,{x:cC+0.1,y:div6y,w:wC-0.2,h:0,line:{color:C.border,pt:0.5}});
  fld(s2, cC,div6y+0.06,wC,0.34,'DESCRIÇÃO OBJETIVA',{val:v('desc_desc')});
  fld(s2, cC,div6y+0.42,wC*0.50,0.26,'ETAPA EM QUE CAIU',{val:v('etapa_caiu')});
  fld(s2, cC+wC*0.52,div6y+0.42,wC*0.48,0.26,'PODE RETORNAR?',{val:v('retorno')});
  fld(s2, cC,div6y+0.70,wC,0.26,'CONDIÇÃO PARA RETORNO',{val:v('condicao')});

  // Bloco 7 — full width
  const b7tops=[T2+0.28+0.05+STAR_CRITERIA.length*0.24+0.64, mitY+0.70, div6y+0.98];
  const b7y=Math.max(...b7tops)+0.14;
  hdr(s2, M,b7y,13.1,0.26,'7','RECOMENDAÇÃO FINAL','0A7A60');
  const recOpts=[{val:'avancar',lbl:'Avançar para estruturação formal',bg:'D2EDE5',bc:C.tealM},{val:'continuar',lbl:'Continuar refinamento',bg:'EAF3DE',bc:C.green},{val:'pausar',lbl:'Pausar',bg:'FAF0D7',bc:C.amber},{val:'descontinuar',lbl:'Descontinuar',bg:'FCEBEB',bc:C.redM}];
  const rw=2.60, rh=0.40, rg=0.10;
  recOpts.forEach((r,i)=>{
    const active=p.rec===r.val;
    s2.addShape(pres.ShapeType.rect,{x:M+i*(rw+rg),y:b7y+0.28,w:rw,h:rh,fill:{color:active?r.bg:'EDEBE6'},line:{color:active?r.bc:C.border,pt:active?1.5:0.4}});
    s2.addShape(pres.ShapeType.oval,{x:M+i*(rw+rg)+0.1,y:b7y+0.28+(rh-0.14)/2,w:0.14,h:0.14,fill:{color:active?r.bc:C.white},line:{color:active?r.bc:C.border,pt:0.8}});
    s2.addText(r.lbl,{x:M+i*(rw+rg)+0.30,y:b7y+0.28,w:rw-0.35,h:rh,fontSize:7.5,bold:active,color:active?C.text:C.muted,valign:'middle',margin:0});
  });
  const justX=M+recOpts.length*(rw+rg)+0.08;
  fld(s2, justX,b7y+0.28,13.1-justX+M,rh,'JUSTIFICATIVA EXECUTIVA',{val:v('justificativa')});
  const sigY=b7y+0.28+rh+0.08;
  fld(s2, M,sigY,3.6,0.30,'REVISADO POR',{val:v('revisor')});
  fld(s2, M+3.7,sigY,2.0,0.30,'DATA',{val:v('data_rev')});
  fld(s2, M+5.8,sigY,7.3,0.30,'COMENTÁRIOS / RESSALVAS',{val:v('comentarios')});

  const fileName = (p.nome||'projeto_pdi').replace(/[^a-zA-Z0-9_\-]/g,'_')+'_canvas.pptx';
  pres.writeFile({fileName}).then(()=>toast(`PPTX "${fileName}" gerado com sucesso!`)).catch(e=>toast('Erro ao gerar PPTX: '+e));
}

// ── UTILS ─────────────────────────────────────────────────────────
function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(()=>t.classList.remove('show'), 3500);
}

init();
</script>
</body>
</html>
