# -*- coding: utf-8 -*-
import os
import io
import sqlite3
import webbrowser
from threading import Timer
from datetime import date, datetime

from flask import (
    Flask,
    render_template_string,
    request,
    session,
    redirect,
    url_for,
    send_file,
)
from openpyxl import Workbook

app = Flask(__name__)
# SE ESTIVER LENDO ISSO: ESTA É A VERSÃO CORRETA (V3.13) COM PEDIATRIA
app.secret_key = "oftalmoprev_bir_v313_pediatria"

# ==================================================
# 0. LOGIN (USUÁRIO / SENHA)
# ==================================================
USERS = {
    "admin": "admin123",
    "operador": "1234",
}

LOGIN_TEMPLATE = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Login - OftalmoPrev</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { background:#f4f7f9; font-family:'Segoe UI',sans-serif; }
    .card { border:none; border-radius:14px; box-shadow:0 8px 22px rgba(0,0,0,.10); }
    .header-top { background:#002b5c; color:#fff; padding:18px 0; border-bottom:5px solid #007bff; }
  </style>
</head>
<body>
<div class="header-top text-center">
  <h4 class="mb-0">OFTALMOPREV - ACESSO RESTRITO</h4>
  <div class="small">Build de Integração Restrita (BIR v3.13)</div>
</div>

<div class="container py-5" style="max-width: 520px;">
  <div class="card p-4">
    <h5 class="fw-bold mb-3">Entrar</h5>

    {% if erro %}
      <div class="alert alert-danger small">{{ erro }}</div>
    {% endif %}

    <form method="post">
      <div class="mb-3">
        <label class="small fw-bold">Usuário</label>
        <input class="form-control" name="username" required autofocus>
      </div>
      <div class="mb-3">
        <label class="small fw-bold">Senha</label>
        <input type="password" class="form-control" name="password" required>
      </div>
      <button class="btn btn-primary w-100 py-2 fw-bold">Entrar</button>
    </form>

    <div class="text-secondary small mt-3">
      * Uso restrito. Seus acessos podem ser registrados para auditoria.
    </div>
  </div>
</div>
</body>
</html>
"""

@app.before_request
def exigir_login():
    # libera rotas públicas
    if request.endpoint in ("login", "static"):
        return
    # libera logout (pra evitar loop caso alguém abra /logout sem estar logado)
    if request.endpoint == "logout":
        return
    # protege TODO o resto
    if not session.get("user"):
        return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    erro = ""
    if request.method == "POST":
        u = (request.form.get("username") or "").strip()
        p = request.form.get("password") or ""
        if u in USERS and USERS[u] == p:
            session["user"] = u
            return redirect(url_for("index"))
        erro = "Usuário ou senha inválidos."
    return render_template_string(LOGIN_TEMPLATE, erro=erro)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ==================================================
# 0.1 PACIENTES (CADASTRO: NOME + TELEFONE + NASC)
# ==================================================
PACIENTES_DB = os.path.join(os.path.dirname(__file__), "pacientes.db")

def _db_conn():
    conn = sqlite3.connect(PACIENTES_DB)
    conn.row_factory = sqlite3.Row
    return conn

def _init_pacientes_db():
    conn = _db_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pacientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT NOT NULL,
            data_nascimento TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()

_init_pacientes_db()

def salvar_paciente(nome, telefone, data_nascimento, paciente_id=None):
    nome = (nome or "").strip().upper()
    telefone = (telefone or "").strip()
    data_nascimento = (data_nascimento or "").strip()

    if not nome or not telefone:
        return None

    conn = _db_conn()
    cur = conn.cursor()

    if paciente_id:
        cur.execute(
            "UPDATE pacientes SET nome=?, telefone=?, data_nascimento=? WHERE id=?",
            (nome, telefone, data_nascimento, int(paciente_id))
        )
        conn.commit()
        conn.close()
        return int(paciente_id)

    cur.execute(
        "INSERT INTO pacientes (nome, telefone, data_nascimento) VALUES (?, ?, ?)",
        (nome, telefone, data_nascimento)
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return int(new_id)

def buscar_paciente(paciente_id):
    try:
        pid = int(paciente_id)
    except:
        return None
    conn = _db_conn()
    row = conn.execute("SELECT * FROM pacientes WHERE id=?", (pid,)).fetchone()
    conn.close()
    return row

def listar_pacientes():
    conn = _db_conn()
    rows = conn.execute("SELECT * FROM pacientes ORDER BY id DESC").fetchall()
    conn.close()
    return rows

PACIENTES_TEMPLATE = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Pacientes - OftalmoPrev</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { background:#f4f7f9; font-family:'Segoe UI',sans-serif; }
    .card { border:none; border-radius:12px; box-shadow:0 4px 10px rgba(0,0,0,0.08); }
    .topbar { background:#002b5c; color:#fff; padding:14px 0; border-bottom:4px solid #007bff; }
    .badge-pill { border-radius:999px; }
  </style>
</head>
<body>
<div class="topbar text-center">
  <h5 class="mb-0">Pacientes cadastrados</h5>
  <div class="small text-white-50">Ações: editar cadastro e exportar Excel (Nome + Telefone)</div>
</div>

<div class="container py-4">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <div class="small text-secondary">
      Total: <b>{{total}}</b>
    </div>
    <div class="d-flex gap-2">
      <a class="btn btn-outline-primary" href="/pacientes.xlsx">Baixar Excel</a>
      <a class="btn btn-secondary" href="/">Voltar</a>
    </div>
  </div>

  <div class="card p-3">
    <div class="table-responsive">
      <table class="table table-striped align-middle mb-0">
        <thead>
          <tr>
            <th style="width:80px;">ID</th>
            <th>Nome</th>
            <th style="width:220px;">Telefone</th>
            <th style="width:170px;">Nascimento</th>
            <th style="width:120px;"></th>
          </tr>
        </thead>
        <tbody>
        {% for p in pacientes %}
          <tr>
            <td><span class="badge bg-dark badge-pill">{{p['id']}}</span></td>
            <td class="fw-bold">{{p['nome']}}</td>
            <td>{{p['telefone']}}</td>
            <td>{{p['data_nascimento'] or ""}}</td>
            <td class="text-end">
              <a class="btn btn-sm btn-primary" href="/?paciente_id={{p['id']}}">
                Editar
              </a>
            </td>
          </tr>
        {% endfor %}
        {% if total == 0 %}
          <tr>
            <td colspan="5" class="text-center text-secondary py-4">
              Nenhum paciente cadastrado ainda.
            </td>
          </tr>
        {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</div>
</body>
</html>
"""

@app.route("/pacientes", methods=["GET"])
def pacientes():
    rows = listar_pacientes()
    return render_template_string(PACIENTES_TEMPLATE, pacientes=rows, total=len(rows))

@app.route("/pacientes.xlsx", methods=["GET"])
def pacientes_xlsx():
    rows = listar_pacientes()

    wb = Workbook()
    ws = wb.active
    ws.title = "Pacientes"
    ws.append(["Nome", "Telefone"])
    for p in rows:
        ws.append([p["nome"], p["telefone"]])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"pacientes_{date.today().strftime('%Y-%m-%d')}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ==================================================
# 1. BANCO DE DADOS (USANDO IDs SEM ACENTO PARA SEGURANÇA)
# ==================================================

DB_REGRAS = {
    # --- REFRATIVO E CÓRNEA ---
    "miopia": {
        "label": "Miopia > -1.00",
        "score": 3,
        "exames": ["Mapeamento de Retina – AO", "Retinografia Simples – AO"],
        "just": "Rastreio de fragilidade periférica e degenerações latentes.",
        "ref": "AAO PPP Comprehensive Evaluation",
        "link": "https://www.aao.org/education/preferred-practice-pattern/comprehensive-adult-medical-eye-evaluation-ppp"
    },
    "astigmatismo": {
        "label": "Astigmatismo > -1.50",
        "score": 2,
        "exames": ["Ceratoscopia / Topografia – AO", "Paquimetria de Córnea – AO"],
        "just": "Avaliação de curvatura e espessura para descarte de ectasias.",
        "ref": "Global Consensus on Keratoconus",
        "link": "https://pubmed.ncbi.nlm.nih.gov/25901970/"
    },
    "ceratocone": {
        "label": "Suspeita de ceratocone",
        "score": 4,
        "exames": ["Ceratoscopia / Topografia – AO", "Paquimetria de Córnea – AO"],
        "just": "Investigação estrutural corneana baseada em padrões de curvatura.",
        "ref": "AAO Corneal PPP",
        "link": "https://www.aao.org/education/preferred-practice-pattern/corneal-ectasia-ppp"
    },
    "hipermetropia": {
        "label": "Hipermetropia > +2.50",
        "score": 2,
        "exames": ["Gonioscopia – AO"],
        "just": "Rastreio preventivo de ângulo estreito em hipermétropes.",
        "ref": "AAO POAG PPP",
        "link": "https://www.aao.org/education/preferred-practice-pattern/primary-open-angle-glaucoma-ppp"
    },
    "olho_seco": {
        "label": "Olho seco",
        "score": 2,
        "exames": ["Teste de Shirmmer – AO", "Ceratoscopia – AO"],
        "just": "Avaliação de superfície ocular e filme lacrimal.",
        "ref": "TFOS DEWS II",
        "link": "https://www.tfosdewsneureport.org/"
    },

    # --- GLAUCOMA E PIO ---
    "camara_rasa": {
        "label": "Câmara anterior rasa",
        "score": 4,
        "exames": ["Gonioscopia – AO"],
        "just": "Risco elevado de fechamento angular primário.",
        "ref": "CBO Consenso Glaucoma",
        "link": "https://www.cbo.com.br/novo/medico/pdf/Diretrizes_CBO_AMB_CFM.pdf"
    },
    "pio_alta": {
        "label": "PIO > 19mmHg",
        "score": 5,
        "exames": ["Paquimetria – AO", "Campimetria – AO", "OCT de Nervo Óptico – AO"],
        "just": "Investigação de HT Ocular e risco OHTS.",
        "ref": "OHTS Study / AAO POAG",
        "link": "https://www.aao.org/education/preferred-practice-pattern/primary-open-angle-glaucoma-ppp"
    },
    "escavacao": {
        "label": "Escavação > 0.5",
        "score": 4,
        "exames": ["OCT de Nervo Óptico – AO", "Campimetria – AO"],
        "just": "Avaliação estrutural e funcional do nervo óptico.",
        "ref": "AAO POAG PPP",
        "link": "https://www.aao.org/education/preferred-practice-pattern/primary-open-angle-glaucoma-ppp"
    },
    "susp_glaucoma": {
        "label": "Suspeita de glaucoma",
        "score": 4,
        "exames": ["OCT de Nervo Óptico – AO", "Paquimetria – AO", "Campimetria – AO"],
        "just": "Rastreio multimodal para detecção precoce.",
        "ref": "CBO / ICO Guidelines",
        "link": "https://www.cbo.com.br/novo/medico/pdf/Diretrizes_CBO_AMB_CFM.pdf"
    },
    "glaucoma_ok": {
        "label": "Glaucoma confirmado",
        "score": 5,
        "exames": ["Curva Tensional Diária ou TSH – AO", "Campimetria – AO", "Gonioscopia – AO"],
        "just": "Monitoramento de progressão e estabilidade tensional.",
        "ref": "SOE Guidelines",
        "link": "https://www.soe.org/guidelines/"
    },

    # --- RETINA E SISTÊMICO ---
    "lesao_retina": {
        "label": "Lesão em retina",
        "score": 5,
        "exames": ["Mapeamento de Retina – AO", "OCT de Mácula – AO"],
        "just": "Monitoramento de integridade retiniana e camadas neurais.",
        "ref": "CBO Diretrizes Retina",
        "link": "https://www.cbo.com.br/novo/medico/pdf/Diretrizes_CBO_AMB_CFM.pdf"
    },
    "diabetes": {
        "label": "Diabetes",
        "score": 5,
        "exames": ["OCT de Mácula – AO", "Retinografia Colorida – AO"],
        "just": "Rastreio de retinopatia diabética (ETDRS).",
        "ref": "AAO Diabetic Retinopathy",
        "link": "https://www.aao.org/education/preferred-practice-pattern/diabetic-retinopathy-ppp"
    },
    "hipertensao": {
        "label": "Hipertensão",
        "score": 3,
        "exames": ["Mapeamento de Retina – AO", "Retinografia Colorida – AO"],
        "just": "Avaliação de alterações microvasculares sistêmicas.",
        "ref": "Diretrizes SBC/CBO",
        "link": "https://www.cbo.com.br/novo/medico/pdf/Diretrizes_CBO_AMB_CFM.pdf"
    },
    "trauma": {
        "label": "Trauma ocular",
        "score": 4,
        "exames": ["USG Ocular – AO", "Gonioscopia – AO", "Mapeamento de Retina"],
        "just": "Avaliação de danos estruturais e risco de recessão angular.",
        "ref": "Ocular Trauma Score",
        "link": "https://pubmed.ncbi.nlm.nih.gov/12028607/"
    },
    "uveite": {
        "label": "Suspeita de uveíte",
        "score": 4,
        "exames": ["OCT de Mácula – AO", "USG Ocular – AO"],
        "just": "Pesquisa de focos inflamatórios e complicações maculares.",
        "ref": "IUSG Guidelines",
        "link": "https://www.iusg.net/"
    },

    # --- PROTOCOLOS ESPECIAIS E PEDIÁTRICOS ---
    "pos_cirurgia": {
        "label": "Check-up Cirurgia ocular > 1 ano",
        "score": 2,
        "exames": ["Microscopia Especular – AO", "Mapeamento de Retina"],
        "just": "Monitoramento endotelial e integridade pós-cirúrgica.",
        "ref": "AAO Corneal Endothelial",
        "link": "https://www.aao.org/education/preferred-practice-pattern/corneal-endothelial-ppp"
    },
    "ped_miopia": {
        "label": "Criança com Miopia (2-15 anos)",
        "score": 3,
        "exames": ["Retinografia Simples – AO", "Biometria Óptica – AO", "Mapeamento de Retina –na– AO".replace("na–", "–"), "Mapeamento de Retina – AO"],
        "just": "Protocolo de controle de miopia: monitoramento axial e fundoscopia.",
        "ref": "CBO/SOE Miopia Infantil",
        "link": "https://www.cbo.com.br/"
    },
    "ped_astig": {
        "label": "Criança c/ Miopia + Astigmatismo",
        "score": 3,
        "exames": ["Topografia – AO", "Biometria Óptica – AO", "Mapeamento de Retina – AO"],
        "just": "Rastreio combinado: ectasias e crescimento axial.",
        "ref": "CBO/SOE Miopia Infantil",
        "link": "https://www.cbo.com.br/"
    }
}

# Organização Visual
LAYOUT_INTERFACE = {
    "Distúrbios Refracionais e Córnea": ["miopia", "astigmatismo", "ceratocone", "hipermetropia", "olho_seco"],
    "Glaucoma e Tonometria": ["pio_alta", "escavacao", "susp_glaucoma", "glaucoma_ok", "camara_rasa"],
    "Patologias Sistémica e Inflamatórias": ["diabetes", "hipertensao", "lesao_retina", "uveite", "trauma"],
    "Protocolos Especiais e Pediátricos": ["pos_cirurgia", "ped_miopia", "ped_astig"]
}

TODOS_EXAMES_POSSIVEIS = [
    "Mapeamento de Retina – AO", "Retinografia Simples – AO", "Retinografia Colorida – AO",
    "Ceratoscopia / Topografia – AO", "Paquimetria de Córnea – AO", "Paquimetria – AO",
    "Gonioscopia – AO", "Campimetria – AO", "OCT de Nervo Óptico – AO", "OCT de Mácula – AO",
    "Curva Tensional Diária ou TSH – AO", "Microscopia Especular – AO", "USG Ocular – AO",
    "Teste de Shirmmer – AO", "Biometria Óptica – AO"
]

QUEIXAS_LIST = [
    "Dificuldade visual para longe",
    "Visão embaçada / flutuação",
    "Fotofobia / Lacremejamento",
    "Pós-operatório oftalmológico",
    "Dificuldade visual para perto",
    "Dor ocular / pressão ocular",
    "Trauma ocular recente"
]
HISTORICO_LIST = [
    "Uso de óculos ou lentes",
    "Glaucoma / hipertensão ocular",
    "Retinopatia / DMRI / uveíte",
    "Uso de corticoides / imunossupressores",
    "Acompanhamento periódico",
    "Doenças sistêmicas (DM, HAS)"
]

# ==================================================
# 2. LÓGICA DE PROCESSAMENTO
# ==================================================

def calcular_idade(data_nascimento_str):
    if not data_nascimento_str:
        return ""
    try:
        dn = datetime.strptime(data_nascimento_str, "%Y-%m-%d").date()
        hoje = date.today()
        return hoje.year - dn.year - ((hoje.month, hoje.day) < (dn.month, dn.day))
    except:
        return ""

def processar_dados_clinicos(ids_selecionados, idade):
    exames_set = set()
    links_list = []
    justificativas_list = []
    achados_nomes = []
    score_total = 0

    IDS_PEDIATRICOS = ["ped_miopia", "ped_astig"]

    for id_item in ids_selecionados:
        if id_item in IDS_PEDIATRICOS and isinstance(idade, int) and idade > 15:
            print(f"DEBUG: Bloqueado item {id_item} para paciente de {idade} anos.")
            continue

        dados = DB_REGRAS.get(id_item)
        if dados:
            score_total += dados["score"]
            exames_set.update(dados["exames"])
            achados_nomes.append(dados["label"])

            if not any(l["nome"] == dados["ref"] for l in links_list):
                links_list.append({"nome": dados["ref"], "link": dados["link"]})

            justificativas_list.append(f"[{dados['label']}]: {dados['just']}")

    if isinstance(idade, int) and idade > 60:
        score_total += 2
        justificativas_list.append("[Idade > 60]: Fator de risco senil adicionado (+2).")

    return sorted(list(exames_set)), links_list, justificativas_list, achados_nomes, min(score_total, 10)

# ==================================================
# 3. INTERFACE E IMPRESSÃO
# ==================================================

TEMPLATE_INTERFACE = """
<!doctype html>
<html lang="pt-br">
<head>
    <meta charset="utf-8">
    <title>OftalmoPrev - CDSS (BIR v3.13)</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f9; font-family: 'Segoe UI', sans-serif; }
        .header-top { background: #002b5c; color: white; padding: 20px 0; border-bottom: 5px solid #007bff; }
        .card { border: none; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.08); }
        .section-header { border-left: 5px solid #007bff; padding-left: 12px; margin: 15px 0; font-weight: bold; color: #002b5c; }
        .score-badge { font-size: 1.5rem; font-weight: bold; padding: 10px 20px; border-radius: 50px; color: white; }
        .bg-low { background-color: #28a745; }
        .bg-med { background-color: #ffc107; color: #333; }
        .bg-high { background-color: #dc3545; }
        .ref-link { text-decoration: none; margin-right: 5px; margin-bottom: 5px; display: inline-block; }
        .topbar-right { display:flex; gap:10px; justify-content:center; align-items:center; }
        .btn-plus {
            width: 38px; height: 38px;
            display:flex; align-items:center; justify-content:center;
            border-radius: 10px;
        }
    </style>
</head>
<body>
<div class="header-top text-center">
    <div class="topbar-right">
        <div>
            <h3 class="mb-0">OftalmoPrev – Sistema de Suporte a Decisão</h3>
            <p class="small mb-0">Build de Integração Restrita (BIR) - Python 3.12 Compatible</p>
            <div class="small">Logado como: <b>{{user}}</b> | <a href="/logout" class="text-white-50">Sair</a></div>
        </div>
    </div>
</div>

<div class="container py-4">
    <form method="post">
        <div class="card p-4 mb-4">
            <div class="row g-3 align-items-end">
                <input type="hidden" name="paciente_id" value="{{paciente_id}}">
                <div class="col-md-5">
                    <label class="small fw-bold">Paciente</label>
                    <input class="form-control" name="nome" value="{{nome}}">
                </div>
                <div class="col-md-3">
                    <label class="small fw-bold">Telefone</label>
                    <input class="form-control" name="telefone" value="{{telefone}}" placeholder="(DD) 9xxxx-xxxx">
                </div>
                <div class="col-md-3">
                    <label class="small fw-bold">Data Nasc.</label>
                    <input type="date" class="form-control" name="data_nascimento" value="{{data_nascimento}}">
                </div>
                <div class="col-md-1">
                    <label class="small fw-bold">Idade</label>
                    <div class="d-flex gap-2">
                        <input class="form-control bg-light" name="idade" value="{{idade}}" readonly>
                        <a class="btn btn-outline-light border btn-plus" href="/pacientes" title="Ver pacientes" target="_blank">+</a>
                    </div>
                </div>
            </div>
            <div class="small text-secondary mt-2">
                * Ao analisar, o paciente (Nome + Telefone) é salvo/atualizado automaticamente.
            </div>
        </div>

        <div class="row">
            <div class="col-md-4">
                <div class="section-header">ANAMNESE</div>
                <div class="card p-3 mb-3">
                    <h6 class="text-secondary small fw-bold">QUEIXAS</h6>
                    {% for q in QUEIXAS %}
                        <div class="form-check">
                            <input class="form-check-input" type="checkbox" name="queixas" value="{{q}}" {% if q in queixas_sel %}checked{% endif %}>
                            <label class="small">{{q}}</label>
                        </div>
                    {% endfor %}
                    <hr>
                    <h6 class="text-secondary small fw-bold">HISTÓRICO</h6>
                    {% for h in HISTORICO %}
                        <div class="form-check">
                            <input class="form-check-input" type="checkbox" name="historico" value="{{h}}" {% if h in historico_sel %}checked{% endif %}>
                            <label class="small">{{h}}</label>
                        </div>
                    {% endfor %}
                </div>
            </div>

            <div class="col-md-8">
                <div class="section-header">ACHADOS CLÍNICOS</div>
                <div class="row">
                    {% for cat, id_list in LAYOUT.items() %}
                    <div class="col-md-6 mb-3">
                        <div class="card p-3 h-100">
                            <h6 class="text-primary small fw-bold">{{cat}}</h6>
                            {% for id_item in id_list %}
                            <div class="form-check">
                                <input class="form-check-input" type="checkbox" name="achados" value="{{id_item}}" {% if id_item in achados_sel %}checked{% endif %}>
                                <label class="small">{{ DB_REGRAS[id_item]['label'] }}</label>
                            </div>
                            {% endfor %}
                        </div>
                    </div>
                    {% endfor %}
                </div>
            </div>
        </div>

        <button class="btn btn-primary w-100 py-3 mt-3 fw-bold shadow">ANALISAR RISCO & PROTOCOLOS</button>

        {% if exames or score > 0 %}
        <div class="card p-4 mt-4 border-start border-success border-5 shadow">
            <div class="d-flex justify-content-between align-items-center mb-4">
                <h5 class="text-success fw-bold m-0">PROPOSTA TÉCNICA</h5>
                <div class="text-center">
                    <span class="score-badge {% if score < 4 %}bg-low{% elif score < 7 %}bg-med{% else %}bg-high{% endif %}">
                        Risco: {{score}}/10
                    </span>
                </div>
            </div>

            {% if links %}
            <div class="mb-3">
                <label class="small fw-bold text-secondary">DIRETRIZES APLICÁVEIS:</label><br>
                {% for l in links %}
                <a href="{{l.link}}" target="_blank" class="badge bg-primary ref-link">📚 {{l.nome}}</a>
                {% endfor %}
            </div>
            {% endif %}

            <div class="row mb-3">
                {% for e in exames %}
                <div class="col-md-6"><div class="p-2 border bg-white mb-1 small fw-bold">✓ {{e}}</div></div>
                {% endfor %}
            </div>

            <div class="mb-3">
                <label class="small fw-bold text-primary">OBSERVAÇÕES MÉDICAS:</label>
                <textarea class="form-control" name="exames_extras" rows="2"></textarea>
            </div>

            {% for j in just_raw %}<input type="hidden" name="just_hidden" value="{{j}}">{% endfor %}
            {% for l in links %}<input type="hidden" name="ref_hidden" value="{{l.nome}}">{% endfor %}
            {% for e in exames %}<input type="hidden" name="exames_hidden" value="{{e}}">{% endfor %}
            {% for n in achados_nomes %}<input type="hidden" name="nomes_hidden" value="{{n}}">{% endfor %}
            <input type="hidden" name="score_hidden" value="{{score}}">

            <button formaction="/imprimir" formtarget="_blank" class="btn btn-dark w-100 py-3 mt-4 fw-bold">GERAR PRESCRIÇÃO A4</button>
        </div>
        {% endif %}
    </form>
</div>
</body>
</html>
"""

TEMPLATE_PRESCRICAO = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
    @page { size: A4; margin: 2.5cm 2cm 2cm 2cm; }
    body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.4; color: #333; }
    .header { text-align: center; border-bottom: 2px solid #002b5c; margin-bottom: 15px; }
    .patient-box { background: #f2f2f2; padding: 10px; border-radius: 5px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; }
    .checklist { display: grid; grid-template-columns: 1fr 1fr; gap: 5px; margin-bottom: 15px; }
    .box-full { display: inline-block; width: 12px; height: 12px; background-color: #002b5c; border: 1px solid #000; margin-right: 5px; color: white; font-size: 10px; text-align: center; line-height: 12px;}
    .box-empty { display: inline-block; width: 12px; height: 12px; border: 1px solid #000; margin-right: 5px; }
    .just-box { border-left: 4px solid #002b5c; background: #f9f9f9; padding: 10px; font-size: 9pt; font-style: italic; margin-top: 10px; }
    .footer-lgpd { position: fixed; bottom: 0; width: 100%; font-size: 8pt; text-align: center; color: #777; border-top: 1px solid #eee; padding-top: 5px; }
    .score-print { border: 2px solid #000; padding: 5px 10px; font-weight: bold; border-radius: 5px; }
</style>
</head>
<body>
    <div class="header">
        <h2 style="margin:0; color: #002b5c;">GUIA DE RASTREAMENTO OFTALMOLÓGICO</h2>
        <small>Protocolo Baseado em Evidências (AAO / CBO / SOE)</small>
    </div>

    <div class="patient-box">
        <div>
            <strong>Paciente:</strong> {{nome}} {% if idade %}({{idade}} anos){% endif %}<br>
            <strong>Telefone:</strong> {{telefone}}<br>
            <strong>Data:</strong> {{data}}
        </div>
        <div class="score-print">Risco Clínico: {{score}}/10</div>
    </div>

    <div style="margin-bottom: 15px;">
        <strong>Indicação Clínica:</strong>
        {% if achados_nomes %}{{ achados_nomes | join(', ') }}{% else %}Rotina{% endif %}
    </div>

    <h4>1. Propedêutica Indicada (Checklist):</h4>
    <div class="checklist">
        {% for item in checklist %}
        <div>{% if item.marcado %}<span class="box-full">X</span>{% else %}<span class="box-empty"></span>{% endif %} {{item.nome}}</div>
        {% endfor %}
    </div>

    {% if exames_extras %}
    <div style="margin-top: 10px; border: 1px solid #ccc; padding: 10px;">
        <strong>Observações Complementares:</strong><br>{{exames_extras}}
    </div>
    {% endif %}

    <div class="just-box">
        <strong>Justificativa Técnica da Solicitação:</strong><br>
        {% for j in justificativas %}
        {{j}}<br>
        {% endfor %}
    </div>

    <div style="margin-top: 15px; font-size: 9pt;">
        <strong>Referências Bibliográficas Consultadas:</strong><br>
        {{ referencias | join(' | ') }}
    </div>

    <div class="footer-lgpd">
        Este documento é gerado por Sistema de Apoio à Decisão Clínica (CDSS). O processamento de dados observa a LGPD (Lei 13.709/2018).<br>
        Assinatura do Responsável
    </div>
    <script>window.onload = function() { window.print(); }</script>
</body>
</html>
"""

# ==================================================
# 4. ROTAS FLASK
# ==================================================
@app.route("/", methods=["GET", "POST"])
def index():
    # Se veio para editar um paciente (clicou "Editar" na lista)
    paciente_id_qs = request.args.get("paciente_id", "").strip()
    paciente = buscar_paciente(paciente_id_qs) if paciente_id_qs else None

    dados = {
        "paciente_id": paciente["id"] if paciente else "",
        "nome": paciente["nome"] if paciente else "",
        "telefone": paciente["telefone"] if paciente else "",
        "data_nascimento": paciente["data_nascimento"] if paciente else "",
        "idade": calcular_idade(paciente["data_nascimento"]) if paciente else "",
        "queixas_sel": [],
        "historico_sel": [],
        "achados_sel": [],
        "exames": [],
        "just_raw": [],
        "links": [],
        "achados_nomes": [],
        "score": 0
    }

    if request.method == "POST":
        print("-" * 30)
        print("DEBUG: Recebendo POST do formulário...")

        paciente_id_form = (request.form.get("paciente_id") or "").strip()
        dados["paciente_id"] = paciente_id_form

        dados["nome"] = (request.form.get("nome", "") or "").upper()
        dados["telefone"] = (request.form.get("telefone", "") or "").strip()
        dados["data_nascimento"] = request.form.get("data_nascimento", "")
        dados["queixas_sel"] = request.form.getlist("queixas")
        dados["historico_sel"] = request.form.getlist("historico")
        dados["achados_sel"] = request.form.getlist("achados")

        print(f"DEBUG: Achados (IDs) recebidos: {dados['achados_sel']}")

        idade = calcular_idade(dados["data_nascimento"])
        dados["idade"] = idade

        # Salva/atualiza paciente SOMENTE se tiver nome + telefone
        saved_id = salvar_paciente(
            dados["nome"],
            dados["telefone"],
            dados["data_nascimento"],
            paciente_id_form if paciente_id_form else None
        )
        if saved_id:
            dados["paciente_id"] = str(saved_id)

        exames, links, justs, achados_nomes, score = processar_dados_clinicos(dados["achados_sel"], idade)

        dados.update({
            "exames": exames,
            "links": links,
            "just_raw": justs,
            "achados_nomes": achados_nomes,
            "score": score
        })

    return render_template_string(
        TEMPLATE_INTERFACE,
        QUEIXAS=QUEIXAS_LIST,
        HISTORICO=HISTORICO_LIST,
        LAYOUT=LAYOUT_INTERFACE,
        DB_REGRAS=DB_REGRAS,
        user=session.get("user", ""),
        **dados
    )

@app.route("/imprimir", methods=["POST"])
def imprimir():
    exames_sugeridos = request.form.getlist("exames_hidden")
    exames_extras = request.form.get("exames_extras", "").strip()
    justificativas = request.form.getlist("just_hidden")
    referencias = request.form.getlist("ref_hidden")
    achados_nomes = request.form.getlist("nomes_hidden")
    score = request.form.get("score_hidden", "0")

    nome = request.form.get("nome", "")
    telefone = request.form.get("telefone", "")
    idade = request.form.get("idade", "")

    checklist_final = []
    sugeridos_set = set(exames_sugeridos)
    for ex in TODOS_EXAMES_POSSIVEIS:
        checklist_final.append({"nome": ex, "marcado": ex in sugeridos_set})

    return render_template_string(
        TEMPLATE_PRESCRICAO,
        nome=nome,
        telefone=telefone,
        idade=idade,
        score=score,
        checklist=checklist_final,
        exames_extras=exames_extras,
        justificativas=justificativas,
        achados_nomes=achados_nomes,
        referencias=referencias,
        data=date.today().strftime("%d/%m/%Y")
    )

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

if __name__ == "__main__":
    Timer(1.5, open_browser).start()
    app.run(port=5000)
