import os
from functools import wraps
from pathlib import Path
from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect,
    url_for, flash, send_file, jsonify, session
)
from werkzeug.utils import secure_filename

import database as db
import ocr
import contract_generator as cg

load_dotenv()

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "pdf"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "viver-de-obra-2026-secret")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.url))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.url))
        if session.get("user_role") != "admin":
            flash("Acesso restrito a administradores.", "danger")
            return redirect(url_for("index"))
        return f(*args, **kwargs)
    return decorated


def is_admin():
    return session.get("user_role") == "admin"


# Injeta variáveis globais nos templates
@app.context_processor
def inject_globals():
    return {
        "current_user": session.get("user_login"),
        "current_role": session.get("user_role"),
        "is_admin": is_admin(),
    }


# ---------------------------------------------------------------------------
# Login / Logout
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("index"))
    if request.method == "POST":
        usuario = db.autenticar_usuario(
            request.form["login"].strip(),
            request.form["senha"],
        )
        if usuario:
            session["user_id"] = usuario["id"]
            session["user_login"] = usuario["login"]
            session["user_role"] = usuario["role"]
            next_url = request.args.get("next") or url_for("index")
            return redirect(next_url)
        flash("Login ou senha incorretos.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Usuários (admin)
# ---------------------------------------------------------------------------

@app.route("/usuarios")
@admin_required
def usuarios():
    lista = db.listar_usuarios()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@admin_required
def novo_usuario():
    if request.method == "POST":
        login_val = request.form["login"].strip()
        senha = request.form["senha"]
        role = request.form.get("role", "visualizador")
        if not login_val or not senha:
            flash("Login e senha são obrigatórios.", "danger")
        else:
            try:
                db.criar_usuario(login_val, senha, role)
                flash(f"Usuário '{login_val}' criado!", "success")
                return redirect(url_for("usuarios"))
            except Exception:
                flash("Esse login já existe.", "danger")
    return render_template("usuario_form.html", usuario=None, action="novo")


@app.route("/usuarios/<int:uid>/editar", methods=["GET", "POST"])
@admin_required
def editar_usuario(uid):
    lista = db.listar_usuarios()
    usuario = next((u for u in lista if u["id"] == uid), None)
    if not usuario:
        flash("Usuário não encontrado.", "danger")
        return redirect(url_for("usuarios"))
    if request.method == "POST":
        login_val = request.form["login"].strip()
        role = request.form.get("role", "visualizador")
        nova_senha = request.form.get("senha", "").strip() or None
        db.atualizar_usuario(uid, login_val, role, nova_senha)
        flash("Usuário atualizado!", "success")
        return redirect(url_for("usuarios"))
    return render_template("usuario_form.html", usuario=usuario, action="editar")


@app.route("/usuarios/<int:uid>/excluir", methods=["POST"])
@admin_required
def excluir_usuario(uid):
    if uid == session.get("user_id"):
        flash("Você não pode excluir seu próprio usuário.", "danger")
    else:
        db.excluir_usuario(uid)
        flash("Usuário removido.", "info")
    return redirect(url_for("usuarios"))


# ---------------------------------------------------------------------------
# Catálogos: Categorias, Unidades, Formas de Pagamento
# ---------------------------------------------------------------------------

@app.route("/catalogos")
@admin_required
def catalogos():
    return render_template(
        "catalogos.html",
        categorias=db.listar_categorias(apenas_ativas=False),
        unidades=db.listar_unidades(apenas_ativas=False),
        formas=db.listar_formas_pagamento(apenas_ativas=False),
    )


@app.route("/catalogos/categoria/nova", methods=["POST"])
@admin_required
def nova_categoria():
    nome = request.form.get("nome", "").strip()
    if nome:
        db.criar_categoria(nome)
        flash(f"Categoria '{nome}' adicionada.", "success")
    return redirect(url_for("catalogos"))


@app.route("/catalogos/categoria/<int:cid>/excluir", methods=["POST"])
@admin_required
def excluir_categoria(cid):
    db.excluir_categoria(cid)
    flash("Categoria removida.", "info")
    return redirect(url_for("catalogos"))


@app.route("/catalogos/unidade/nova", methods=["POST"])
@admin_required
def nova_unidade():
    nome = request.form.get("nome", "").strip()
    if nome:
        db.criar_unidade(nome)
        flash(f"Unidade '{nome}' adicionada.", "success")
    return redirect(url_for("catalogos"))


@app.route("/catalogos/unidade/<int:uid>/excluir", methods=["POST"])
@admin_required
def excluir_unidade(uid):
    db.excluir_unidade(uid)
    flash("Unidade removida.", "info")
    return redirect(url_for("catalogos"))


@app.route("/catalogos/forma-pagamento/nova", methods=["POST"])
@admin_required
def nova_forma_pagamento():
    descricao = request.form.get("descricao", "").strip()
    if descricao:
        db.criar_forma_pagamento(descricao)
        flash(f"Forma de pagamento adicionada.", "success")
    return redirect(url_for("catalogos"))


@app.route("/catalogos/forma-pagamento/<int:fid>/excluir", methods=["POST"])
@admin_required
def excluir_forma_pagamento(fid):
    db.excluir_forma_pagamento(fid)
    flash("Forma de pagamento removida.", "info")
    return redirect(url_for("catalogos"))


# ---------------------------------------------------------------------------
# Dashboard principal
# ---------------------------------------------------------------------------

@app.route("/")
@login_required
def index():
    clientes = db.listar_clientes()
    executoras = db.listar_executoras()
    servicos = db.listar_servicos()
    resumo = db.resumo_geral()
    cliente_id = request.args.get("cliente_id", type=int)
    cliente_sel = None
    metricas = None
    if cliente_id:
        cliente_sel = db.buscar_cliente(cliente_id)
        metricas = db.dashboard_cliente(cliente_id)
        contratos = db.listar_contratos(cliente_id=cliente_id)
    else:
        contratos = db.listar_contratos()
    return render_template("index.html",
                           contratos=contratos,
                           clientes=clientes,
                           executoras=executoras,
                           servicos=servicos,
                           resumo=resumo,
                           cliente_sel=cliente_sel,
                           metricas=metricas,
                           cliente_id=cliente_id)


# ---------------------------------------------------------------------------
# Dashboard por Cliente
# ---------------------------------------------------------------------------

@app.route("/clientes/<int:cid>/dashboard")
@login_required
def dashboard_cliente(cid):
    cliente = db.buscar_cliente(cid)
    if not cliente:
        flash("Cliente não encontrado.", "danger")
        return redirect(url_for("clientes"))
    metricas = db.dashboard_cliente(cid)
    contratos = db.listar_contratos(cliente_id=cid)
    return render_template(
        "dashboard_cliente.html",
        cliente=cliente,
        metricas=metricas,
        contratos=contratos,
    )


# ---------------------------------------------------------------------------
# Clientes
# ---------------------------------------------------------------------------

@app.route("/clientes")
@login_required
def clientes():
    lista = db.listar_clientes(apenas_ativos=False)
    return render_template("clientes.html", clientes=lista)


@app.route("/clientes/novo", methods=["GET", "POST"])
@login_required
def novo_cliente():
    ocr_dados = {}
    if request.method == "POST":
        action = request.form.get("action")
        if action == "ocr":
            arquivos = request.files.getlist("documentos")
            ocr_dados = processar_uploads(arquivos)
            if not ocr_dados:
                flash("Nenhum documento processado. Preencha manualmente.", "warning")
            return render_template("cliente_form.html", cliente=None,
                                   action="novo", ocr_dados=ocr_dados)
        dados = _dados_cliente_do_form(request.form)
        if not dados["nome"]:
            flash("Nome é obrigatório.", "danger")
            return render_template("cliente_form.html", cliente=None,
                                   action="novo", ocr_dados={})
        db.criar_cliente(dados)
        flash("Cliente cadastrado com sucesso!", "success")
        return redirect(url_for("clientes"))
    return render_template("cliente_form.html", cliente=None, action="novo", ocr_dados={})


@app.route("/clientes/<int:cid>/editar", methods=["GET", "POST"])
@login_required
def editar_cliente(cid):
    cliente = db.buscar_cliente(cid)
    if not cliente:
        flash("Cliente não encontrado.", "danger")
        return redirect(url_for("clientes"))
    if request.method == "POST":
        dados = _dados_cliente_do_form(request.form)
        db.atualizar_cliente(cid, dados)
        flash("Cliente atualizado!", "success")
        return redirect(url_for("clientes"))
    return render_template("cliente_form.html", cliente=cliente, action="editar", ocr_dados={})


@app.route("/clientes/<int:cid>/excluir", methods=["POST"])
@admin_required
def excluir_cliente(cid):
    db.excluir_cliente(cid)
    flash("Cliente removido.", "info")
    return redirect(url_for("clientes"))


@app.route("/clientes/<int:cid>/historico")
@login_required
def historico_cliente(cid):
    cliente = db.buscar_cliente(cid)
    if not cliente:
        flash("Cliente não encontrado.", "danger")
        return redirect(url_for("clientes"))
    contratos = db.listar_contratos(cliente_id=cid)
    return render_template("historico_cliente.html", cliente=cliente, contratos=contratos)


def _dados_cliente_do_form(form):
    return {
        "nome": form.get("nome", "").strip(),
        "nacionalidade": form.get("nacionalidade", "brasileiro(a)").strip(),
        "estado_civil": form.get("estado_civil", "solteiro(a)").strip(),
        "profissao": form.get("profissao", "").strip(),
        "cpf": form.get("cpf", "").strip(),
        "rg": form.get("rg", "").strip(),
        "uf_rg": form.get("uf_rg", "").strip(),
        "logradouro": form.get("logradouro", "").strip(),
        "complemento": form.get("complemento", "").strip(),
    }


# ---------------------------------------------------------------------------
# Executoras
# ---------------------------------------------------------------------------

@app.route("/executoras")
@login_required
def executoras():
    lista = db.listar_executoras(apenas_ativas=False)
    return render_template("executoras.html", executoras=lista)


@app.route("/executoras/nova", methods=["GET", "POST"])
@login_required
def nova_executora():
    ocr_dados = {}
    if request.method == "POST":
        action = request.form.get("action")
        if action == "ocr":
            arquivos = request.files.getlist("documentos")
            ocr_dados = processar_uploads(arquivos)
            if not ocr_dados:
                flash("Nenhum documento processado. Preencha manualmente.", "warning")
            return render_template("executora_form.html", executora=None,
                                   action="nova", ocr_dados=ocr_dados)
        dados = _dados_executora_do_form(request.form)
        if not dados["nome"]:
            flash("Nome é obrigatório.", "danger")
            return render_template("executora_form.html", executora=None,
                                   action="nova", ocr_dados={})
        db.criar_executora(dados)
        flash("Executora cadastrada com sucesso!", "success")
        return redirect(url_for("executoras"))
    return render_template("executora_form.html", executora=None, action="nova", ocr_dados={})


@app.route("/executoras/<int:eid>/editar", methods=["GET", "POST"])
@login_required
def editar_executora(eid):
    executora = db.buscar_executora(eid)
    if not executora:
        flash("Executora não encontrada.", "danger")
        return redirect(url_for("executoras"))
    if request.method == "POST":
        dados = _dados_executora_do_form(request.form)
        db.atualizar_executora(eid, dados)
        flash("Executora atualizada!", "success")
        return redirect(url_for("executoras"))
    return render_template("executora_form.html", executora=executora,
                           action="editar", ocr_dados={})


@app.route("/executoras/<int:eid>/excluir", methods=["POST"])
@admin_required
def excluir_executora(eid):
    db.excluir_executora(eid)
    flash("Executora removida.", "info")
    return redirect(url_for("executoras"))


def _dados_executora_do_form(form):
    tipo = form.get("tipo_pessoa", "PF")
    return {
        "tipo_pessoa": tipo,
        "nome": form.get("nome", "").strip(),
        "cpf": form.get("cpf", "").strip() if tipo == "PF" else "",
        "cnpj": form.get("cnpj", "").strip() if tipo == "PJ" else "",
        "endereco": form.get("endereco", "").strip(),
        "telefone": form.get("telefone", "").strip(),
    }


def _snapshot_cliente_form(form) -> dict:
    return {k: form.get(k, "") for k in [
        "nome_contratante", "nacionalidade_contratante", "estado_civil_contratante",
        "profissao_contratante", "cpf_contratante", "rg_contratante",
        "uf_rg_contratante", "logradouro_contratante", "complemento_contratante",
    ]}


def _snapshot_executora_form(form) -> dict:
    return {k: form.get(k, "") for k in [
        "nome_executora", "tipo_pessoa_exec", "cpf_executora",
        "cnpj_executora", "endereco_executora",
    ]}


# ---------------------------------------------------------------------------
# Catálogo de Serviços
# ---------------------------------------------------------------------------

@app.route("/servicos")
@login_required
def servicos():
    lista = db.listar_servicos(apenas_ativos=False)
    return render_template("servicos.html", servicos=lista)


@app.route("/servicos/novo", methods=["GET", "POST"])
@login_required
def novo_servico():
    if request.method == "POST":
        nome = request.form["nome"].strip()
        descricao = request.form["descricao"].strip()
        unidade = request.form.get("unidade", "M²").strip()
        preco = request.form.get("preco_padrao", "").strip()
        percentual = request.form.get("percentual_padrao", "").strip()
        categoria = request.form.get("categoria", "").strip()
        if not nome or not descricao:
            flash("Nome e descrição são obrigatórios.", "danger")
            return render_template("servico_form.html", servico=None, action="novo",
                                   categorias=db.listar_categorias(),
                                   unidades=db.listar_unidades())
        preco_val = float(preco.replace(",", ".")) if preco else None
        perc_val = float(percentual.replace(",", ".")) if percentual else None
        db.criar_servico(nome, descricao, unidade, preco_val, categoria, perc_val)
        flash("Serviço cadastrado!", "success")
        return redirect(url_for("servicos"))
    return render_template("servico_form.html", servico=None, action="novo",
                           categorias=db.listar_categorias(),
                           unidades=db.listar_unidades())


@app.route("/servicos/<int:sid>/editar", methods=["GET", "POST"])
@login_required
def editar_servico(sid):
    servico = db.buscar_servico(sid)
    if not servico:
        flash("Serviço não encontrado.", "danger")
        return redirect(url_for("servicos"))
    if request.method == "POST":
        nome = request.form["nome"].strip()
        descricao = request.form["descricao"].strip()
        unidade = request.form.get("unidade", "M²").strip()
        preco = request.form.get("preco_padrao", "").strip()
        percentual = request.form.get("percentual_padrao", "").strip()
        categoria = request.form.get("categoria", "").strip()
        preco_val = float(preco.replace(",", ".")) if preco else None
        perc_val = float(percentual.replace(",", ".")) if percentual else None
        db.atualizar_servico(sid, nome, descricao, unidade, preco_val, categoria, perc_val)
        flash("Serviço atualizado!", "success")
        return redirect(url_for("servicos"))
    return render_template("servico_form.html", servico=servico, action="editar",
                           categorias=db.listar_categorias(),
                           unidades=db.listar_unidades())


@app.route("/servicos/<int:sid>/excluir", methods=["POST"])
@admin_required
def excluir_servico(sid):
    db.excluir_servico(sid)
    flash("Serviço removido.", "info")
    return redirect(url_for("servicos"))


# ---------------------------------------------------------------------------
# Novo Contrato — Passo 1: Tipo
# ---------------------------------------------------------------------------

@app.route("/contrato/novo")
@login_required
def escolher_tipo():
    tipos = cg.listar_tipos_contrato()
    return render_template("escolher_tipo.html", tipos=tipos)


# ---------------------------------------------------------------------------
# Novo Contrato — Passo 2: Selecionar/Cadastrar Partes
# ---------------------------------------------------------------------------

@app.route("/contrato/novo/<tipo>/partes", methods=["GET", "POST"])
@login_required
def selecionar_partes(tipo):
    if tipo not in cg.CONTRATOS:
        flash("Tipo inválido.", "danger")
        return redirect(url_for("escolher_tipo"))

    clientes_lista = db.listar_clientes()
    executoras_lista = db.listar_executoras()

    if request.method == "POST":
        action = request.form.get("action")

        if action == "ocr_cliente":
            arquivos = request.files.getlist("docs_cliente")
            dados = processar_uploads(arquivos)
            session["ocr_cliente"] = dados
            session["saved_exec_form"] = _snapshot_executora_form(request.form)
            session["saved_exec_id"] = request.form.get("executora_id", "")
            return redirect(url_for("selecionar_partes", tipo=tipo))

        if action == "ocr_executora":
            arquivos = request.files.getlist("docs_executora")
            dados = processar_uploads(arquivos)
            session["ocr_executora"] = dados
            session["saved_cli_form"] = _snapshot_cliente_form(request.form)
            session["saved_cli_id"] = request.form.get("cliente_id", "")
            return redirect(url_for("selecionar_partes", tipo=tipo))

        cliente_id = request.form.get("cliente_id") or None
        executora_id = request.form.get("executora_id") or None

        if not cliente_id and request.form.get("salvar_cliente"):
            dados_cli = {
                "nome": request.form.get("nome_contratante", "").strip(),
                "nacionalidade": request.form.get("nacionalidade_contratante", "brasileiro(a)").strip(),
                "estado_civil": request.form.get("estado_civil_contratante", "solteiro(a)").strip(),
                "profissao": request.form.get("profissao_contratante", "").strip(),
                "cpf": request.form.get("cpf_contratante", "").strip(),
                "rg": request.form.get("rg_contratante", "").strip(),
                "uf_rg": request.form.get("uf_rg_contratante", "").strip(),
                "logradouro": request.form.get("logradouro_contratante", "").strip(),
                "complemento": request.form.get("complemento_contratante", "").strip(),
            }
            if dados_cli["nome"]:
                cliente_id = str(db.criar_cliente(dados_cli))
                flash(f"Cliente '{dados_cli['nome']}' salvo!", "success")

        if not executora_id and request.form.get("salvar_executora"):
            tipo_pessoa = request.form.get("tipo_pessoa_exec", "PF")
            dados_exec = {
                "tipo_pessoa": tipo_pessoa,
                "nome": request.form.get("nome_executora", "").strip(),
                "cpf": request.form.get("cpf_executora", "").strip() if tipo_pessoa == "PF" else "",
                "cnpj": request.form.get("cnpj_executora", "").strip() if tipo_pessoa == "PJ" else "",
                "endereco": request.form.get("endereco_executora", "").strip(),
                "telefone": "",
            }
            if dados_exec["nome"]:
                executora_id = str(db.criar_executora(dados_exec))
                flash(f"Executora '{dados_exec['nome']}' salva!", "success")

        session["contrato_tipo"] = tipo
        session["contrato_cliente_id"] = int(cliente_id) if cliente_id else None
        session["contrato_executora_id"] = int(executora_id) if executora_id else None
        session["form_cliente"] = {
            "nome_contratante": request.form.get("nome_contratante", "").strip(),
            "nacionalidade_contratante": request.form.get("nacionalidade_contratante", "").strip(),
            "estado_civil_contratante": request.form.get("estado_civil_contratante", "").strip(),
            "profissao_contratante": request.form.get("profissao_contratante", "").strip(),
            "cpf_contratante": request.form.get("cpf_contratante", "").strip(),
            "rg_contratante": request.form.get("rg_contratante", "").strip(),
            "uf_rg_contratante": request.form.get("uf_rg_contratante", "").strip(),
            "logradouro_contratante": request.form.get("logradouro_contratante", "").strip(),
            "complemento_contratante": request.form.get("complemento_contratante", "").strip(),
        }
        session["form_executora"] = {
            "nome_executora": request.form.get("nome_executora", "").strip(),
            "tipo_pessoa_exec": request.form.get("tipo_pessoa_exec", "PF"),
            "cpf_executora": request.form.get("cpf_executora", "").strip(),
            "cnpj_executora": request.form.get("cnpj_executora", "").strip(),
            "endereco_executora": request.form.get("endereco_executora", "").strip(),
        }
        return redirect(url_for("formulario_contrato", tipo=tipo))

    ocr_cliente = session.pop("ocr_cliente", {})
    ocr_executora = session.pop("ocr_executora", {})
    saved_cli_form = session.pop("saved_cli_form", {})
    saved_exec_form = session.pop("saved_exec_form", {})
    saved_cli_id = session.pop("saved_cli_id", "")
    saved_exec_id = session.pop("saved_exec_id", "")

    return render_template(
        "selecionar_partes.html",
        tipo=tipo,
        info=cg.CONTRATOS[tipo],
        clientes=clientes_lista,
        executoras=executoras_lista,
        ocr_cliente=ocr_cliente,
        ocr_executora=ocr_executora,
        saved_cli_form=saved_cli_form,
        saved_exec_form=saved_exec_form,
        saved_cli_id=saved_cli_id,
        saved_exec_id=saved_exec_id,
    )


# ---------------------------------------------------------------------------
# Novo Contrato — Passo 3: Serviços, datas e geração
# ---------------------------------------------------------------------------

@app.route("/contrato/novo/<tipo>/formulario", methods=["GET", "POST"])
@login_required
def formulario_contrato(tipo):
    if tipo not in cg.CONTRATOS:
        flash("Tipo inválido.", "danger")
        return redirect(url_for("escolher_tipo"))

    servicos_catalogo = db.listar_servicos()
    formas_pagamento = db.listar_formas_pagamento()

    cliente_id = session.get("contrato_cliente_id")
    executora_id = session.get("contrato_executora_id")
    form_cliente = session.get("form_cliente", {})
    form_executora = session.get("form_executora", {})

    if cliente_id:
        c = db.buscar_cliente(cliente_id)
        if c:
            form_cliente = {
                "nome_contratante": c["nome"],
                "nacionalidade_contratante": c["nacionalidade"],
                "estado_civil_contratante": c["estado_civil"],
                "profissao_contratante": c["profissao"] or "",
                "cpf_contratante": c["cpf"] or "",
                "rg_contratante": c["rg"] or "",
                "uf_rg_contratante": c["uf_rg"] or "",
                "logradouro_contratante": c["logradouro"] or "",
                "complemento_contratante": c["complemento"] or "",
            }

    if executora_id:
        e = db.buscar_executora(executora_id)
        if e:
            form_executora = {
                "nome_executora": e["nome"],
                "tipo_pessoa_exec": e.get("tipo_pessoa", "PF"),
                "cpf_executora": e["cpf"] or "",
                "cnpj_executora": e.get("cnpj") or "",
                "endereco_executora": e["endereco"] or "",
            }

    dados_iniciais = {**form_cliente, **form_executora}

    if request.method == "POST":
        dados = {
            "nome_contratante": request.form.get("nome_contratante", "").strip(),
            "nacionalidade_contratante": request.form.get("nacionalidade_contratante", "").strip(),
            "estado_civil_contratante": request.form.get("estado_civil_contratante", "").strip(),
            "profissao_contratante": request.form.get("profissao_contratante", "").strip(),
            "cpf_contratante": request.form.get("cpf_contratante", "").strip(),
            "rg_contratante": request.form.get("rg_contratante", "").strip(),
            "uf_rg_contratante": request.form.get("uf_rg_contratante", "").strip(),
            "logradouro_contratante": request.form.get("logradouro_contratante", "").strip(),
            "complemento_contratante": request.form.get("complemento_contratante", "").strip(),
            "nome_executora": request.form.get("nome_executora", "").strip(),
            "tipo_pessoa_exec": request.form.get("tipo_pessoa_exec", "PF"),
            "cpf_executora": request.form.get("cpf_executora", "").strip(),
            "cnpj_executora": request.form.get("cnpj_executora", "").strip(),
            "endereco_executora": request.form.get("endereco_executora", "").strip(),
            "data_inicio": request.form.get("data_inicio", "").strip(),
            "prazo_execucao": request.form.get("prazo_execucao", "").strip(),
            "data_conclusao": request.form.get("data_conclusao", "").strip(),
            "valor_multa_fmt": request.form.get("valor_multa_fmt", "").strip(),
            "valor_multa_extenso": request.form.get("valor_multa_extenso", "").strip(),
            "cidade_assinatura": request.form.get("cidade_assinatura", "Anápolis-GO").strip(),
            "data_assinatura": request.form.get("data_assinatura", "").strip(),
        }

        servicos_desc = request.form.getlist("servico_desc[]")
        servicos_unid = request.form.getlist("servico_unidade[]")
        servicos_qtd  = request.form.getlist("servico_qtd[]")
        servicos_val  = request.form.getlist("servico_valor[]")
        servicos_prazo = request.form.getlist("servico_prazo[]")

        servicos_list = []
        valor_total = 0.0
        for i, desc in enumerate(servicos_desc):
            if desc.strip():
                val_str = servicos_val[i] if i < len(servicos_val) else ""
                try:
                    val_num = float(val_str.replace("R$", "").replace(".", "").replace(",", ".").strip())
                    qtd_num = float((servicos_qtd[i] if i < len(servicos_qtd) else "1").replace(",", ".") or "1")
                    valor_total += val_num * qtd_num
                except (ValueError, IndexError):
                    pass
                servicos_list.append({
                    "num": i + 1,
                    "descricao": desc.strip(),
                    "unidade": servicos_unid[i] if i < len(servicos_unid) else "M²",
                    "qtd": servicos_qtd[i] if i < len(servicos_qtd) else "1",
                    "valor": val_str,
                    "prazo": servicos_prazo[i] if i < len(servicos_prazo) else "",
                })

        if not dados["nome_contratante"]:
            flash("Nome do contratante é obrigatório.", "danger")
            return render_template("formulario_contrato.html", tipo=tipo,
                                   info=cg.CONTRATOS[tipo], dados=dados_iniciais,
                                   servicos_catalogo=servicos_catalogo,
                                   formas_pagamento=formas_pagamento)

        try:
            caminho = cg.gerar_contrato(tipo, dados, servicos_list)
            cid = db.salvar_contrato(
                tipo, cliente_id, executora_id,
                dados["nome_contratante"], dados["cpf_contratante"],
                dados["nome_executora"], dados.get("cpf_executora", ""),
                str(caminho), valor_total if valor_total > 0 else None,
            )
            for k in ["contrato_tipo", "contrato_cliente_id", "contrato_executora_id",
                      "form_cliente", "form_executora"]:
                session.pop(k, None)
            flash("Contrato gerado com sucesso!", "success")
            return redirect(url_for("download_contrato", cid=cid))
        except Exception as e:
            flash(f"Erro ao gerar contrato: {e}", "danger")

    return render_template("formulario_contrato.html", tipo=tipo,
                           info=cg.CONTRATOS[tipo], dados=dados_iniciais,
                           servicos_catalogo=servicos_catalogo,
                           formas_pagamento=formas_pagamento)


# ---------------------------------------------------------------------------
# Recibo
# ---------------------------------------------------------------------------

@app.route("/contrato/<int:cid>/recibo", methods=["GET", "POST"])
@login_required
def recibo_contrato(cid):
    contrato = db.buscar_contrato(cid)
    if not contrato:
        flash("Contrato não encontrado.", "danger")
        return redirect(url_for("index"))
    return render_template("recibo.html", contrato=contrato)


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

@app.route("/contrato/<int:cid>/download")
@login_required
def download_contrato(cid):
    contratos = db.listar_contratos()
    contrato = next((c for c in contratos if c["id"] == cid), None)
    if not contrato or not Path(contrato["caminho_arquivo"]).exists():
        flash("Arquivo não encontrado.", "danger")
        return redirect(url_for("index"))
    return send_file(contrato["caminho_arquivo"], as_attachment=True,
                     download_name=Path(contrato["caminho_arquivo"]).name)


# ---------------------------------------------------------------------------
# APIs JSON
# ---------------------------------------------------------------------------

@app.route("/api/cliente/<int:cid>")
@login_required
def api_cliente(cid):
    c = db.buscar_cliente(cid)
    return jsonify(c) if c else (jsonify({"erro": "não encontrado"}), 404)


@app.route("/api/executora/<int:eid>")
@login_required
def api_executora(eid):
    e = db.buscar_executora(eid)
    return jsonify(e) if e else (jsonify({"erro": "não encontrado"}), 404)


@app.route("/api/servico/<int:sid>")
@login_required
def api_servico(sid):
    s = db.buscar_servico(sid)
    return jsonify(s) if s else (jsonify({"erro": "não encontrado"}), 404)


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def processar_uploads(files):
    dados_ocr = []
    erros = []
    for arquivo in files:
        if arquivo and arquivo.filename and allowed_file(arquivo.filename):
            filename = secure_filename(arquivo.filename)
            caminho = UPLOAD_DIR / filename
            arquivo.save(str(caminho))
            try:
                resultado = ocr.extrair_dados_documento(str(caminho))
                dados_ocr.append(resultado)
            except Exception as e:
                erros.append(f"Erro em '{arquivo.filename}': {e}")
    for e in erros:
        flash(e, "warning")
    return ocr.mesclar_dados(dados_ocr) if dados_ocr else {}


with app.app_context():
    db.init_db()
    cg.preparar_todos_templates()


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
