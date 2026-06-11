import os
import hashlib
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.environ.get("SUPABASE_DB_URL")


def get_db():
    conn = psycopg2.connect(DATABASE_URL)
    return conn


def _serialize(val):
    import datetime
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.isoformat()
    return val


def _rows_to_dicts(cursor):
    cols = [d[0] for d in cursor.description]
    return [{c: _serialize(v) for c, v in zip(cols, row)} for row in cursor.fetchall()]


def _row_to_dict(cursor, row):
    if row is None:
        return None
    cols = [d[0] for d in cursor.description]
    return {c: _serialize(v) for c, v in zip(cols, row)}


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            login TEXT NOT NULL UNIQUE,
            senha_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'visualizador',
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TIMESTAMP DEFAULT NOW()
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS categorias (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS unidades_medida (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL UNIQUE,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS formas_pagamento (
            id SERIAL PRIMARY KEY,
            descricao TEXT NOT NULL,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS servicos (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            descricao TEXT NOT NULL,
            unidade TEXT NOT NULL DEFAULT 'M²',
            preco_padrao REAL,
            percentual_padrao REAL,
            categoria TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM information_schema.columns
                WHERE table_name='servicos' AND column_name='percentual_padrao'
            ) THEN
                ALTER TABLE servicos ADD COLUMN percentual_padrao REAL;
            END IF;
        END $$;
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS clientes (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            nacionalidade TEXT DEFAULT 'brasileiro(a)',
            estado_civil TEXT DEFAULT 'solteiro(a)',
            profissao TEXT,
            cpf TEXT,
            rg TEXT,
            uf_rg TEXT,
            logradouro TEXT,
            complemento TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TIMESTAMP DEFAULT NOW()
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS executoras (
            id SERIAL PRIMARY KEY,
            tipo_pessoa TEXT NOT NULL DEFAULT 'PF',
            nome TEXT NOT NULL,
            cpf TEXT,
            cnpj TEXT,
            endereco TEXT,
            telefone TEXT,
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TIMESTAMP DEFAULT NOW()
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS contratos (
            id SERIAL PRIMARY KEY,
            tipo TEXT NOT NULL,
            cliente_id INTEGER REFERENCES clientes(id),
            executora_id INTEGER REFERENCES executoras(id),
            nome_contratante TEXT,
            cpf_contratante TEXT,
            nome_executora TEXT,
            cpf_executora TEXT,
            valor_total REAL,
            caminho_arquivo TEXT,
            criado_em TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("""
        DO $$ BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM information_schema.columns
                WHERE table_name='contratos' AND column_name='valor_total'
            ) THEN
                ALTER TABLE contratos ADD COLUMN valor_total REAL;
            END IF;
        END $$;
    """)

    # Dados padrão
    cur.execute("SELECT COUNT(*) FROM usuarios")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO usuarios (login, senha_hash, role) VALUES (%s, %s, %s)",
            ("admin", _hash_senha("admin123"), "admin"),
        )

    for unid in ["M²", "M³", "ML", "VB", "UN", "CJ", "HR", "KG", "L"]:
        cur.execute(
            "INSERT INTO unidades_medida (nome) VALUES (%s) ON CONFLICT (nome) DO NOTHING",
            (unid,),
        )

    for cat in ["Estrutura", "Alvenaria", "Revestimento", "Instalações", "Acabamento", "Outros"]:
        cur.execute(
            "INSERT INTO categorias (nome) VALUES (%s) ON CONFLICT (nome) DO NOTHING",
            (cat,),
        )

    for fp in [
        "À vista na conclusão",
        "50% no início / 50% na conclusão",
        "30 dias após conclusão",
        "Medição mensal",
        "Conforme cronograma",
    ]:
        cur.execute("SELECT COUNT(*) FROM formas_pagamento WHERE descricao=%s", (fp,))
        if cur.fetchone()[0] == 0:
            cur.execute("INSERT INTO formas_pagamento (descricao) VALUES (%s)", (fp,))

    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Autenticação
# ---------------------------------------------------------------------------

def _hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode()).hexdigest()


def autenticar_usuario(login: str, senha: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios WHERE login=%s AND ativo=1", (login,))
    row = _row_to_dict(cur, cur.fetchone())
    cur.close()
    conn.close()
    if row and row["senha_hash"] == _hash_senha(senha):
        return row
    return None


def listar_usuarios():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, login, role, ativo, criado_em FROM usuarios ORDER BY login")
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def criar_usuario(login: str, senha: str, role: str) -> int:
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO usuarios (login, senha_hash, role) VALUES (%s,%s,%s) RETURNING id",
        (login, _hash_senha(senha), role),
    )
    uid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return uid


def atualizar_usuario(uid: int, login: str, role: str, nova_senha: str = None):
    conn = get_db()
    cur = conn.cursor()
    if nova_senha:
        cur.execute(
            "UPDATE usuarios SET login=%s, role=%s, senha_hash=%s WHERE id=%s",
            (login, role, _hash_senha(nova_senha), uid),
        )
    else:
        cur.execute("UPDATE usuarios SET login=%s, role=%s WHERE id=%s", (login, role, uid))
    conn.commit()
    cur.close()
    conn.close()


def excluir_usuario(uid: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE usuarios SET ativo=0 WHERE id=%s", (uid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Categorias
# ---------------------------------------------------------------------------

def listar_categorias(apenas_ativas=True):
    conn = get_db()
    cur = conn.cursor()
    q = "SELECT * FROM categorias WHERE ativo=1 ORDER BY nome" if apenas_ativas else \
        "SELECT * FROM categorias ORDER BY nome"
    cur.execute(q)
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def criar_categoria(nome: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO categorias (nome) VALUES (%s) ON CONFLICT (nome) DO NOTHING RETURNING id",
        (nome,),
    )
    row = cur.fetchone()
    conn.commit()
    cur.close()
    conn.close()
    return row[0] if row else None


def excluir_categoria(cid: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE categorias SET ativo=0 WHERE id=%s", (cid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Unidades de Medida
# ---------------------------------------------------------------------------

def listar_unidades(apenas_ativas=True):
    conn = get_db()
    cur = conn.cursor()
    q = "SELECT * FROM unidades_medida WHERE ativo=1 ORDER BY nome" if apenas_ativas else \
        "SELECT * FROM unidades_medida ORDER BY nome"
    cur.execute(q)
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def criar_unidade(nome: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO unidades_medida (nome) VALUES (%s) ON CONFLICT (nome) DO NOTHING RETURNING id",
        (nome,),
    )
    row = cur.fetchone()
    conn.commit()
    cur.close()
    conn.close()
    return row[0] if row else None


def excluir_unidade(uid: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE unidades_medida SET ativo=0 WHERE id=%s", (uid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Formas de Pagamento
# ---------------------------------------------------------------------------

def listar_formas_pagamento(apenas_ativas=True):
    conn = get_db()
    cur = conn.cursor()
    q = "SELECT * FROM formas_pagamento WHERE ativo=1 ORDER BY descricao" if apenas_ativas else \
        "SELECT * FROM formas_pagamento ORDER BY descricao"
    cur.execute(q)
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def criar_forma_pagamento(descricao: str) -> int:
    conn = get_db()
    cur = conn.cursor()
    cur.execute("INSERT INTO formas_pagamento (descricao) VALUES (%s) RETURNING id", (descricao,))
    fid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return fid


def excluir_forma_pagamento(fid: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE formas_pagamento SET ativo=0 WHERE id=%s", (fid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Serviços
# ---------------------------------------------------------------------------

def listar_servicos(apenas_ativos=True):
    conn = get_db()
    cur = conn.cursor()
    if apenas_ativos:
        cur.execute("SELECT * FROM servicos WHERE ativo=1 ORDER BY categoria, nome")
    else:
        cur.execute("SELECT * FROM servicos ORDER BY categoria, nome")
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def buscar_servico(sid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM servicos WHERE id=%s", (sid,))
    row = _row_to_dict(cur, cur.fetchone())
    cur.close()
    conn.close()
    return row


def criar_servico(nome, descricao, unidade, preco_padrao, categoria, percentual_padrao=None):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO servicos (nome, descricao, unidade, preco_padrao, categoria, percentual_padrao)
           VALUES (%s,%s,%s,%s,%s,%s) RETURNING id""",
        (nome, descricao, unidade, preco_padrao, categoria, percentual_padrao),
    )
    sid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return sid


def atualizar_servico(sid, nome, descricao, unidade, preco_padrao, categoria, percentual_padrao=None):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """UPDATE servicos SET nome=%s, descricao=%s, unidade=%s, preco_padrao=%s,
           categoria=%s, percentual_padrao=%s WHERE id=%s""",
        (nome, descricao, unidade, preco_padrao, categoria, percentual_padrao, sid),
    )
    conn.commit()
    cur.close()
    conn.close()


def excluir_servico(sid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE servicos SET ativo=0 WHERE id=%s", (sid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Clientes
# ---------------------------------------------------------------------------

def listar_clientes(apenas_ativos=True):
    conn = get_db()
    cur = conn.cursor()
    if apenas_ativos:
        cur.execute("SELECT * FROM clientes WHERE ativo=1 ORDER BY nome")
    else:
        cur.execute("SELECT * FROM clientes ORDER BY nome")
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def buscar_cliente(cid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE id=%s", (cid,))
    row = _row_to_dict(cur, cur.fetchone())
    cur.close()
    conn.close()
    return row


def criar_cliente(dados: dict) -> int:
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO clientes
           (nome, nacionalidade, estado_civil, profissao, cpf, rg, uf_rg, logradouro, complemento)
           VALUES (%(nome)s, %(nacionalidade)s, %(estado_civil)s, %(profissao)s,
                   %(cpf)s, %(rg)s, %(uf_rg)s, %(logradouro)s, %(complemento)s)
           RETURNING id""",
        dados,
    )
    cid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return cid


def atualizar_cliente(cid, dados: dict):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """UPDATE clientes SET nome=%(nome)s, nacionalidade=%(nacionalidade)s,
           estado_civil=%(estado_civil)s, profissao=%(profissao)s, cpf=%(cpf)s,
           rg=%(rg)s, uf_rg=%(uf_rg)s, logradouro=%(logradouro)s,
           complemento=%(complemento)s WHERE id=%(id)s""",
        {**dados, "id": cid},
    )
    conn.commit()
    cur.close()
    conn.close()


def excluir_cliente(cid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE clientes SET ativo=0 WHERE id=%s", (cid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Executoras
# ---------------------------------------------------------------------------

def listar_executoras(apenas_ativas=True):
    conn = get_db()
    cur = conn.cursor()
    if apenas_ativas:
        cur.execute("SELECT * FROM executoras WHERE ativo=1 ORDER BY nome")
    else:
        cur.execute("SELECT * FROM executoras ORDER BY nome")
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def buscar_executora(eid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM executoras WHERE id=%s", (eid,))
    row = _row_to_dict(cur, cur.fetchone())
    cur.close()
    conn.close()
    return row


def criar_executora(dados: dict) -> int:
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO executoras (tipo_pessoa, nome, cpf, cnpj, endereco, telefone)
           VALUES (%(tipo_pessoa)s, %(nome)s, %(cpf)s, %(cnpj)s, %(endereco)s, %(telefone)s)
           RETURNING id""",
        dados,
    )
    eid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return eid


def atualizar_executora(eid, dados: dict):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """UPDATE executoras SET tipo_pessoa=%(tipo_pessoa)s, nome=%(nome)s, cpf=%(cpf)s,
           cnpj=%(cnpj)s, endereco=%(endereco)s, telefone=%(telefone)s WHERE id=%(id)s""",
        {**dados, "id": eid},
    )
    conn.commit()
    cur.close()
    conn.close()


def excluir_executora(eid):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE executoras SET ativo=0 WHERE id=%s", (eid,))
    conn.commit()
    cur.close()
    conn.close()


# ---------------------------------------------------------------------------
# Contratos
# ---------------------------------------------------------------------------

def salvar_contrato(tipo, cliente_id, executora_id, nome_contratante, cpf_contratante,
                    nome_executora, cpf_executora, caminho, valor_total=None):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO contratos
           (tipo, cliente_id, executora_id, nome_contratante, cpf_contratante,
            nome_executora, cpf_executora, caminho_arquivo, valor_total)
           VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
        (tipo, cliente_id, executora_id, nome_contratante, cpf_contratante,
         nome_executora, cpf_executora, caminho, valor_total),
    )
    cid = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return cid


def buscar_contrato(cid: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """SELECT c.*, cl.nome as cliente_nome, cl.cpf as cliente_cpf,
                  cl.logradouro as cliente_logradouro,
                  e.endereco as executora_endereco, e.cnpj as executora_cnpj,
                  e.cpf as executora_cpf_doc
           FROM contratos c
           LEFT JOIN clientes cl ON c.cliente_id = cl.id
           LEFT JOIN executoras e ON c.executora_id = e.id
           WHERE c.id=%s""",
        (cid,),
    )
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows[0] if rows else None


def listar_contratos(cliente_id=None):
    conn = get_db()
    cur = conn.cursor()
    if cliente_id:
        cur.execute(
            """SELECT c.*, cl.nome as cliente_nome
               FROM contratos c LEFT JOIN clientes cl ON c.cliente_id = cl.id
               WHERE c.cliente_id=%s ORDER BY c.criado_em DESC""",
            (cliente_id,),
        )
    else:
        cur.execute(
            """SELECT c.*, cl.nome as cliente_nome
               FROM contratos c LEFT JOIN clientes cl ON c.cliente_id = cl.id
               ORDER BY c.criado_em DESC LIMIT 100"""
        )
    rows = _rows_to_dicts(cur)
    cur.close()
    conn.close()
    return rows


def dashboard_cliente(cid: int) -> dict:
    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        "SELECT COUNT(*), COALESCE(SUM(valor_total),0) FROM contratos WHERE cliente_id=%s",
        (cid,),
    )
    total_contratos, total_gasto = cur.fetchone()

    cur.execute(
        "SELECT DISTINCT nome_executora FROM contratos WHERE cliente_id=%s AND nome_executora IS NOT NULL",
        (cid,),
    )
    executoras = [r[0] for r in cur.fetchall()]

    cur.execute(
        "SELECT tipo, COUNT(*) FROM contratos WHERE cliente_id=%s GROUP BY tipo ORDER BY COUNT(*) DESC",
        (cid,),
    )
    por_tipo = [{"tipo": r[0], "qtd": r[1]} for r in cur.fetchall()]

    cur.execute(
        """SELECT TO_CHAR(criado_em,'YYYY-MM') as mes, COALESCE(SUM(valor_total),0)
           FROM contratos WHERE cliente_id=%s
           GROUP BY mes ORDER BY mes DESC LIMIT 12""",
        (cid,),
    )
    por_mes = list(reversed([{"mes": r[0], "total": float(r[1])} for r in cur.fetchall()]))

    cur.close()
    conn.close()
    return {
        "total_contratos": total_contratos,
        "total_gasto": float(total_gasto),
        "executoras": executoras,
        "por_tipo": por_tipo,
        "por_mes": por_mes,
    }


def resumo_geral() -> dict:
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM clientes WHERE ativo=1")
    total_clientes = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM executoras WHERE ativo=1")
    total_executoras = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM servicos WHERE ativo=1")
    total_servicos = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*), COALESCE(SUM(valor_total),0) FROM contratos")
    total_contratos, valor_total = cur.fetchone()

    cur.execute(
        "SELECT tipo, COUNT(*) FROM contratos GROUP BY tipo ORDER BY COUNT(*) DESC LIMIT 1"
    )
    row = cur.fetchone()
    tipo_mais_usado = row[0] if row else "—"

    cur.execute(
        """SELECT TO_CHAR(criado_em,'YYYY-MM') as mes, COUNT(*), COALESCE(SUM(valor_total),0)
           FROM contratos GROUP BY mes ORDER BY mes DESC LIMIT 12"""
    )
    por_mes = list(reversed([{"mes": r[0], "qtd": r[1], "total": float(r[2])} for r in cur.fetchall()]))

    cur.execute(
        "SELECT tipo, COUNT(*) FROM contratos GROUP BY tipo ORDER BY COUNT(*) DESC"
    )
    por_tipo = [{"tipo": r[0], "qtd": r[1]} for r in cur.fetchall()]

    cur.close()
    conn.close()
    return {
        "total_clientes": total_clientes,
        "total_executoras": total_executoras,
        "total_servicos": total_servicos,
        "total_contratos": total_contratos,
        "valor_total": float(valor_total),
        "tipo_mais_usado": tipo_mais_usado,
        "por_mes": por_mes,
        "por_tipo": por_tipo,
    }
