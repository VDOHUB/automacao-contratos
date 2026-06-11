"""
Extrai dados de documentos pessoais (RG, CNH, CPF, CNPJ) via Groq (gratuito).
Suporta imagens (JPG, PNG, WEBP) e PDF.
"""

from groq import Groq
import base64
import json
import os
import re
import tempfile
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

PROMPT = """Analise a imagem deste documento brasileiro e extraia os dados em JSON.
Retorne APENAS o JSON, sem texto adicional, sem markdown, sem ```json.

Para RG ou Identidade:
{"tipo": "rg", "nome": "", "cpf": "", "rg": "", "data_nascimento": "", "naturalidade": "", "filiacao_mae": "", "filiacao_pai": ""}

Para CNH:
{"tipo": "cnh", "nome": "", "cpf": "", "rg": "", "data_nascimento": "", "naturalidade": ""}

Para comprovante de endereço:
{"tipo": "endereco", "logradouro": "", "numero": "", "complemento": "", "bairro": "", "cidade": "", "uf": "", "cep": ""}

Para cartão CPF:
{"tipo": "cpf", "nome": "", "cpf": ""}

Para cartão CNPJ ou contrato social:
{"tipo": "cnpj", "razao_social": "", "cnpj": "", "endereco": "", "nome": ""}

Regras:
- Se não conseguir identificar algum campo, deixe como string vazia ""
- Formate CPF como XXX.XXX.XXX-XX
- Formate CNPJ como XX.XXX.XXX/XXXX-XX
- Mantenha o RG no formato original do documento
- O nome deve estar em letras maiúsculas como aparece no documento
"""


def _pdf_para_imagens(caminho_pdf: str) -> list[str]:
    """Converte páginas de PDF em imagens temporárias PNG. Retorna lista de caminhos."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise RuntimeError("pypdfium2 não instalado. Execute: pip install pypdfium2")

    pdf = pdfium.PdfDocument(caminho_pdf)
    caminhos = []
    tmpdir = tempfile.mkdtemp()

    for i, page in enumerate(pdf):
        bitmap = page.render(scale=2)  # 2x para melhor qualidade OCR
        pil_img = bitmap.to_pil()
        out_path = os.path.join(tmpdir, f"page_{i}.png")
        pil_img.save(out_path)
        caminhos.append(out_path)
        if i >= 2:  # máximo 3 páginas por PDF
            break

    pdf.close()
    return caminhos


def _extrair_de_imagem(caminho: str) -> dict:
    """Envia uma imagem para o Groq e retorna os dados extraídos."""
    path = Path(caminho)
    mime_map = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp",
    }
    mime_type = mime_map.get(path.suffix.lower(), "image/jpeg")

    with open(caminho, "rb") as f:
        image_b64 = base64.b64encode(f.read()).decode("utf-8")

    response = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_b64}"}},
                {"type": "text", "text": PROMPT},
            ],
        }],
        temperature=0,
        max_tokens=1024,
    )

    raw = response.choices[0].message.content.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"erro": "Não foi possível extrair dados", "raw": raw}


def extrair_dados_documento(caminho_arquivo: str) -> dict:
    """Extrai dados de imagem ou PDF. Para PDF, processa até 3 páginas."""
    path = Path(caminho_arquivo)

    if path.suffix.lower() == ".pdf":
        imagens = _pdf_para_imagens(caminho_arquivo)
        if not imagens:
            return {"erro": "PDF sem páginas extraídas"}
        resultados = [_extrair_de_imagem(img) for img in imagens]
        return mesclar_dados(resultados)
    else:
        return _extrair_de_imagem(caminho_arquivo)


def mesclar_dados(docs: list) -> dict:
    merged = {}
    for doc in docs:
        if "erro" in doc:
            continue
        for key, value in doc.items():
            if key == "tipo":
                continue
            # CNPJ: mapa razao_social → nome para uniformidade
            if key == "razao_social" and value and not merged.get("nome"):
                merged["nome"] = value
                continue
            if value and not merged.get(key):
                merged[key] = value
    return merged
