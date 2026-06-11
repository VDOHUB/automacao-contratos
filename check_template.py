import zipfile, re, sys

tipo = sys.argv[1] if len(sys.argv) > 1 else "estrutura"

paths = {
    "estrutura": "0\xb0 - Modelos de Contratos/MODELO CONTRATO LUIZ CIVIL - Estrutura (1).docx",
    "alvenaria": "0\xb0 - Modelos de Contratos/MODELO CONTRATO LUIZ - Alvenaria 1\xb0PAV (1).docx",
    "arrimo":    "0\xb0 - Modelos de Contratos/MODELO CONTRATO LUIZ CIVIL - Arrimo (1).docx",
}

path = paths.get(tipo, paths["estrutura"])
print(f"Lendo: {path}")

with zipfile.ZipFile(path) as z:
    xml = z.read('word/document.xml').decode('utf-8')

# Merge red runs first (same as preprocessor)
para_re = re.compile(r'(<w:p[ >].*?</w:p>)', re.DOTALL)
run_re = re.compile(r'<w:r[ >].*?</w:r>', re.DOTALL)

merged_texts = []
for para_m in para_re.finditer(xml):
    para = para_m.group(1)
    if 'w:highlight w:val="red"' not in para:
        continue
    # collect consecutive red runs
    groups = []
    for run in run_re.findall(para):
        is_red = 'w:highlight w:val="red"' in run
        if groups and groups[-1][0] == is_red:
            groups[-1][1].append(run)
        else:
            groups.append((is_red, [run]))
    for is_red, runs in groups:
        if not is_red:
            continue
        texts = []
        for r in runs:
            texts.extend(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r))
        merged = ''.join(texts).strip()
        if merged:
            merged_texts.append(merged)

print(f"\n=== Textos VERMELHOS (apos merge) em '{tipo}' ===")
for t in merged_texts:
    print(repr(t))
