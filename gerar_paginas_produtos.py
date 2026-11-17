import os
import json
import urllib.request
import re

API_URL = "https://script.google.com/macros/s/AKfycbwFc9yH6Z-Rfuzr4R-3Dw05v8M0LKDbfxbmjOKEzecvaxMVlfEf5GPVn2Fe8sdiEhXF/exec"
DEFAULT_OG = "https://tflife.provacanoa.tech/og-image.jpg"
OUTPUT_DIR = "p"

def get_direct_img(url):
    if not url:
        return DEFAULT_OG
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m:
        file_id = m.group(1)
        return f"https://lh3.googleusercontent.com/d/{file_id}=w900"
    return url

os.makedirs(OUTPUT_DIR, exist_ok=True)

req = urllib.request.Request(API_URL, headers={'User-Agent': 'Mozilla/5.0'})
try:
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        produtos = data.get('produtos', [])
except Exception as e:
    print("Erro ao buscar produtos da API:", e)
    produtos = []

if not produtos:
    produtos = [
        {"id": "pmwjz80ad", "nome": "Camisa DRYFit", "preco": 150, "descricao": "Uma Linda Camisa de tecido que transpira.", "imagem": "https://lh3.googleusercontent.com/d/1Pz2I7E2Yc4r2L_w2EFBp-lrGnPsjI_CV=w900"}
    ]

def gerar_html(p_id, file_suffix, nome, preco, desc, img_raw, foto_idx=0, cor_nome=""):
    preco_fmt = f"R$ {float(preco):.2f}".replace('.', ',')
    desc_clean = (desc or f"Confira {nome} na TF Life!").replace('"', '&quot;').replace('\n', ' ')
    img_url = get_direct_img(img_raw)
    
    titulo_og = f"{nome} · {preco_fmt}"
    if cor_nome:
        titulo_og = f"{nome} · {cor_nome} · {preco_fmt}"

    redirect_url = f"/index.html?p={p_id}"
    if foto_idx > 0:
        redirect_url += f"&foto={foto_idx}"

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>{titulo_og} - TF Life</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta property="og:type" content="product">
  <meta property="og:site_name" content="TF Life">
  <meta property="og:title" content="{titulo_og}">
  <meta property="og:description" content="{desc_clean}">
  <meta property="og:image" content="{img_url}">
  <meta property="og:image:secure_url" content="{img_url}">
  <meta property="og:image:type" content="image/jpeg">
  <meta property="og:image:width" content="900">
  <meta property="og:image:height" content="900">
  <meta property="og:image:alt" content="{titulo_og} - TF Life">
  <meta property="og:url" content="https://tflife.provacanoa.tech/p/{p_id}{file_suffix}.html">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{titulo_og}">
  <meta name="twitter:description" content="{desc_clean}">
  <meta name="twitter:image" content="{img_url}">
  <meta http-equiv="refresh" content="0;url={redirect_url}">
  <script>
    var dest = '{redirect_url}';
    if (location.search) {{
      var s = location.search.replace(/^\\?/, '&');
      dest += s;
    }}
    window.location.replace(dest);
  </script>
</head>
<body style="font-family:sans-serif;text-align:center;padding:40px;color:#1E4D35;background:#FAFAF7">
  <h2>{nome} {f'({cor_nome})' if cor_nome else ''}</h2>
  <p style="color:#666">{desc_clean}</p>
  <p><strong>{preco_fmt}</strong></p>
  <p>Redirecionando para a loja...</p>
  <p><a href="{redirect_url}" style="color:#1E4D35;font-weight:bold">Clique aqui se não for redirecionado</a></p>
</body>
</html>"""
    
    file_path = os.path.join(OUTPUT_DIR, f"{p_id}{file_suffix}.html")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Gerado: {file_path} -> Imagem: {img_url}")

for p in produtos:
    p_id = p.get('id')
    nome = p.get('nome', 'Produto')
    preco_base = p.get('preco', 0)
    desc = p.get('descricao_completa') or p.get('descricao') or ""
    vars_list = p.get('variacoes') or []
    
    img_capa = p.get('imagem') or (vars_list[0].get('imagem') if vars_list and len(vars_list) > 0 else "")
    cor_capa = (vars_list[0].get('nome') if vars_list and len(vars_list) > 0 else "")

    # 1. Página principal do produto: p/{id}.html
    gerar_html(p_id, "", nome, preco_base, desc, img_capa, foto_idx=0, cor_nome="")
    gerar_html(p_id, "_0", nome, preco_base, desc, img_capa, foto_idx=0, cor_nome=cor_capa)

    # 2. Páginas por variação/foto: p/{id}_{idx}.html
    if vars_list and len(vars_list) > 1:
        for idx, v in enumerate(vars_list):
            if idx == 0:
                continue
            v_img = v.get('imagem') or img_capa
            v_preco = v.get('preco') if v.get('preco') is not None else preco_base
            v_nome = v.get('nome') or f"Opção {idx+1}"
            gerar_html(p_id, f"_{idx}", nome, v_preco, desc, v_img, foto_idx=idx, cor_nome=v_nome)

print("Todas as páginas com variações foram geradas com sucesso!")
