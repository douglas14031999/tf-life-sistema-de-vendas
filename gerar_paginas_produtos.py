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
        {"id": "pmwjz80ad", "nome": "Camisa DRYFit", "preco": 150, "descricao": "Uma Linda Camisa de tecido que transpira.", "imagem": "https://lh3.googleusercontent.com/d/1Pz2I7E2Yc4r2L_w2EFBp-lrGnPsjI_CV=w900"},
        {"id": "p001", "nome": "Copo térmico 500 ml", "preco": 59.9, "descricao": "Mantém gelado por horas", "imagem": DEFAULT_OG},
        {"id": "p002", "nome": "Garrafa squeeze 900 ml", "preco": 39.9, "descricao": "Tampa antivazamento", "imagem": DEFAULT_OG},
        {"id": "p003", "nome": "Quentinha frango e batata-doce", "preco": 24.9, "descricao": "Feita no dia, 350 g", "imagem": DEFAULT_OG},
        {"id": "p007", "nome": "Camiseta dry-fit TF Life", "preco": 79.9, "descricao": "P ao GG, várias cores", "imagem": DEFAULT_OG}
    ]

for p in produtos:
    p_id = p.get('id')
    nome = p.get('nome', 'Produto')
    preco = p.get('preco', 0)
    preco_fmt = f"R$ {float(preco):.2f}".replace('.', ',')
    desc = p.get('descricao_completa') or p.get('descricao') or f"Confira {nome} na TF Life!"
    desc_clean = desc.replace('"', '&quot;').replace('\n', ' ')

    img_raw = p.get('imagem')
    if not img_raw and p.get('variacoes') and len(p['variacoes']) > 0:
        img_raw = p['variacoes'][0].get('imagem')
    
    img_url = get_direct_img(img_raw)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>{nome} · {preco_fmt} - TF Life</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta property="og:type" content="product">
  <meta property="og:site_name" content="TF Life">
  <meta property="og:title" content="{nome} · {preco_fmt}">
  <meta property="og:description" content="{desc_clean}">
  <meta property="og:image" content="{img_url}">
  <meta property="og:image:secure_url" content="{img_url}">
  <meta property="og:image:type" content="image/jpeg">
  <meta property="og:image:width" content="900">
  <meta property="og:image:height" content="900">
  <meta property="og:image:alt" content="{nome} - TF Life">
  <meta property="og:url" content="https://tflife.provacanoa.tech/p/{p_id}.html">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{nome} · {preco_fmt}">
  <meta name="twitter:description" content="{desc_clean}">
  <meta name="twitter:image" content="{img_url}">
  <meta http-equiv="refresh" content="0;url=/index.html?p={p_id}">
  <script>
    var dest = '/index.html?p={p_id}';
    if (location.search) {{
      var s = location.search.replace(/^\\?/, '&');
      dest += s;
    }}
    window.location.replace(dest);
  </script>
</head>
<body style="font-family:sans-serif;text-align:center;padding:40px;color:#1E4D35;background:#FAFAF7">
  <h2>{nome}</h2>
  <p style="color:#666">{desc_clean}</p>
  <p><strong>{preco_fmt}</strong></p>
  <p>Redirecionando para a loja...</p>
  <p><a href="/index.html?p={p_id}" style="color:#1E4D35;font-weight:bold">Clique aqui se não for redirecionado</a></p>
</body>
</html>"""

    file_path = os.path.join(OUTPUT_DIR, f"{p_id}.html")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(html)

print("Páginas geradas com sucesso!")
