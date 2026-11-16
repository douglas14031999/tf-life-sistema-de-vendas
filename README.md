# TF Life - Sistema de Vendas

Sistema completo de e-commerce e gestão de pedidos integrado ao Google Sheets e WhatsApp.

## 🚀 Funcionalidades

### 🛍️ Loja Virtual (`index.html` / `TF_Life___Loja.html`)
- Catálogo de produtos com fotos, categorias e variações de tamanho/cor.
- Carrinho lateral moderno com miniaturas, ajuste de quantidade e botão de remoção.
- Identificação automática do cliente ao digitar o número de telefone (preenche nome e endereço salvos).
- Sistema de **Cupons de Desconto**:
  - Desconto em porcentagem (%) ou valor fixo (R$).
  - Validação de limite de uso e data de expiração.
- Finalização de pedido com envio formatado direto para o WhatsApp e gravação simultânea na planilha.

### ⚙️ Painel de Administração (`admin.html` / `TF_Life___Administração.html`)
- Gerenciamento de Produtos (adicionar, editar, excluir, estoque e fotos).
- Acompanhamento de Pedidos em tempo real.
- Base de Clientes cadastrados.
- Gestão de Cupons de Desconto com regras customizadas.

### 📊 Backend (`Code.gs`)
- Script do Google Apps Script conectado ao Google Sheets com tabelas para:
  - `Produtos`
  - `Pedidos`
  - `Clientes`
  - `Cupons`

---

## 🌐 Publicação no GitHub Pages

Para acessar o site online através do GitHub Pages:
1. No repositório, acesse **Settings** > **Pages**.
2. Em **Build and deployment**, selecione a branch `main` e a pasta `/ (root)`.
3. Clique em **Save**.
4. A loja estará disponível em: `https://douglas14031999.github.io/tf-life-sistema-de-vendas/`
