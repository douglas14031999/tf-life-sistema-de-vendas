# TF Life - Sistema de Vendas

Sistema completo de e-commerce PWA e gestão de pedidos integrado ao Google Sheets, Google Drive e WhatsApp.

---

## 🚀 Funcionalidades

### 🛍️ Loja Virtual (`index.html` / `TF_Life___Loja.html`)
- **PWA (Progressive Web App)**: Instalável na tela inicial de celulares e desktops, com funcionamento offline e detecção inteligente de app já instalado (sem exibir card repetido).
- **Reconhecimento Automático de Dispositivos**: Identifica o aparelho do cliente via token seguro (`tflife_device_id`). Ao abrir o carrinho, preenche automaticamente telefone, nome e endereço em 0ms.
- **Catálogo Otimizado**: Fotos com compressão, pré-carregamento com cache local e carregamento dinâmico ultrarrápido.
- **Variações de Produtos**: Seleção de fotos por cor, escolha de tamanhos e controle de grade de estoque em tempo real.
- **Sistema de Cupons de Desconto**:
  - Desconto em porcentagem (%) ou valor fixo (R$).
  - Validação de limite de uso e data de validade.
- **Finalização de Pedido**: Envio formatado direto para o WhatsApp da loja e gravação simultânea na planilha com baixa de estoque.

### ⚙️ Painel de Administração (`admin.html` / `TF_Life___Administração.html`)
- **Login Instantâneo em Pacote Único**: Retorna a autenticação e todo o catálogo de produtos/pedidos em uma única requisição.
- **Miniaturas Ultraleves**: Miniaturas da tabela solicitam resoluções compactas (`sz=w100`), reduzindo em ~98% o peso das imagens na listagem.
- **Acompanhamento de Pedidos**: Sincronização em tempo real de pedidos, faturamento e pendências com cálculo direto dos totais gerais.
- **Gestão Completa**: Produtos, Categorias, Clientes, Cupons e Configurações de WhatsApp e faixas de destaque.

### 📊 Backend de Alta Performance (`Code.gs`)
- **Memoização em Memória (`_memo`)**: Evita leituras redundantes de abas dentro da mesma execução.
- **Cache GZIP Compactado**: Compacta o catálogo da loja e contorna limitações de tamanho de cache no Apps Script.
- **Fila de Limpeza do Drive**: Exclusão assíncrona de fotos antigas sem travar a interface do usuário.
- **Concorrência Otimizada**: Requisições de leitura (catálogo, validação de cupons, busca de clientes) rodam sem bloqueio global de travas (`LockService`).

---

## ⚡ Como Atualizar e Publicar o Backend no Google Apps Script

Sempre que houver atualizações no arquivo [`Code.gs`](Code.gs), siga este passo a passo para atualizar a API do Google Apps Script sem alterar a URL `/exec`:

1. Abra a sua planilha **"TF Life – Loja (dados)"** no Google Sheets.
2. No menu superior, acesse **Extensões** > **Apps Script**.
3. Selecione todo o código no editor, apague e cole o conteúdo completo do arquivo [`Code.gs`](Code.gs).
4. No canto superior direito, clique em **Implantar** > **Gerenciar implantações**.
5. Clique no ícone de **lápis (Editar)** na implantação ativa:
   - Em **Versão**, selecione **"Nova versão"**.
   - Verifique se as configurações estão mantidas:
     - **Executar como**: *Eu (seu e-mail)*
     - **Quem tem acesso**: *Qualquer pessoa*
   - Clique em **Implantar**.
6. *(Recomendado para Máxima Performance)*:
   - No menu do editor do Apps Script, selecione a função **`instalarGatilhos`** na barra de ferramentas superior e clique em **Executar**.
   - Conceda a autorização da sua conta Google quando solicitado.
   - Isso ativará um gatilho automático a cada 5 minutos que:
     - Remove fotos antigas e deletadas do Google Drive em segundo plano sem travar o painel.
     - Mantém o catálogo da loja sempre aquecido no cache da nuvem para resposta instantânea aos clientes.

---

## 🌐 Publicação no GitHub Pages

O site é hospedado através do GitHub Pages:
1. No repositório, acesse **Settings** > **Pages**.
2. Em **Build and deployment**, selecione a branch `main` e a pasta `/ (root)`.
3. Domínio personalizado configurado: `https://tflife.provacanoa.tech` (ou `https://douglas14031999.github.io/tf-life-sistema-de-vendas/`).
