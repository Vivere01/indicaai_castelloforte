# 🪒 Indica Aí — Castello Forte Barbearia

Formulário web do **Programa de Indicação** da Castello Forte Barbearia.  
Permite que assinantes indiquem até 5 amigos, com os dados sendo salvos automaticamente na planilha do Google Sheets.

---

## 📁 Arquivos

| Arquivo | Descrição |
|---|---|
| `index.html` | Formulário web completo (front-end) |
| `code.gs` | Google Apps Script — backend que recebe e salva os dados |

---

## 🚀 Como configurar a integração com o Google Sheets

### 1. Publicar o Apps Script

1. Acesse [script.google.com](https://script.google.com/) e clique em **Novo projeto**
2. Apague o código padrão e cole o conteúdo de `code.gs`
3. Clique em **Implantar** → **Nova implantação**
4. Configure:
   - **Tipo:** App da Web
   - **Executar como:** Eu (sua conta Google)
   - **Quem pode acessar:** Qualquer pessoa
5. Clique em **Implantar** e copie a **URL do Web App** gerada

### 2. Conectar ao formulário

Abra `index.html` e localize esta linha no `<script>`:

```js
var SCRIPT_URL = "COLE_AQUI_A_URL_DO_WEB_APP";
```

Substitua `"COLE_AQUI_A_URL_DO_WEB_APP"` pela URL copiada no passo anterior.

### 3. Estrutura da planilha

O script cria automaticamente uma aba chamada **"Indicações"** na planilha com as colunas:

| Data/Hora | Nome Assinante | WhatsApp Assinante | Indicado 1 — Nome | Indicado 1 — WhatsApp | ... | Status |
|---|---|---|---|---|---|---|

> A planilha utilizada: `1uun-641r9XMwCLrcrrDT0h-Iurh0_kP_VipnsjKQ_IY`

---

## 💡 Funcionalidades do formulário

- ✅ Coleta nome e WhatsApp do assinante
- ✅ Coleta dados de até 5 indicados (1 obrigatório, demais opcionais)
- ✅ Máscara automática de telefone: `(99) 99999-9999`
- ✅ Validação de campos obrigatórios
- ✅ Estado de carregamento durante o envio
- ✅ Tela de sucesso após envio
- ✅ Banner de erro se o envio falhar
- ✅ Layout responsivo (mobile-first)

---

## 🌐 Como hospedar

O formulário é um arquivo HTML estático — pode ser hospedado em:

- **GitHub Pages** (grátis, basta habilitar nas configurações do repositório)
- **Netlify / Vercel** (arraste e solte a pasta)
- **Qualquer servidor web** (Apache, Nginx etc.)

### GitHub Pages (recomendado)

1. Vá em **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / pasta: `/ (root)`
4. O formulário estará disponível em: `https://Vivere01.github.io/indicaai_castelloforte/`

---

## ⚠️ Observações

- O Apps Script deve ser publicado com acesso **"Qualquer pessoa"** para aceitar envios sem login.
- Por usar `mode: 'no-cors'` no fetch, o navegador não lê a resposta do servidor — o envio é considerado bem-sucedido ao não gerar erro de rede.
- Mantenha a URL do script em segredo para evitar abusos.
