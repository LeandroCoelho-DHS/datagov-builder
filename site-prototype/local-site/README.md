# Data Hawk Solution — protótipo local

Site estático (HTML + CSS puros, sem build, sem dependências) para testar o
protótipo no navegador.

## Como testar

**Opção 1 — abrir direto no navegador**
Dê duplo clique em `index.html` (ou arraste para uma aba do navegador).
Toda a navegação entre páginas funciona normalmente.

**Opção 2 — servidor local (recomendado, evita eventuais bloqueios de `file://`)**

```bash
cd local-site
python3 -m http.server 8000
```

Depois acesse `http://localhost:8000` no navegador.

## Arquivos

- `index.html` — página inicial
- `servicos.html` — Serviços & Produtos
- `contato.html` — Contato (com formulário de exemplo, sem backend)
- `style.css` — estilos compartilhados entre as três páginas

## Observações

- O formulário de contato não envia dados a lugar nenhum — é só uma
  demonstração visual/funcional dos campos (mostra um alerta ao enviar).
- Textos entre colchetes, como `[e-mail de contato]`, são placeholders a
  preencher com os dados reais da consultoria.
- As fontes (Newsreader e Archivo) e cores usam Google Fonts e `oklch()`,
  que exigem um navegador atual (Chrome, Edge, Firefox ou Safari recentes)
  e conexão com a internet para carregar as fontes.
