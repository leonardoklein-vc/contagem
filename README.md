# 🗂️ Sistema Integrado de Contagem de Documentos de Arquivos

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Bootstrap](https://img.shields.io/badge/Bootstrap_5-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)
![OCR API](https://img.shields.io/badge/API_Integration-OCR.space-ff69b4?style=for-the-badge)

Uma Single Page Application (SPA) desenvolvida 100% no ecossistema Google Workspace para transformar a gestão documental hospitalar. O sistema substitui uma infraestrutura legada de mais de 70 abas de planilhas por um banco de dados centralizado com interface web moderna.

---

## 🚨 O Desafio Operacional
O controle de documentos e prontuários era feito de forma manual, disperso em dezenas de guias. 
* **Fragmentação:** Impossível extrair métricas cruzadas sem trabalho manual intensivo.
* **Lentidão:** O volume de dados travava a interface padrão das planilhas.
* **Retrabalho:** Lançamentos repetitivos e propensos a erro humano.

---

## 💡 A Solução e Arquitetura
Como arquiteto da solução, desenvolvi um sistema onde o Google Apps Script (GAS) atua como backend/API, servindo um frontend responsivo (HTML/JS/Bootstrap) que os usuários operam sem nunca tocar na planilha base.

### ✨ Features de Destaque

* **🤖 Integração OCR (Optical Character Recognition):** Conexão com a API do *OCR.space* para ler documentos escaneados e preencher formulários de lançamento automaticamente, reduzindo o *data entry* manual.
* **🧠 Autocomplete Inteligente (Algoritmo VIP):** O sistema varre o histórico de uso e coloca os "Tipos Documentais" mais frequentes no topo das sugestões de busca, otimizando o tempo de preenchimento.
* **📊 Analytics Dinâmico (Chart.js):** Geração de gráficos em tempo real (Evolução Mensal, Rankings por Local/Tipo) consumindo diretamente a base de dados consolidada.
* **✂️ Estúdio PDF Integrado:** Ferramenta injetada via *Lazy Load* (carrega apenas quando solicitada para poupar memória) para manipulação de PDFs diretamente no navegador usando `pdf-lib`.
* **🏷️ Motor de Etiquetas:** Geração visual e impressão automatizada de rótulos de rastreamento com cálculo lógico de numeração e paginação exata para impressão física.*

<img width="1438" height="744" alt="fac4d59d-ce62-4271-9fa1-b3dba700fe8e" src="https://github.com/user-attachments/assets/f732a9a4-ad81-439d-adff-3a14e8be6d7f" />
---

## ⚙️ Engenharia e Performance
Para garantir fluidez em um ambiente web servido pelo GAS:
1. **Cache e Otimização de Fila:** Uso intenso do `CacheService` e processamento em lote para evitar chamadas excessivas à API do Google Sheets.
2. **Auto-Save (Drafts):** Implementação de `localStorage` no frontend. Se o navegador for fechado acidentalmente, o rascunho dos lançamentos é recuperado na próxima sessão.
3. **UX Focada em Produtividade:** Navegação de formulários otimizada para uso exclusivo do teclado (Eventos de `Enter` customizados pulando entre campos).
