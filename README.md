# Consign 🗄️

Sistema web para gestão operacional de guarda-volumes hospitalares, com controle de armários para visitantes e acompanhantes, emissão de termo de responsabilidade, evidências por imagem e auditoria de ponta a ponta.

---

## 📌 Descrição objetiva

O **Consign** é um WebApp construído com **Google Apps Script + Google Sheets** para organizar a rotina de guarda-volumes em ambiente hospitalar. O sistema centraliza cadastro, movimentações, devoluções, termo assinado e histórico operacional em uma única interface.

---

## ❗ Problema que o sistema resolve

Em operações hospitalares, o controle de armários costuma ser distribuído entre anotações manuais, planilhas paralelas e comunicação informal. Isso gera:

- baixa rastreabilidade de quem usou cada armário;
- risco de perda de informações e evidências;
- lentidão na liberação de armários;
- dificuldade para auditoria e conferência de plantão;
- falhas no registro de termos e devoluções.

O **Consign** padroniza esse fluxo com registro estruturado, validações operacionais e histórico auditável.

---

## ⚙️ Principais funcionalidades

- **Gestão de armários em tempo real** (livre, em uso, próximo do horário, vencido, contingência).
- **Cadastro de visitantes e acompanhantes** com vínculo de paciente, leito, unidade e responsável.
- **Termo de responsabilidade com fluxo guiado** (paciente, acompanhante, revisão e assinatura).
- **Geração de PDF do termo** com armazenamento no Google Drive.
- **Registro obrigatório de imagens** (volumes, movimentações, entrega e contingências).
- **Indexação de evidências** na aba `Registro de Imagens` para consulta posterior.
- **Módulo de liberações** por período, paciente e prontuário.
- **Achados e perdidos** com histórico separado.
- **Painel de monitoramento (dashboard)** com filtros e indicadores operacionais.
- **Auditoria de eventos (LOGS)** para rastrear ações críticas.

---

## 🧱 Tecnologias utilizadas

- **Google Apps Script** (back-end e publicação do WebApp)
- **Google Sheets** (base de dados operacional)
- **HTML5 + CSS3 + JavaScript** (interface do usuário)
- **Google Drive API via Apps Script** (armazenamento de PDFs e imagens)
- **jsPDF** (exportação de PDF no front-end, quando aplicável)

---

## 🗂️ Estrutura do projeto

```text
Consign/
├── CODE.gs                # Back-end Apps Script (rotas, regras e persistência)
├── index.html             # Tela de login/entrada
├── script.html            # Lógica front-end e chamadas para o WebApp
├── style.html             # Estilos da interface
├── MANUAL_DO_USUARIO.md   # Guia funcional detalhado da operação
├── relatorio.md           # Relatório técnico de revisão
└── README.md              # Documentação principal do projeto
```

---

## 🔄 Fluxo de funcionamento

1. **Login** do usuário autorizado.
2. **Seleção de perfil e unidade** para contextualizar a operação.
3. **Cadastro de uso do armário** (visitante/acompanhante).
4. **Registro de evidências obrigatórias** nos pontos críticos do processo.
5. **Emissão/finalização do termo** (quando aplicável) com geração de PDF.
6. **Liberação do armário** e atualização automática dos status.
7. **Auditoria e consulta histórica** via logs, históricos e registro de imagens.

---

## 🖼️ Capturas de tela

> Substitua os blocos abaixo pelas imagens reais do sistema.

### Tela de login
`![Login do Consign](./docs/screenshots/login.png)`

### Monitor operacional
`![Monitor de armários](./docs/screenshots/monitor.png)`

### Termo de responsabilidade
`![Fluxo do termo](./docs/screenshots/termo.png)`

### Registro de imagens / evidências
`![Registro de evidências](./docs/screenshots/evidencias.png)`

---

## ▶️ Como executar

### Pré-requisitos

- Conta Google com acesso ao **Apps Script**, **Sheets** e **Drive**.
- Planilha principal do projeto com permissões de edição.

### Passo a passo

1. Crie (ou copie) um projeto no **Google Apps Script**.
2. Adicione os arquivos do repositório (`CODE.gs`, `index.html`, `script.html`, `style.html`).
3. Configure os IDs de pastas/planilhas nas constantes do `CODE.gs`.
4. Verifique se as abas operacionais existem (ou se são criadas automaticamente).
5. Publique em **Implantar → Implantar como aplicativo da web**.
6. Defina o acesso conforme sua política interna (usuários autorizados).
7. Abra a URL do WebApp e valide os fluxos principais (cadastro, termo, liberação e logs).

---

## 🚀 Melhorias futuras

- controle de sessão com token assinado e expiração no servidor;
- endurecimento de segurança (autorização por ação, proteção anti-clickjacking);
- trilha de auditoria expandida com métricas de SLA por unidade;
- dashboard executivo com séries históricas e comparativos por turno;
- notificações automáticas para armários vencidos e pendências críticas.

---

## 👤 Autor

Desenvolvido por **Lucas Leal**.  
Projeto focado em operação real, rastreabilidade e eficiência em ambiente hospitalar.
