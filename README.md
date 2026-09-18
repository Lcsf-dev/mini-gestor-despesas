# 💰 Mini Gestor de Despesas Rápidas

Aplicação desktop desenvolvida em Visual Basic 6 (VB6) integrada ao Microsoft SQL Server, utilizando Stored Procedures e ADO (ActiveX Data Objects). 

O objetivo do projeto é fornecer um utilitário simples e rápido para registro e acompanhamento de gastos pessoais diários.

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Visual Basic 6 (VB6)
- **Banco de Dados:** Microsoft SQL Server
- **Conectividade:** ADO (ActiveX Data Objects 2.8)
- **Arquitetura de Dados:** Stored Procedures (`INSERT`, `SELECT`, `DELETE`)

---

## 📁 Estrutura do Repositório

```text
MiniGestorDespesas/
├── database/
│   ├── 01_create_database.sql
│   ├── 02_create_table.sql
│   └── 03_stored_procedures.sql
└── src/
    ├── MiniDespesas.vbp
    ├── frmPrincipal.frm
    └── modConexao.bas