# 💰 Mini Gestor de Despesas

Aplicação desktop em **Visual Basic 6 e SQL Server** para cadastrar, consultar, editar e excluir despesas, com totalização dos valores registrados.

## 🎯 O que o sistema resolve

Centraliza descrição, data e valor de cada gasto em uma lista. Permite corrigir lançamentos sem excluir e cadastrar novamente, consultar o total e remover um registro após confirmação.

## ✨ Funcionalidades

- Cadastro com descrição, data e valor positivo, limitado a duas casas decimais.
- Edição do registro selecionado e cancelamento da edição.
- Exclusão definitiva com confirmação, mostrando descrição e valor.
- Filtro de mês e ano, iniciado no mês atual, com total apenas das despesas do período. Meses sem lançamentos mostram total zero, sem apagar o histórico.
- Interface em azul-escuro e verde, fonte Segoe UI e tabela redimensionável.
- Duplo clique para editar; **F2** limpa os campos/cancela a edição; **F5** atualiza a lista.

## 🧩 Arquitetura e decisões

O formulário chama stored procedures por `ADODB.Command`, com parâmetros tipados. Valores monetários usam `Currency` no VB6 e `DECIMAL(10,2)` no SQL Server; a soma em memória usa Decimal para comportar o total de muitos registros.

Edição e exclusão comparam os valores originais com os do banco. Se outro processo modificar ou remover o registro, a operação é recusada e o usuário deve atualizar a lista. Essa comparação não substitui um histórico de auditoria ou controle por `rowversion`.

As quatro procedures novas possuem prefixo `usp_Despesas`. O script de evolução mantém a tabela, seus dados e as procedures anteriores. A tela atual utiliza as novas procedures e permite consultar cada mês separadamente.

## 📂 Organização

```text
src/                         Projeto, formulário e conexão VB6
database/01_create_database.sql
database/02_create_table.sql
database/03_stored_procedures.sql
database/04_modernizacao_crud.sql
database/05_filtro_mensal.sql
assets/mini-despesas.ico      Ícone do atalho
bin/MiniDespesas.exe          Compilação local, fora do Git
```

## 🚀 Instalação e execução

1. Disponibilize a instância SQL Server `.\SQLEXPRESS` com autenticação Windows.
2. Para uma instalação nova, execute os scripts SQL na ordem **01, 02, 03, 04, 05**. Em uma instalação já modernizada (script 04 aplicado), aplique somente **05** para adicionar o filtro mensal.
3. Garanta que a conta Windows usada pelo aplicativo tenha acesso ao banco `DB_MiniDespesas` e às operações necessárias.
4. Compile `src/MiniDespesas.vbp` com VB6 para `bin/MiniDespesas.exe`, ou utilize um executável já compilado.
5. Abra o executável ou o atalho **Mini Gestor de Despesas**, na raiz da instalação local.

O aplicativo utiliza ADO 2.8, `SQLOLEDB`, runtime VB6 e o controle **MSFLXGRD.OCX**. Uma instalação em outro computador precisa disponibilizar essas dependências de 32 bits. O editor VB6 é necessário para compilar/editar, não para usar o executável. O SSMS é uma ferramenta administrativa opcional.

## 🎨 Ícone e atalho

O ícone original está em `assets/mini-despesas.ico`. O atalho local aponta para `bin/MiniDespesas.exe`, utiliza `bin` como pasta de trabalho e o arquivo `.ico` como ícone. Se mover a pasta ou trocar de computador, recrie o atalho com os novos caminhos. Atalhos `.lnk` não são publicados porque armazenam caminhos locais.

## 📌 Escopo e validação

- A versão foi compilada com sucesso no ambiente local.
- A migração foi aplicada sem alterar os dois registros existentes.
- Não foram executados testes funcionais desta atualização; a interface em execução e outras escalas de Windows ainda precisam ser verificadas pelo usuário.
- O layout tem tamanho mínimo; não oferece adaptação específica para telas pequenas ou múltiplos monitores com escalas diferentes.
- Exclusões são definitivas. Não há lixeira, histórico de alterações, autenticação própria, exportação ou categorias nesta versão.
- A leitura de valores segue o separador decimal configurado no Windows. A data é informada em `dd/mm/aaaa`.
- Fontes VB6 estão em Windows-1252. Scripts SQL e documentação estão em UTF-8.

O repositório contém os fontes e scripts de estrutura, sem executáveis ou dados pessoais.

## 📅 Consulta mensal

Ao abrir, a tela seleciona o mês e ano atuais. Escolha outro mês, digite o ano e clique em **Filtrar mês**. O total corresponde à lista carregada; F5 atualiza o período aplicado. A troca de período fica bloqueada durante a edição: salve ou cancele primeiro. Se salvar uma despesa em outro mês, um aviso informa onde consultá-la.

Não há exclusão ou fechamento automático na virada do mês. Com a janela aberta, o período permanece selecionado; ao reabrir, o mês atual é selecionado novamente.
