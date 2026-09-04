# 🏗️ Arquitetura do Sistema

O sistema foi desenvolvido em **Microsoft Excel com VBA**, utilizando UserForms para a interface e tabelas estruturadas para organização e armazenamento dos dados.

A estrutura foi organizada de forma a separar as principais responsabilidades do sistema, facilitando sua manutenção e evolução.

---

## 🖥️ Camada de interface

A interação com o usuário é realizada principalmente através de **UserForms** desenvolvidos em VBA.

Entre as principais interfaces estão:

* Menu principal
* Cadastro de produtos
* Cadastro de categorias
* Entrada de estoque
* Saída de estoque
* Ajuste de estoque
* Consulta de estoque
* Histórico de movimentações

---

## ⚙️ Camada de regras e processamento

A lógica do sistema é executada através de procedimentos, funções e módulos VBA.

Essa camada é responsável por:

* Validação dos dados;
* Geração automática de códigos;
* Registro de entradas;
* Registro de saídas;
* Ajustes de estoque;
* Atualização dos saldos;
* Registro do histórico de movimentações;
* Controle de estoque disponível;
* Automação de tarefas;
* Rotinas de backup.

---

## 🗃️ Camada de dados

As informações são organizadas em tabelas estruturadas dentro do arquivo Excel.

Entre as principais estruturas estão:

* `BD_Produtos`
* `BD_Categorias`
* `BD_Entradas`
* `BD_Saidas`
* `BD_Estoque`
* `BD_Movimentos`

Essas estruturas armazenam os dados utilizados pelas diferentes funcionalidades do sistema.

---

## 🔄 Fluxo de movimentação

O controle de estoque considera diferentes tipos de movimentação:

**Entrada**

Produto → Registro da entrada → Atualização do estoque → Registro da movimentação

**Saída**

Produto → Registro da saída → Atualização do estoque → Registro da movimentação

**Ajuste**

Produto → Ajuste de quantidade → Atualização do estoque → Registro da movimentação

O histórico permite acompanhar as operações realizadas no sistema.

---

## 💾 Backup automático

O sistema possui uma rotina de backup automático executada durante o encerramento.

Antes de criar um novo backup, o sistema verifica se já existe um backup realizado na mesma data.

A rotina mantém no máximo **15 backups**, removendo automaticamente os arquivos mais antigos quando o limite é atingido.

---

## 🔐 Controle e organização

O projeto também possui estruturas relacionadas a usuários e permissões, permitindo organizar o acesso às funcionalidades do sistema.

---

## 🧩 Tecnologias

* Microsoft Excel
* VBA (Visual Basic for Applications)
* UserForms
* Tabelas estruturadas
* Automação de processos

---

## 🎯 Objetivo da arquitetura

A organização do projeto busca facilitar:

* Manutenção do sistema;
* Identificação das responsabilidades de cada componente;
* Evolução das funcionalidades;
* Organização dos dados;
* Redução de processos manuais;
* Rastreabilidade das movimentações.
