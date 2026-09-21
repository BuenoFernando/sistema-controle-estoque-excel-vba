# 📦 Sistema de Controle de Estoque — Excel/VBA

Sistema de controle de estoque desenvolvido em **Microsoft Excel e VBA**, permitindo **cadastrar produtos, controlar estoque, registrar entradas e saídas, realizar ajustes e acompanhar o histórico de movimentações**.

O projeto foi desenvolvido como parte do meu portfólio profissional, com foco na aplicação prática de conhecimentos em **VBA, automação de processos, gerenciamento de dados, desenvolvimento de sistemas e análise de processos**.

---

## 📌 Sobre o Projeto

O sistema simula o funcionamento de uma solução interna para gerenciamento de estoque, permitindo centralizar informações de produtos e controlar suas movimentações.

A aplicação permite controlar o fluxo de estoque:

**Cadastro → Entrada → Estoque → Saída → Ajuste → Histórico**

Além do controle das movimentações, o sistema possui **cadastros, consultas, geração automática de códigos e rotina de backup**, proporcionando maior organização e rastreabilidade das informações.

---

## 🎯 Objetivos

O projeto foi desenvolvido com os seguintes objetivos:

* 💻 Praticar desenvolvimento de aplicações utilizando VBA.
* 📊 Trabalhar com dados estruturados dentro do Excel.
* 📦 Desenvolver um sistema para controle de produtos e estoque.
* 🔄 Automatizar processos de entrada, saída e ajuste de estoque.
* 🖥️ Desenvolver interfaces utilizando UserForms.
* 🔎 Implementar consultas e organização das informações.
* 📋 Registrar o histórico das movimentações.
* 💾 Implementar rotina de backup automático.
* ⚙️ Aplicar regras de negócio em uma aplicação prática.
* 💼 Desenvolver um projeto completo para portfólio profissional.

---

## ⚙️ Funcionalidades

### 📦 Cadastro de Produtos

Permite realizar o cadastro e gerenciamento dos produtos do estoque.

Principais informações e operações:

* Cadastro de produtos
* Edição de produtos
* Cadastro de categorias
* Geração automática de códigos
* Consulta das informações cadastradas

### 📊 Controle de Estoque

Permite acompanhar as quantidades disponíveis dos produtos e manter o controle dos saldos de estoque.

O sistema possibilita:

* Consulta do estoque atual
* Controle de estoque disponível
* Atualização automática dos saldos
* Acompanhamento das movimentações

### 📥 Entrada de Estoque

Permite registrar entradas de produtos no estoque, atualizando automaticamente as quantidades disponíveis.

### 📤 Saída de Estoque

Permite registrar saídas de produtos e atualizar os respectivos saldos.

### 🔧 Ajuste de Estoque

Permite realizar ajustes quando é necessário corrigir a quantidade registrada no sistema.

### 📋 Histórico de Movimentações

O sistema mantém o registro das movimentações realizadas, permitindo acompanhar as operações de entrada, saída e ajustes.

### 💾 Backup Automático

O sistema possui uma rotina automatizada de backup que:

* Realiza backup do arquivo do sistema.
* Controla a execução do backup diário.
* Mantém os últimos 15 backups.
* Remove automaticamente os backups mais antigos.

---

## 💻 Tecnologias Utilizadas

| Tecnologia                  | Utilização                                                    |
| --------------------------- | ------------------------------------------------------------- |
| 📗 **Microsoft Excel**      | Plataforma utilizada para desenvolvimento do sistema          |
| 💻 **VBA**                  | Programação e automação das funcionalidades                   |
| 🖥️ **UserForms**           | Desenvolvimento das interfaces gráficas                       |
| 📊 **Tabelas estruturadas** | Organização e armazenamento dos dados                         |
| ⚙️ **VBA**                  | Implementação das regras de negócio e processos automatizados |
| 🌿 **Git**                  | Controle de versão                                            |
| 🐙 **GitHub**               | Hospedagem e publicação do projeto                            |

---

## 🗄️ Estrutura de Dados

O sistema utiliza estruturas de dados dentro do próprio Excel para armazenar e organizar as informações relacionadas ao controle de estoque.

Entre as principais áreas estão:

```text
Produtos
Categorias
Fornecedores
Entradas
Saídas
Movimentos
Estoque
```

As movimentações são utilizadas para manter o controle dos saldos de estoque e registrar o histórico das operações realizadas.

O sistema aplica regras de negócio em VBA para:

* Cadastro de produtos
* Consulta de informações
* Registro de entradas
* Registro de saídas
* Ajustes de estoque
* Atualização dos saldos
* Registro das movimentações
* Controle dos backups

---

## 📁 Estrutura do Projeto

```text
sistema-controle-estoque-excel-vba/
│
├── Sistema_Controle_Estoque.xlsm
│
├── imagens/
│   ├── menu.JPG
│   ├── cadastro mercadorias.JPG
│   ├── editar produto.JPG
│   ├── consulta estoque.JPG
│   ├── entrada.JPG
│   ├── saida.JPG
│   ├── acerto de estoque.JPG
│   └── historico movimentos.JPG
│
├── README.md
└── .gitignore
```

### 🧩 Organização do Sistema

O sistema é organizado em módulos responsáveis pelas diferentes áreas da aplicação, incluindo:

**Produtos**

Responsável pelo cadastro, edição e gerenciamento dos produtos.

**Categorias**

Responsável pelo cadastro e organização das categorias.

**Entradas**

Responsável pelo registro das entradas de produtos no estoque.

**Saídas**

Responsável pelo registro das saídas e atualização dos saldos.

**Movimentos**

Responsável pelo histórico das movimentações realizadas.

**Estoque**

Responsável pela consulta e controle dos saldos disponíveis.

**Usuários**

Estrutura destinada ao gerenciamento de usuários do sistema.

**Backup**

Responsável pela rotina automatizada de cópia e retenção dos arquivos de backup.

---

## 🔄 Fluxo do Sistema

```text
                  ┌─────────────────────┐
                  │    MENU PRINCIPAL   │
                  └──────────┬──────────┘
                             │
        ┌────────────────────┼────────────────────┐
        │                    │                    │
        ▼                    ▼                    ▼
    Produtos             Estoque            Movimentos
        │                    │                    │
        ▼                    ▼                    ▼
   Cadastro             Entradas /             Histórico
   e Edição               Saídas
        │                    │
        └────────────────────┼────────────────────┘
                             │
                             ▼
                    Atualização do Estoque
                             │
                             ▼
                         Backup
```

---

## 🖼️ Demonstração

### 🏠 Menu Principal

![Menu Principal](imagens/menu.JPG)

Tela inicial utilizada para acessar as principais funcionalidades do sistema.

### 📦 Cadastro de Mercadorias

![Cadastro de Mercadorias](imagens/cadastro%20mercadorias.JPG)

Tela utilizada para cadastrar e gerenciar os produtos.

### ✏️ Editar Produto

![Editar Produto](imagens/editar%20produto.JPG)

Permite localizar e alterar informações dos produtos cadastrados.

### 📊 Consulta de Estoque

![Consulta de Estoque](imagens/consulta%20estoque.JPG)

Tela utilizada para consultar as quantidades disponíveis em estoque.

### 📥 Entrada de Estoque

![Entrada de Estoque](imagens/entrada.JPG)

Permite registrar a entrada de produtos e atualizar o estoque.

### 📤 Saída de Estoque

![Saída de Estoque](imagens/saida.JPG)

Permite registrar a saída de produtos e atualizar os respectivos saldos.

### 🔧 Ajuste de Estoque

![Ajuste de Estoque](imagens/acerto%20de%20estoque.JPG)

Permite realizar ajustes nas quantidades registradas no sistema.

### 📋 Histórico de Movimentações

![Histórico de Movimentações](imagens/historico%20movimentos.JPG)

Permite consultar o histórico das movimentações realizadas.

---

## 🚀 Como Executar o Projeto

### 1️⃣ Pré-requisito

É necessário possuir o **Microsoft Excel para Windows** com suporte à execução de macros VBA.

### 2️⃣ Baixar o Projeto

Faça o download do arquivo `.xlsm` disponível neste repositório.

### 3️⃣ Abrir o Sistema

Abra o arquivo utilizando o **Microsoft Excel para Windows**.

### 4️⃣ Habilitar as Macros

Caso o Excel solicite autorização, habilite as **macros/conteúdo** para permitir a execução do sistema.

### 5️⃣ Utilizar o Sistema

Após a abertura, utilize o **Menu Principal** para acessar as funcionalidades disponíveis.

> **Observação:** o sistema foi desenvolvido em VBA e requer o Microsoft Excel para Windows com suporte a macros habilitado.

---

## 📑 Controle e Backup

O sistema possui uma rotina de backup automatizado para aumentar a segurança das informações.

O processo permite:

* 💾 Realizar backup automático.
* 📅 Controlar a execução diária do backup.
* 🗂️ Manter os últimos 15 arquivos.
* 🧹 Excluir automaticamente os backups mais antigos.

Essa funcionalidade busca reduzir o risco de perda das informações armazenadas no sistema.

---

## 🧠 Conceitos Demonstrados

Este projeto demonstra conhecimentos práticos em:

* 💻 VBA
* 📊 Microsoft Excel
* 🖥️ Desenvolvimento de interfaces com UserForms
* 🧩 Programação orientada a eventos
* 📦 Regras de negócio
* 📋 CRUD e gerenciamento de informações
* 🔄 Controle de movimentações
* 📈 Controle de estoque
* 💾 Automação de backups
* 🔎 Consultas e organização de dados
* ⚙️ Automação de processos
* 📁 Organização de sistemas
* 🌿 Controle de versão com Git
* 🐙 Publicação de projetos no GitHub

---

## 🔮 Possíveis Evoluções

Como projeto de portfólio, o sistema foi desenvolvido de forma **enxuta e funcional**, mantendo o foco nas principais operações de controle de estoque.

Algumas possibilidades de evolução seriam:

* 🔐 Implementação completa de login e controle de permissões.
* 📊 Criação de dashboards gerenciais.
* 📑 Ampliação dos relatórios.
* 🔔 Implementação de alertas de estoque mínimo.
* 📦 Melhorias no controle de fornecedores.
* 📈 Indicadores adicionais de movimentação.
* 🗄️ Migração dos dados para um banco de dados externo.
* 🔌 Integração com outras aplicações ou sistemas.

**Essas funcionalidades não fazem parte da versão atual.**

---

## 💼 Objetivo Profissional

Este projeto faz parte do meu portfólio de transição para a área de **Tecnologia da Informação**, demonstrando a aplicação prática de conhecimentos em **programação, automação, análise de sistemas, gestão de dados e desenvolvimento de soluções para processos empresariais**.

O objetivo é desenvolver soluções **simples, organizadas e funcionais** para problemas encontrados em ambientes corporativos.

---

## 👨‍💻 Autor

**Fernando Bueno**

**Engenheiro de Computação | Analista de Sistemas | Analista de TI**

### 🔗 Contato

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Fernando%20Bueno-blue?style=for-the-badge\&logo=linkedin)](https://www.linkedin.com/in/fernando-cesar-bueno/)

[![GitHub](https://img.shields.io/badge/GitHub-BuenoFernando-black?style=for-the-badge\&logo=github)](https://github.com/BuenoFernando)
