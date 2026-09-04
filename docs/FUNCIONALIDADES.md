# ⚙️ Funcionalidades do Sistema

## 📦 Cadastro de produtos

O sistema permite realizar o cadastro e gerenciamento dos produtos armazenados no estoque.

Entre as informações utilizadas estão:

* Código do produto
* Descrição
* Categoria
* Quantidade em estoque
* Estoque disponível
* Informações relacionadas ao produto

O código do produto é gerado automaticamente pelo sistema.

---

## 📊 Consulta de estoque

A consulta de estoque permite visualizar as informações atuais dos produtos e acompanhar suas quantidades.

---

## 📥 Entrada de estoque

A funcionalidade de entrada permite registrar o recebimento de produtos.

Durante o processo são registradas informações relacionadas à operação e o estoque é atualizado de acordo com a quantidade movimentada.

A movimentação também pode ser registrada no histórico do sistema.

---

## 📤 Saída de estoque

A funcionalidade de saída permite registrar a retirada de produtos do estoque.

O processo considera a quantidade disponível e atualiza as informações de estoque após a operação.

---

## 🔧 Acerto de estoque

O acerto de estoque permite corrigir diferenças identificadas entre a quantidade registrada no sistema e a quantidade física existente.

Essa funcionalidade é importante para manter a consistência das informações de estoque.

---

## 📋 Histórico de movimentações

O sistema mantém um histórico das movimentações realizadas.

As operações podem incluir:

* Entrada
* Saída
* Ajuste

O histórico permite acompanhar as alterações realizadas no estoque e fornece maior rastreabilidade das operações.

---

## 🔢 Geração automática de códigos

O sistema possui uma rotina VBA para geração automática dos códigos dos produtos.

Essa funcionalidade reduz a necessidade de controle manual dos identificadores e ajuda a evitar duplicidade de códigos.

---

## 🖥️ UserForms

A interface do sistema utiliza **UserForms do VBA** para facilitar a interação do usuário.

Os formulários permitem realizar operações como:

* Cadastro
* Consulta
* Entrada
* Saída
* Ajuste
* Edição de produtos

A utilização de formulários proporciona uma experiência mais próxima de uma aplicação tradicional, mesmo utilizando o Excel como plataforma.

---

## 🔄 Automação com VBA

O VBA é utilizado para automatizar diferentes processos do sistema.

Entre as automações implementadas estão:

* Manipulação de dados
* Atualização de estoque
* Registro de movimentações
* Geração de códigos
* Pesquisas e filtros
* Atualização de informações na interface
* Validação de operações

---

## 📈 Organização e manutenção dos dados

As informações são organizadas em diferentes tabelas e estruturas, separando dados de cadastro, estoque e movimentações.

Essa organização facilita:

* Consulta dos dados
* Manutenção do sistema
* Desenvolvimento de novas funcionalidades
* Identificação de responsabilidades de cada estrutura

---

## 🎯 Objetivo das funcionalidades

As funcionalidades foram desenvolvidas buscando transformar processos que normalmente seriam realizados manualmente em processos automatizados.

O projeto demonstra a aplicação prática de conceitos de:

**Programação + Banco de Dados Estruturado + Interface + Automação + Regras de Negócio**

utilizando Excel e VBA.
