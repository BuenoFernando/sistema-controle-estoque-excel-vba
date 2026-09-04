# 🏗️ Arquitetura do Sistema

## 📌 Visão geral

O **Sistema de Controle de Estoque** foi desenvolvido utilizando **Microsoft Excel e VBA**, combinando interface gráfica através de UserForms, tabelas estruturadas e rotinas de automação.

A solução foi projetada para organizar o cadastro de produtos e controlar as principais operações relacionadas ao estoque, como entradas, saídas, ajustes e consulta do histórico de movimentações.

---

## 🧱 Estrutura do sistema

O sistema é organizado em diferentes áreas responsáveis por funções específicas.

### Interface

A interação com o usuário é realizada principalmente através de **UserForms**, permitindo executar operações sem a necessidade de manipular diretamente as tabelas de dados.

Principais telas:

* Menu principal
* Cadastro de mercadorias
* Edição de produtos
* Consulta de estoque
* Entrada de estoque
* Saída de estoque
* Acerto de estoque
* Histórico de movimentos

---

## 🗄️ Estrutura de dados

O projeto utiliza tabelas estruturadas para armazenar e organizar as informações.

Entre as principais estruturas estão:

* `BD_Produtos`
* `BD_Categorias`
* `BD_Marcas`
* `BD_Fornecedores`
* `BD_Entradas`
* `BD_Saidas`
* `BD_Estoque`
* `BD_Movimentos`

A separação das informações em diferentes estruturas facilita a organização dos dados e a manutenção do sistema.

---

## ⚙️ Camada de programação

A lógica do sistema é implementada em **VBA (Visual Basic for Applications)**.

As rotinas VBA são responsáveis por executar tarefas como:

* Cadastro e edição de produtos
* Geração automática de códigos
* Registro de entradas
* Registro de saídas
* Ajustes de estoque
* Atualização dos saldos
* Registro das movimentações
* Consultas e filtros
* Automatização de processos

---

## 🔄 Fluxo de movimentação

As movimentações de estoque seguem uma lógica centralizada para manter o histórico das operações.

### Entrada

```text
Produto
   ↓
Registro da entrada
   ↓
Atualização do estoque
   ↓
Registro da movimentação
```

### Saída

```text
Produto
   ↓
Registro da saída
   ↓
Atualização do estoque
   ↓
Registro da movimentação
```

### Ajuste

```text
Produto
   ↓
Acerto de estoque
   ↓
Atualização do saldo
   ↓
Registro da movimentação
```

Esse modelo permite manter um histórico das operações realizadas sobre o estoque.

---

## 🔢 Geração de código de produto

O sistema possui uma rotina VBA responsável pela **geração automática do código dos produtos**, reduzindo a necessidade de criação manual dos identificadores.

---

## 📊 Controle de estoque

O sistema trabalha com informações relacionadas ao estoque atual e ao estoque disponível.

O estoque disponível considera também a quantidade reservada:

```text
Estoque disponível =
Estoque atual - Estoque reservado
```

Essa abordagem permite diferenciar a quantidade fisicamente disponível da quantidade que já está comprometida.

---

## 🎯 Objetivo técnico

Além de solucionar um problema prático de controle de estoque, o projeto foi desenvolvido como uma aplicação prática de conhecimentos relacionados a:

* Programação
* Lógica de programação
* VBA
* Estruturação de dados
* Automação
* Desenvolvimento de interfaces
* Análise de sistemas
* Organização de processos

O projeto representa uma aplicação prática de conceitos de desenvolvimento de software utilizando uma ferramenta amplamente utilizada no ambiente corporativo.

---

## 🚀 Possíveis evoluções

O sistema pode continuar evoluindo através da implementação de novos recursos, como:

* Relatórios gerenciais
* Dashboard de indicadores
* Melhorias de desempenho
* Controle de usuários e permissões
* Backup automatizado
* Integração com banco de dados
* Migração futura para uma aplicação web ou desktop
* Integração com outras ferramentas e sistemas
