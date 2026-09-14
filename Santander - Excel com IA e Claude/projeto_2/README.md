Markdown
# 📑 Agregador de Dados para Imposto de Renda (IRPF)
### 📑 Income Tax Data Aggregator (IRPF)

[![DIO](https://img.shields.io/badge/DIO-Bootcamp-openrss?style=for-the-badge&logo=dio&color=7308a3)](https://www.dio.me/)
[![Santander](https://img.shields.io/badge/Santander-Excel%20com%20IA%20e%20Claude-ec0000?style=for-the-badge&logo=santander&logoColor=white)](https://www.dio.me/)
[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)

---

🌐 **Languages / Idiomas**: [Português](#-português) | [English](#-english)

---

## 🇧🇷 Português

### 📌 Sumário
- [Visão Geral](#-visão-geral)
- [Funcionalidades e Módulos da Planilha](#-funcionalidades-e-módulos-da-planilha)
- [Modelagem de Dados e Fórmulas](#-modelagem-de-dados-e-fórmulas)
- [Estrutura do Repositório](#-estrutura-do-repositório)
- [Como Executar o Projeto](#-como-executar-o-projeto)
- [Tecnologias Utilizadas](#-tecnologias-utilizadas)
- [Autor](#-autor)

---

### 📝 Visão Geral

O projeto consiste em uma ferramenta de centralização e organização de dados fiscais desenvolvida em Microsoft Excel (`Projeto_2.xlsx`) para apoiar o preenchimento da Declaração de Imposto de Renda Pessoa Física (IRPF). A solução automatiza a consolidação de saldos bancários, valida informações de instituições financeiras e organiza os históricos de receitas e holerites em um ambiente estruturado.

---

### ⚙️ Funcionalidades e Módulos da Planilha

#### Aba `TITULAR` (Dados Cadastrais)
* **Cadastro Fiscal de Pessoa Física**: Registro padronizado de CPF, Título de Eleitor, endereço completo, contatos e dados do cônjuge.
* **Validação de Status Fiscal**: Flags de verificação para alteração de dados cadastrais, dependência de cônjuge e residência no exterior.

#### Aba `INFORMES` (Rendimentos Bancários)
* **Consolidação Bancária**: Registro de posições financeiras por instituição financeira com referência a anexos de comprovantes em PDF.
* **Soma Automatizada de Saldos**: Cálculo dinâmico do patrimônio financeiro total informado.

#### Aba `NOTAS` (Extratos e Holerites)
* **Registro de Entradas**: Controle cronológico de receitas mensais categorizadas (ex: holerites e rendimentos do trabalho).

#### Aba `BANCOS` (Matriz de Referência)
* **Tabela de Apoio de Instituições Financeiras**: Matriz contendo o código de compensação e nome oficial dos bancos para validação e padronização.

---

### 🧮 Modelagem de Dados e Fórmulas

A inteligência do modelo baseia-se na padronização de registros e fórmulas de consolidação:

#### 1. Consolidação de Posição Patrimonial Total
Soma automática dos valores informados nas contas bancárias cadastradas:

`=SOMA(D11; D17; D22)`

#### 2. Padronização de Lista de Bancos
Utilização do intervalo da aba `BANCOS` (`A2:A51`) para validação de dados e consistência no preenchimento dos informes.

---

### 📁 Estrutura do Repositório

```text
desafios_dio/
└── Santander - Excel com IA e Claude/
    └── projeto_2/
        ├── Projeto_2.xlsx
        └── README.md 
```
### 🚀 Como Executar o Projeto
Clonar o Repositório:

Bash
git clone [https://github.com/ThiagoSarnaglia1986/desafios_dio.git](https://github.com/ThiagoSarnaglia1986/desafios_dio.git)

Navegar até a Pasta:

Bash
cd "desafios_dio/Santander - Excel com IA e Claude/projeto_2"
Abrir a Planilha:

Abra o arquivo Projeto_2.xlsx no Microsoft Excel (2016 ou superior).

Preencha os campos nas abas TITULAR, INFORMES e NOTAS para consolidar as informações fiscais.

###🛠️ Tecnologias Utilizadas
Microsoft Excel: Modelagem de dados fiscais, validação de entradas, fórmulas de soma (SOMA/SUM) e estrutura relacional de apoio.

Git & GitHub: Versionamento de código e documentação.

Markdown: Estruturação técnica do projeto.

###👤 Autor
Desenvolvido por Thiago Queiroz Sarnaglia

Data Analyst Portfolio

🐙 GitHub: @ThiagoSarnaglia1986

📁 Repositório: desafios_dio

🇺🇸 English
###📌 Table of Contents
Overview

Spreadsheet Modules and Features

Data Modeling and Formulas

Repository Structure

How to Run the Project

Technologies Used

Author

###📝 Overview
This project consists of an Income Tax Data Aggregator developed in Microsoft Excel (Projeto_2.xlsx) to streamline and structure financial data required for Individual Income Tax Returns (IRPF). The tool automates bank balance consolidation, validates financial institution data, and organizes monthly income streams and paystub records in a structured repository.

###⚙️ Spreadsheet Modules and Features
####TITULAR Sheet (Taxpayer Details)
* Taxpayer Information: Standardized record of CPF, voter ID, full address, contact details, and spouse information.

* Tax Status Flags: Boolean fields for address change tracking, spouse dependency, and foreign residence status.

###INFORMES Sheet (Bank Income Statements)
* Bank Position Consolidation: Tracking financial balances per banking institution alongside document attachment references (PDFs).

* Automated Total Sum: Dynamic calculation of total reported financial assets.

####NOTAS Sheet (Income & Paystub Records)
* Revenue Tracking: Chronological log of categorized monthly revenues (e.g., salary, paystubs, and secondary income).

####BANCOS Sheet (Reference Table)
* Financial Institutions Matrix: Support table containing clearing codes and official bank names for data validation and standardized lookup.

###🧮 Data Modeling and Formulas
The core logic relies on standardized records and aggregation formulas:

####1. Total Wealth Consolidation
Automated summation of bank balances across registered financial accounts:

=SUM(D11; D17; D22)

####2. Bank List Standardizing
Data validation referencing the BANCOS range (A2:A51) ensuring input consistency across statements.

### 📁 Repository Structure
```Plaintext
desafios_dio/
└── Santander - Excel com IA e Claude/
    └── projeto_2/
        ├── Projeto_2.xlsx
        └── README.md
```
🚀 How to Run the Project
Clone the Repository:

Bash
git clone [https://github.com/ThiagoSarnaglia1986/desafios_dio.git](https://github.com/ThiagoSarnaglia1986/desafios_dio.git)
Navigate to the Directory:

Bash
cd "desafios_dio/Santander - Excel com IA e Claude/projeto_2"
Open the Spreadsheet:

Open Projeto_2.xlsx in Microsoft Excel (2016 or newer).

Fill in the fields under TITULAR, INFORMES, and NOTAS to consolidate your tax information.

###🛠️ Technologies Used
Microsoft Excel: Financial data modeling, data validation, summation functions (SUM/SOMA), and relational reference structures.

Git & GitHub: Version control and project documentation.

Markdown: Technical documentation formatting.

###👤 Author
Developed by Thiago Queiroz Sarnaglia

Data Analyst Portfolio

🐙 GitHub: @ThiagoSarnaglia1986

📁 Repository: desafios_dio
