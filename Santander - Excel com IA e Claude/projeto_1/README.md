# 📊 Simulador de Investimentos em FIIs & Alocação por Perfil de Risco
### 📊 Real Estate Investment Simulator (FIIs) & Risk Profile Allocation

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

O projeto consiste em uma planilha interativa desenvolvida em Microsoft Excel (`projeto_1.xlsx`) para planejamento financeiro e investimentos em Fundos Imobiliários (FIIs). A ferramenta automatiza a projeção de acúmulo patrimonial via juros compostos a longo prazo, calcula a renda passiva esperada em dividendos e gera uma recomendação personalizada de distribuição da carteira entre diferentes segmentos de FIIs com base no perfil de risco do investidor.

---

### ⚙️ Funcionalidades e Módulos da Planilha

#### Aba `APP` (Dashboard Principal)
* **Planejamento Orçamentário**: Registro de salário e taxa de rendimento da carteira, gerando uma sugestão automatizada de aporte mensal equivalente a 30% da renda.
* **Simulador de Investimento Mensal**: Parâmetros de entrada editáveis para valor do aporte, prazo em anos e taxa mensal esperada.
* **Projeção Multicenário**: Tabela comparativa automática projetando a evolução do patrimônio e dos dividendos mensais nos horizontes de **2, 5, 10, 20 e 30 anos**.
* **Alocação Inteligente por Perfil**: Seleção via menu suspenso do perfil do investidor (*Conservador*, *Moderado* ou *Agressivo*).
* **Distribuição por Categoria**: Cálculo dos valores em reais a serem aplicados nas classes de FIIs (*Papel*, *Tijolo*, *Híbridos*, *FOFs*, *Desenvolvimento* e *Hotelarias*).
* **Visualização Gráfica**: Gráfico de colunas integrado representando o percentual recomendado por classe de ativo.

#### Aba `Planilha2` (Matriz de Referência)
* Matriz de apoio contendo a lógica de distribuição percentual por categoria de fundo imobiliário associada a cada perfil de risco.

---

### 🧮 Modelagem de Dados e Fórmulas

A inteligência da planilha fundamenta-se na utilização de intervalos nomeados, matemática financeira e busca dinâmica de dados:

#### 1. Intervalos Nomeados (Defined Names)
Para legibilidade e manutenção das fórmulas, a planilha utiliza os seguintes nomeados na aba `APP`:
* `aporte`: `$D$17`
* `patrimonio`: `$D$20`
* `qtd_anos`: `$D$18`
* `rendimento_carteira`: `$D$13`
* `sugestao`: `$D$14`
* `taxa_mensal`: `$D$19`

#### 2. Projeção Patrimonial (Juros Compostos)
Utilização da função de Valor Futuro (`VF` / `FV`) para cálculo do saldo acumulado:

`=VF(taxa_mensal; qtd_anos * 12; aporte * -1)`

#### 3. Matriz de Busca por Chave Concatenada
Para obter os percentuais recomendados da `Planilha2`, é utilizada uma chave de busca composta (`Perfil-Tipo de FII`):

`=PROCV($C$30&"-"&B56; Planilha2!A3:D20; 4; FALSO)`

---

### 📁 Estrutura do Repositório

O arquivo do projeto encontra-se organizado no repositório no seguinte caminho:

desafios_dio/
└── Santander - Excel com IA e Claude/
└── projeto_1/
├── projeto_1.xlsx
└── README.md


---

### 🚀 Como Executar o Projeto

1. **Clonar o Repositório**:
git clone https://github.com/ThiagoSarnaglia1986/desafios_dio.git

2. **Navegar até a Pasta**:
cd "desafios_dio/Santander - Excel com IA e Claude/projeto_1"

3. **Abrir a Planilha**:
* Abra o arquivo `projeto_1.xlsx` no Microsoft Excel (2016 ou superior).
* Altere as células editáveis na aba `APP` (Salário, Aporte Mensal, Prazo e Perfil) para visualizar a atualização automática dos cenários e gráficos.

---

### 🛠️ Tecnologias Utilizadas

* **Microsoft Excel**: Modelagem financeira, funções de juros compostos (`VF`), busca dinâmica (`PROCV`), intervalos nomeados e gráficos.
* **Git & GitHub**: Versionamento de código e documentação.
* **Markdown**: Estruturação técnica do projeto.

---

### 👤 Autor

Desenvolvido por **Thiago Queiroz Sarnaglia**  
*Data Analyst Portfolio*

* 🐙 **GitHub**: [@ThiagoSarnaglia1986](https://github.com/ThiagoSarnaglia1986)
* 📁 **Repositório**: [desafios_dio](https://github.com/ThiagoSarnaglia1986/desafios_dio/tree/main/Santander%20-%20Excel%20com%20IA%20e%20Claude/projeto_1)

---

## 🇺🇸 English

### 📌 Table of Contents
- [Overview](#-overview)
- [Spreadsheet Modules and Features](#-spreadsheet-modules-and-features)
- [Data Modeling and Formulas](#-data-modeling-and-formulas)
- [Repository Structure](#-repository-structure)
- [How to Run the Project](#-how-to-run-the-project)
- [Technologies Used](#-technologies-used)
- [Author](#-author-1)

---

### 📝 Overview

This project consists of an interactive Microsoft Excel spreadsheet (`projeto_1.xlsx`) designed for personal financial planning and investment simulation in Real Estate Investment Funds (FIIs). The tool automates long-term wealth accumulation projections using compound interest, calculates expected passive dividend income, and generates a personalized portfolio allocation recommendation across different FII sectors based on the investor's risk profile.

---

### ⚙️ Spreadsheet Modules and Features

#### `APP` Sheet (Main Dashboard)
* **Budget Planning**: Input salary and portfolio yield rate to automatically receive an investment contribution suggestion equal to 30% of income.
* **Monthly Investment Simulator**: Editable parameters for contribution amount, timeframe in years, and expected monthly yield.
* **Multi-Scenario Projection**: Automatic comparative table projecting wealth accumulation and monthly dividend generation over **2, 5, 10, 20, and 30-year** horizons.
* **Profile-Based Smart Allocation**: Dropdown menu to select investor risk profile (*Conservative*, *Moderate*, or *Aggressive*).
* **Category Breakdown**: Automatic calculation of monetary values to be invested across FII categories (*Paper*, *Brick*, *Hybrids*, *FoFs*, *Development*, and *Hotels*).
* **Data Visualization**: Integrated column chart displaying the recommended portfolio percentage per asset class.

#### `Planilha2` Sheet (Reference Matrix)
* Background lookup matrix containing percentage allocation rules for each FII sector mapped to investor risk profiles.

---

### 🧮 Data Modeling and Formulas

The core logic relies on named ranges, financial mathematics, and dynamic lookup functions:

#### 1. Named Ranges
To improve formula readability and maintainability, named ranges are used in the `APP` sheet:
* `aporte`: `$D$17`
* `patrimonio`: `$D$20`
* `qtd_anos`: `$D$18`
* `rendimento_carteira`: `$D$13`
* `sugestao`: `$D$14`
* `taxa_mensal`: `$D$19`

#### 2. Wealth Accumulation (Compound Interest)
Uses the Future Value (`FV` / `VF`) function to project accumulated capital:

`=FV(taxa_mensal; qtd_anos * 12; aporte * -1)`

#### 3. Dynamic Lookup via Concatenated Key
To retrieve allocation percentages from `Planilha2`, a composite key is used (`Profile-FII Type`):

`=VLOOKUP($C$30&"-"&B56; Planilha2!A3:D20; 4; FALSE)`

---

### 📁 Repository Structure

The project files are structured within the repository at the following path:

desafios_dio/
└── Santander - Excel com IA e Claude/
└── projeto_1/
├── projeto_1.xlsx
└── README.md


---

### 🚀 How to Run the Project

1. **Clone the Repository**:
git clone https://github.com/ThiagoSarnaglia1986/desafios_dio.git

2. **Navigate to the Directory**:
cd "desafios_dio/Santander - Excel com IA e Claude/projeto_1"

3. **Open the Spreadsheet**:
* Open `projeto_1.xlsx` in Microsoft Excel (2016 or newer).
* Adjust the editable cells in the `APP` tab (Salary, Monthly Contribution, Timeframe, Profile) to view dynamic scenario updates and chart visualizations.

---

### 🛠️ Technologies Used

* **Microsoft Excel**: Financial modeling, compound interest functions (`FV`), dynamic lookup (`VLOOKUP`), named ranges, and data charts.
* **Git & GitHub**: Version control and project documentation.
* **Markdown**: Technical documentation formatting.

---

### 👤 Author

Developed by **Thiago Queiroz Sarnaglia**  
*Data Analyst Portfolio*

* 🐙 **GitHub**: [@ThiagoSarnaglia1986](https://github.com/ThiagoSarnaglia1986)
* 📁 **Repository**: [desafios_dio](https://github.com/ThiagoSarnaglia1986/desafios_dio/tree/main/Santander%20-%20Excel%20com%20IA%20e%20Claude/projeto_1)
