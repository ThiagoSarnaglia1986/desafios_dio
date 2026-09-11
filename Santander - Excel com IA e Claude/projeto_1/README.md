# 📊 Simulador de Investimentos em FIIs & Alocação por Perfil de Risco
### 📊 Real Estate Investment Simulator (FIIs) & Risk Profile Allocation

[![DIO](https://img.shields.io/badge/DIO-Bootcamp-openrss?style=for-the-badge&logo=dio&color=7308a3)](https://www.dio.me/)
[![Santander](https://img.shields.io/badge/Santander-Excel%20com%20IA%20e%20Claude-ec0000?style=for-the-badge&logo=santander&logoColor=white)](https://www.dio.me/)
[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)

---

🌐 **Idiomas / Languages**: [Português](#-português) | [English](#-english)

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
$$\text{Patrimônio} = \text{VF}(\text{taxa\_mensal}; \text{qtd\_anos} \times 12; \text{aporte} \times -1)$$

#### 3. Matriz de Busca por Chave Concatenada
Para obter os percentuais recomendados da `Planilha2`, é utilizada uma chave de busca composta (`Perfil-Tipo de FII`):
```excel
=PROCV($C$30&"-"&B56; Planilha2!A3:D20; 4; FALSO)
