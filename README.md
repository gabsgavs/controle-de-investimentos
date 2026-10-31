# 📊 Ferramenta de Controle de Investimentos e Projeção de FIIs

Projeto desenvolvido como parte do portfólio de análise financeira e modelagem em Excel. Esta ferramenta interativa foi projetada para auxiliar investidores a planejar aportes, projetar o crescimento patrimonial de longo prazo e distribuir os investimentos em Fundos Imobiliários (FIIs) de acordo com seu perfil de risco.

---

## 🔍 1. As 5 Perguntas que a Ferramenta Responde

A planilha foi estruturada para responder de forma automatizada a cinco perguntas essenciais para o planejamento financeiro:

1. **Quanto investir por mês?**
   * **Onde aparece:** Na célula `D17` (Intervalo nomeado: `aporte`), onde o usuário define o valor destinado mensalmente aos investimentos.
2. **Por quantos anos?**
   * **Onde aparece:** Na célula `D18` (Intervalo nomeado: `qtd_anos`), que define o horizonte temporal do investimento.
3. **Qual a taxa de rendimento mensal?**
   * **Onde aparece:** Na célula `D19` (Intervalo nomeado: `taxa_mensal`), que registra a rentabilidade esperada ao mês.
4. **Quanto de patrimônio vai acumular?**
   * **Onde aparece:** Na célula `D20` (Intervalo nomeado: `patrimonio`), calculada automaticamente com base nos aportes, prazo e taxa.
5. **Quanto vai receber de dividendos por mês?**
   * **Onde aparece:** Na célula `D21`, que multiplica o patrimônio acumulado pela taxa de rendimento/dividend yield da carteira.

---

## 🧮 2. Como o VF e o PROCV Entram nos Cálculos

* **Função VF (Valor Futuro / `FV`):**
  Utilizada para calcular o montante acumulado ao longo do tempo com juros compostos e aportes periódicos. 
  * *Fórmula utilizada (Célula `D20` e tabela de cenários):* 
    `=FV(taxa_mensal, qtd_anos*12, aporte*-1)`
  * *Explicação:* A função projeta o valor futuro considerando a taxa mensal, o número total de meses (`anos * 12`) e o aporte mensal. O sinal negativo no aporte (`aporte*-1`) indica a saída de caixa (investimento).

* **Função PROCV (`VLOOKUP`):**
  Utilizada na aba principal para buscar dinamicamente os percentuais de alocação de acordo com o perfil de risco selecionado pelo usuário (`Conservador`, `Moderado` ou `Agressivo`).
  * *Fórmula utilizada:* 
    `=VLOOKUP($C$32&"-"&B37, Planilha2!$A:$D, 4, FALSE)`
  * *Explicação:* A fórmula concatena o perfil selecionado com o tipo de FII (ex: `Moderado-PAPEL`) formando uma chave única que busca o percentual correspondente na base de dados auxiliar, multiplicando-o pelo valor total do aporte.

---

## 🏷️ 3. Intervalos Nomeados Criados

Para facilitar a leitura das fórmulas e estruturação da planilha, foram definidos os seguintes intervalos nomeados no Gerenciador de Nomes do Excel:

* `aporte` $\rightarrow$ `input_data!$D$17`
* `patrimonio` $\rightarrow$ `input_data!$D$20`
* `qtd_anos` $\rightarrow$ `input_data!$D$18`
* `redimento_carteira` $\rightarrow$ `input_data!$D$13`
* `salario` $\rightarrow$ `input_data!$D$12`
* `sugestao_investimento` $\rightarrow$ `input_data!$D$14`
* `taxa_mensal` $\rightarrow$ `input_data!$D$19`

---

## 📈 4. Percentuais de Cada Perfil e Origem

Os percentuais de alocação vieram de premissas clássicas de diversificação no mercado de Fundos Imobiliários, equilibrando a renda de papéis com a valorização de ativos físicos:

| Tipo de FII | Conservador | Moderado | Agressivo |
| :--- | :---: | :---: | :---: |
| **Papel** | 30% | 32% | 50% |
| **Tijolo** | 50% | 35% | 10% |
| **Híbridos** | 10% | 8% (ou 5%) | 5% |
| **FOFs** | 10% | 5% | 5% |
| **Desenvolvimento** | 0% | 10% | 20% |
| **Hotelarias** | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

---

## 🔄 5. O Que Mudou em Relação à Ferramenta do Expert

Em comparação com a ferramenta padrão fornecida pelo expert, implementei as seguintes melhorias estruturais e visuais:
1. **Estrutura de Abas e Nomenclatura Mais Limpa:** Renomeação e organização das abas de parâmetros e dados (`input_data` / `Sheet2`), tornando a navegação mais intuitiva e com padrão profissional de mercado.
2. **Automação da Chave de Busca (PROCV):** Na tabela de parâmetros de FIIs, adicionei a concatenação direta via fórmula nas chaves de busca (`=B3&"-"&C3`), garantindo robustez caso a tabela de perfis seja ordenada ou expandida.
3. **Refinamento de Alocação Tática:** Ajuste fino nos percentuais de FIIs Híbridos no perfil Moderado (otimizando a distribuição de risco) e padronização visual em letras maiúsculas para maior legibilidade.
4. **Projeção de Cenários Dinâmicos Expandida:** Vinculação robusta da tabela de cenários de longo prazo (2 a 30 anos) diretamente aos inputs de aporte e taxa da planilha.

---

## 📸 Simulações da Ferramenta

*(Insira aqui prints da planilha preenchida comparando dois perfis diferentes, por exemplo, o perfil Conservador e o perfil Agressivo com o mesmo valor de aporte, demonstrando a mudança na divisão da carteira).*

---

### 💡 Checklist de Verificação Antes de Enviar:
- [x] Planilha aberta sem erros de fórmula.
- [x] As 5 perguntas respondidas e mapeadas.
- [x] Alterar o perfil altera dinamicamente os aportes e soma exatamente 100%.
- [x] Salário e valores de exemplo utilizados.
- [x] Repositório público e com nome limpo (`controle-de-investimentos`).
