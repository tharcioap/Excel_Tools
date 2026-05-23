# Top Ofenders Hitlist Automation

Este repositório contém a solução automatizada (Macro VBA) para gerar a planilha "Top Ofenders" com base em indicadores de CSAT, SQS e ABS.

## 🚀 Como Utilizar (VBA)
A macro roda diretamente dentro do seu próprio Excel, garantindo máxima segurança e compatibilidade.

1. Abra o seu arquivo Excel `TO Hitlist.xlsx` (ou o arquivo onde você cola seus dados).
2. Aperte `ALT + F11` no teclado para abrir o editor do VBA.
3. No menu superior, clique em `Inserir` > `Módulo`.
4. Abra o arquivo `macro_vba.txt` que está aqui neste repositório. Copie **todo** o código dentro dele e cole na janela branca do Módulo no Excel.
5. Feche a janela do VBA.
6. No Excel, vá na guia `Exibir` > `Macros` > `Exibir Macros` (ou aperte `ALT + F8`).
7. O arquivo possui duas macros. Selecione a que deseja usar (`GerarTopOfenders` ou `Analise_Quartil`) e clique em **Executar**.
8. O sistema vai perguntar quais meses você deseja analisar. Digite-os separados por vírgula (ex: `Jan/2026, Fev/2026`) e aperte OK.
9. A aba "Top Ofenders" será gerada automaticamente com as 4 visões solicitadas!

## 🧠 Entendendo a Matemática (Z-Score)
Para garantir um ranqueamento 100% justo onde indicadores diferentes (porcentagem vs volume) não distorcem a pontuação final, implementamos a padronização estatística (Z-Score).

1. **CSAT (Média Ponderada):** Impacto calculado por `(CSAT Agente - CSAT Global) * Volume`.
2. **SQS (Média Simples):** Impacto calculado apenas por `(SQS Agente - SQS Global)`.
3. **ABS (Média Simples vs Meta):** A pontuação baseia-se na distância entre a meta (5%) e as faltas reais: `(5% - ABS Agente)`.

**O Ranking Final:**
Esses 3 indicadores brutos são transformados em **Notas Z-Score** limitadas de `-3.0` a `+3.0`.
- `-3.0`: Ofensor Máximo extremo naquele KPI.
- `0.0`: O agente está estritamente na média da operação.
- `+3.0`: O agente é Top Performer absoluto naquele KPI.

A nota final da "Hitlist" é a **soma dessas 3 notas** (podendo variar de -9 a +9).
O ranking ordena do funcionário mais negativo (maior ofensor operacional) para o mais positivo.

## 📊 Análise de Quartis (Identificação Crônica)
A segunda macro incluída no código chama-se `Analise_Quartil`. Ela tem um objetivo diferente:
Em vez de olhar para a média global do período, ela analisa **mês a mês** e identifica os colaboradores que caíram no "Quartil 4" (os 25% piores resultados da operação naquele mês específico).

* **Regras de Negócio aplicadas:**
  * Colaboradores que tiveram 0 pesquisas ou 0 monitorias em um determinado mês são inteligentemente ignorados do cálculo daquele mês (para não distorcer o ponto de corte do quartil).
  * O Q4 de CSAT e SQS compreende os 25% com notas *mais baixas*.
  * O Q4 de ABS compreende os 25% com taxas de faltas *mais altas*.
  * A tabela lista quantas vezes (meses) cada colaborador da empresa figurou no pior quartil em cada indicador. Ideal para identificar ofensores "crônicos".

---
# OKR MPS - SQS Variation & Focus Group

## 🚀 Como Utilizar (VBA) para `okr mps.xlsx`
1. Abra o arquivo Excel `okr mps.xlsx`.
2. Aperte `ALT + F11` no teclado para abrir o editor do VBA.
3. No menu superior, clique em `Inserir` > `Módulo`.
4. Abra o arquivo `macro_vba_mps.txt` deste repositório e **copie todo o código**.
5. Cole na janela em branco do módulo no Excel e feche o VBA.
6. Na aba "Rawdata SQS", clique com o botão direito no botão "Atualizar" e selecione `Atribuir Macro`.
7. Escolha a macro chamada `BotaoAtualizar_Click` e clique em OK.

**Pronto!**
Sempre que você clicar em "Atualizar", um painel visual (com checkboxes) abrirá automaticamente. Nele, você pode selecionar Q1 a Q4 e/ou M1 a M12. Quando você clicar em "Executar", a macro fará:
- **Manage Variation SQS:** Irá extrair e calcular dinamicamente (MTD, QTDs, Weeks), agrupando a média de SQS e ordenando Top 25% (Q1), próximos (Q2), etc. O Quartil 4 representa os 25% com notas *mais baixas*.
- **Focus Group:** Vasculha e identifica somente os agentes que ficaram *abaixo da média do Q4* da semana específica e realiza a contagem retroativa `Weeks Below Q4 avg`.

**Observação Matemática (SQS CV):**
O "CV" (Coeficiente de Variação) calcula a Média Geral da Operação (`X`) dividida pelo **Desvio Padrão Populacional** das notas de todos os agentes do período analisado, gerando uma visão perfeitamente limpa sem distorções de amostra. Agentes zerados ou sem nota são sempre desconsiderados.

**Importante sobre MTD:**
Ao selecionar apenas Quarters no painel (sem selecionar um Mês específico), a macro assumirá automaticamente o **último mês do maior Quarter selecionado** para calcular e gerar a coluna do MTD (ex: ao selecionar Quarter 2, o MTD será calculado usando os dados do Mês 6).
