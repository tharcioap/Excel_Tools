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
7. Selecione a macro `GerarTopOfenders` e clique em **Executar**.
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
