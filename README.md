# Top Ofenders Hitlist Automation

Este repositório contém as soluções automatizadas para gerar a planilha "Top Ofenders" com base em indicadores de CSAT, SQS e ABS.

## 🚀 Como Utilizar

Você tem duas opções para rodar as automações:

### 1️⃣ Versão VBA (Direto no Excel)
1. Abra o seu arquivo `TO Hitlist.xlsx`.
2. Aperte `ALT + F11` para abrir o editor do VBA.
3. No menu superior, clique em `Inserir` > `Módulo`.
4. Abra o arquivo `macro_vba.txt` deste projeto, copie todo o código que está lá e cole dentro da janela do Módulo no Excel.
5. Feche a janela do VBA.
6. No Excel, vá na guia `Exibir` > `Macros` > `Exibir Macros`.
7. Selecione `GerarTopOfenders` e clique em **Executar**.
8. Ele vai perguntar quais meses você deseja analisar. É só digitar (ex: `Jan/2026, Fev/2026`) e apertar OK. A aba "Top Ofenders" será gerada automaticamente!

### 2️⃣ Versão Python / Google Colab
Se preferir a robustez de um script Python (ideal para arquivos gigantes ou para usar sem precisar instalar nada através do navegador), leia as instruções no arquivo:
👉 `README_Python_e_Colab.md`

## 🧠 Como os Cálculos Foram Feitos?

A lógica de projeção segue as regras:
- **Impacto Real CSAT/SQS:** `( % Agente - % Global do Período ) * Volume Real do Agente`
- **Projeção ABS (Meta 5%):**
  - Proporção: `95% / (100% - ABS%)`
  - *Volume Projetado CSAT:* Multiplicado pela proporção.
  - *Volume Projetado SQS:* Multiplicado pela proporção (Com Teto de 8 monitorias por mês analisado).
- Os Rankings são gerados do **Impacto Mais Negativo** (pior) para o mais positivo.
