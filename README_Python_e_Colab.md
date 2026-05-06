# Como usar a versão em Python (Duas Opções)

Este projeto automatiza a criação do ranking "Top Ofenders" com base nas abas `csat-sqs` e `abs` da planilha "TO Hitlist.xlsx".

## Opção 1: Google Colab (Sem instalar nada - Recomendado)

Se você não possui Python instalado ou sofre bloqueios no computador da empresa, use o Google Colab:

1. Acesse o site: [Google Colab](https://colab.research.google.com/)
2. Clique em **Novo Notebook** (ou New Notebook).
3. Na lateral esquerda, clique no ícone de uma **pasta** (Arquivos).
4. Faça o upload do seu arquivo **`TO Hitlist.xlsx`** para essa pasta.
5. Copie todo o código de dentro do arquivo `gerar_hitlist.py` e cole na célula principal do Google Colab.
6. Altere a linha de meses caso queira analisar outro período:
   `MESES_ANALISADOS = ['Jan/2026', 'Fev/2026', 'Mar/2026', 'Abr/2026', 'Mai/2026']`
7. Aperte o botão **Play** (ou Shift + Enter).
8. Ele vai gerar o resultado no próprio arquivo `TO Hitlist.xlsx`. Para baixar o resultado pronto para seu computador, clique com o botão direito no arquivo na aba esquerda e escolha **Fazer Download**.

## Opção 2: Rodando no próprio computador (Para quem tem Python instalado)

Se você preferir rodar no seu computador:

1. Instale o Python (versão 3.8 ou superior).
2. Abra o terminal (Prompt de Comando) na pasta do projeto e instale as bibliotecas necessárias:
   ```bash
   pip install pandas openpyxl
   ```
3. Abra o arquivo `gerar_hitlist.py` num bloco de notas, altere a lista `MESES_ANALISADOS` para os meses que deseja e salve.
4. Rode o script executando no terminal:
   ```bash
   python gerar_hitlist.py
   ```
5. Abra o arquivo `TO Hitlist.xlsx` e veja a aba "Top Ofenders".
