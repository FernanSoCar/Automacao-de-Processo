# Projeto: Automação de Processo

Este projeto tem como objetivo automatizar o controle de vendas de uma rede de lojas, gerando indicadores diários e anuais, salvando relatórios de backup e enviando e-mails automáticos para gerentes e diretoria.

## Funcionalidades
- Leitura e tratamento de dados de vendas, lojas e e-mails.
- Geração de relatórios de vendas por loja e por data.
- Cálculo de indicadores: faturamento, diversidade de produtos e ticket médio (diário e anual).
- Backup automático das planilhas de vendas de cada loja.
- Envio automático de e-mails personalizados para cada gerente com os indicadores e planilha anexa.
- Geração de rankings de faturamento (anual e diário) para a diretoria.
- Envio automático de e-mail para a diretoria com os rankings em anexo.

## Estrutura do Projeto
- `Automacao de Processo.ipynb`: Notebook principal com toda a lógica de automação.
- `Bases de Dados/`: Contém os arquivos de entrada (`Vendas.xlsx`, `Lojas.csv`, `Emails.xlsx`).
- `Backup Arquivos Lojas/`: Onde são salvos os backups das planilhas de cada loja e os rankings.
- `src/` e `tests/`: Estrutura para código Python modularizado e testes (opcional).

## Como Executar
1. Instale as dependências necessárias (pandas, win32com, etc).
2. Abra o notebook `Automacao de Processo.ipynb` no Jupyter ou VS Code.
3. Execute as células na ordem, seguindo os passos do notebook.
4. Certifique-se de que o Outlook está instalado e configurado para o envio de e-mails automáticos.

## Requisitos
- Python 3.8+
- pandas
- win32com (pywin32)
- Microsoft Outlook instalado

## Observações
- Os e-mails dos gerentes e diretoria devem estar corretamente cadastrados em `Emails.xlsx`.
- O envio de e-mails é feito via Outlook, que deve estar aberto e configurado.
- O projeto pode ser adaptado para outras plataformas de e-mail ou outros formatos de relatório.

## Exemplo de Indicadores Enviados
- Faturamento do dia e do ano
- Diversidade de produtos vendidos
- Ticket médio do dia e do ano
- Ranking das lojas por faturamento

---

Projeto desenvolvido para fins de estudo e prática de automação de processos com Python.
