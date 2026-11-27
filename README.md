# Bases-de-Dados

## Debêntures ANBIMA

A rotina coleta os arquivos do mercado secundário de debêntures da ANBIMA.
Cada data útil é salva em `downloads/debentures/AAMMDD.xlsx`, e os arquivos
individuais são reunidos em `bases/debentures_base.xlsx`, ordenados por data
de referência e sem linhas duplicadas. Os arquivos já baixados não são refeitos.

O GitHub Actions executa a rotina diariamente às 02:00 UTC (23:00 no horário
de Brasília) e também pode ser iniciado manualmente em **Actions**. A rotina
consulta a data atual em São Paulo e os seis dias úteis anteriores, permitindo
recuperar falhas recentes, mas não realiza carga histórica completa. Datas sem
arquivo disponível são tentadas novamente nas próximas execuções.

Para executar localmente, com Python 3.12:

```sh
pip install -r requirements.txt
python debentures.py
```

Execute os comandos na raiz do repositório. O workflow precisa de permissão de
escrita no conteúdo do repositório para salvar as bases. Erros de rede ou de
processamento causam falha da execução; uma nova execução pode recuperar as
datas ainda dentro da janela. A consolidação está sujeita ao limite do Excel
de 1.048.576 linhas por planilha, incluindo o cabeçalho.
