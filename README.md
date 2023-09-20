# as_xlsx
This is a PL/SQL package written by Anton Scheffer which allows us to export Excel XLSX files from an Oracle Database.

Histórico de alterações:

Autor: Satoru Fukugawa
Data: 27/06/2018
Descrição: 
Criado a procedure query2sheet3, que é uma cópia da procedure  query2sheet2, 
incluindo as seguintes funcionalidades: 
- Possibilidade de iniciar os dados da planilha a partir de uma determinada coluna
- Possibilidade de incluir autofiltro (p_set_autofilter) nas colunas

Autor: Satoru Fukugawa
Data: 06/08/2018
Descrição: 
Alterações na procedure query2sheet3:
- Possibilidade de formatar o relatório com estilos de tabela (p_tablestyle)
- Possibilidade de alterar a largura (p_autofitcolumns) das colunas para obter 
  o melhor ajuste
  
Autor: Satoru Fukugawa
Data: 15/08/2018
Descrição: 
Alterações na procedure query2sheet3:
- Possibilidade de alterar o alinhamento das colunas

Autor: Satoru Fukugawa
Data: 28/01/2019
Descrição: 
Alteração no cálculo da data para ser utilizado o sistema de data 1900 ao
invés de 1904

Autor: Satoru Fukugawa
Data: 03/09/2019
Descrição: 
Quando for utilizado estilo de table, o Excel não permite cabeçalhos com 
o mesmo nome. Dessa forma, caso haja cabeçalhos duplicados, o sistema irá
alterar o nome do cabeçalho incluindo um underline + sequencial no nome
do cabeçalho (Ex: COLUNA, COLUNA_2, COLUNA_3)

Autor: Satoru Fukugawa
Data: 20/03/2020
Descrição: 
Criado a função get_last_row para retornar a última linha da planilha

Autor: Satoru Fukugawa
Data: 26/03/2020
Descrição: 
Alteração na procedure new_sheet - Possibilidade de criar a planilha 
sem as linhas de grade (p_showGridLines)

Autor: Satoru Fukugawa
Data: 28/08/2020
Descrição: 
Possibilidade de utilizar fórmulas
Obs:
1) Foram testadas as principais fórmulas (soma, média, mínimo, máximo). Além da fórmula, 
   também poderá ser utilizado o cálculo matemático. Ex: A1 + A2 + A3
2) As fórmulas deverão ser passadas em inglês, e sem o igual (=) antes da fórmula. 
   Ex: SUM(A1:A3)
3) No caso de utilização de queries/table style, as fórmulas poderão ser adicionadas 
   logo após a última linha do resultado da querie. Neste caso, a rotina entenderá que 
   a última linha é uma linha de totais, e irá “incorporar” o estilo da linha de 
   totais no table style
