create or replace package as_xlsx
is
/**********************************************
**
** Author: Anton Scheffer
** Date: 19-02-2011
** Website: http://technology.amis.nl/blog
** See also: http://technology.amis.nl/blog/?p=10995
**
** Changelog:
**   Date: 21-02-2011
**     Added Aligment, horizontal, vertical, wrapText
**   Date: 06-03-2011
**     Added Comments, MergeCells, fixed bug for dependency on NLS-settings
**   Date: 16-03-2011
**     Added bold and italic fonts
**   Date: 22-03-2011
**     Fixed issue with timezone's set to a region(name) instead of a offset
**   Date: 08-04-2011
**     Fixed issue with XML-escaping from text
**   Date: 27-05-2011
**     Added MIT-license
**   Date: 11-08-2011
**     Fixed NLS-issue with column width
**   Date: 29-09-2011
**     Added font color
**   Date: 16-10-2011
**     fixed bug in add_string
**   Date: 26-04-2012
**     Fixed set_autofilter (only one autofilter per sheet, added _xlnm._FilterDatabase)
**     Added list_validation = drop-down
**   Date: 27-08-2013
**     Added freeze_pane
**   Date: 05-09-2013
**     Performance
**   Date: 14-07-2014
**      Added p_UseXf to query2sheet
**   Date: 23-10-2014
**      Added xml:space="preserve"
**   Date: 29-02-2016
**     Fixed issue with alignment in get_XfId
**     Thank you Bertrand Gouraud
**   Date: 01-04-2017
**     Added p_height to set_row
**   Date: 23-05-2018
**     fixed bug in add_string (thank you David Short)
**     added tabColor to new_sheet
**   Date: 13-06-2018
**     added  c_version
**     added formulas
**   Date: 12-02-2020
**     added sys_refcursor overload of query2sheet
**     use default date format in query2sheet
**     changed to date1904=false
******************************************************************************
******************************************************************************
Copyright (C) 2011, 2020 by Anton Scheffer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

******************************************************************************
******************************************** */
/*
Hist�rico de altera��es:

Autor: Satoru Fukugawa
Data: 27/06/2018
Descri��o: 
Criado a procedure query2sheet3, que � uma c�pia da procedure  query2sheet2, 
incluindo as seguintes funcionalidades: 
- Possibilidade de iniciar os dados da planilha a partir de uma determinada coluna
- Possibilidade de incluir autofiltro (p_set_autofilter) nas colunas

Autor: Satoru Fukugawa
Data: 06/08/2018
Descri��o: 
Altera��es na procedure query2sheet3:
- Possibilidade de formatar o relat�rio com estilos de tabela (p_tablestyle)
- Possibilidade de alterar a largura (p_autofitcolumns) das colunas para obter 
  o melhor ajuste
  
Autor: Satoru Fukugawa
Data: 15/08/2018
Descri��o: 
Altera��es na procedure query2sheet3:
- Possibilidade de alterar o alinhamento das colunas

Autor: Satoru Fukugawa
Data: 28/01/2019
Descri��o: 
Altera��o no c�lculo da data para ser utilizado o sistema de data 1900 ao
inv�s de 1904

Autor: Satoru Fukugawa
Data: 03/09/2019
Descri��o: 
Quando for utilizado estilo de table, o Excel n�o permite cabe�alhos com 
o mesmo nome. Dessa forma, caso haja cabe�alhos duplicados, o sistema ir�
alterar o nome do cabe�alho incluindo um underline + sequencial no nome
do cabe�alho (Ex: COLUNA, COLUNA_2, COLUNA_3)

Autor: Satoru Fukugawa
Data: 20/03/2020
Descri��o: 
Criado a fun��o get_last_row para retornar a �ltima linha da planilha

Autor: Satoru Fukugawa
Data: 26/03/2020
Descri��o: 
Altera��o na procedure new_sheet - Possibilidade de criar a planilha 
sem as linhas de grade (p_showGridLines)

Autor: Satoru Fukugawa
Data: 28/08/2020
Descri��o: 
Possibilidade de utilizar f�rmulas
Obs:
1) Foram testadas as principais f�rmulas (soma, m�dia, m�nimo, m�ximo). Al�m da f�rmula, 
   tamb�m poder� ser utilizado o c�lculo matem�tico. Ex: A1 + A2 + A3
2) As f�rmulas dever�o ser passadas em ingl�s, e sem o igual (=) antes da f�rmula. 
   Ex: SUM(A1:A3)
3) No caso de utiliza��o de queries/table style, as f�rmulas poder�o ser adicionadas 
   logo ap�s a �ltima linha do resultado da querie. Neste caso, a rotina entender� que 
   a �ltima linha � uma linha de totais, e ir� �incorporar� o estilo da linha de 
   totais no table style

*/
--
  type tp_alignment is record
    ( vertical varchar2(11)
    , horizontal varchar2(16)
    , wrapText boolean
    );
  type t_alignment is table of tp_alignment index by pls_integer;
--
  procedure clear_workbook;
--
  procedure new_sheet(p_sheetname varchar2 := null,p_showGridLines boolean := true);
--
  function OraFmt2Excel( p_format varchar2 := null )
  return varchar2;
--
  function get_numFmt( p_format varchar2 := null )
  return pls_integer;
--
  function get_font
    ( p_name varchar2
    , p_family pls_integer := 2
    , p_fontsize number := 11
    , p_theme pls_integer := 1
    , p_underline boolean := false
    , p_italic boolean := false
    , p_bold boolean := false
    , p_rgb varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
  return pls_integer;
--
  function get_fill
    ( p_patternType varchar2
    , p_fgRGB varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
  return pls_integer;
--
  function get_border
    ( p_top varchar2 := 'thin'
    , p_bottom varchar2 := 'thin'
    , p_left varchar2 := 'thin'
    , p_right varchar2 := 'thin'
    )
/*
none
thin
medium
dashed
dotted
thick
double
hair
mediumDashed
dashDot
mediumDashDot
dashDotDot
mediumDashDotDot
slantDashDot
*/
  return pls_integer;
--
  function get_alignment
    ( p_vertical varchar2 := null
    , p_horizontal varchar2 := null
    , p_wrapText boolean := null
    )
/* horizontal
center
centerContinuous
distributed
fill
general
justify
left
right
*/
/* vertical
bottom
center
distributed
justify
top
*/
  return tp_alignment;
--
  procedure cell
    ( p_col pls_integer
    , p_row pls_integer
    , p_value number
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_formula boolean := false
    );
--
  procedure cell
    ( p_col pls_integer
    , p_row pls_integer
    , p_value varchar2
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_formula boolean := false
    );
--
  procedure cell
    ( p_col pls_integer
    , p_row pls_integer
    , p_value date
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_formula boolean := false
    );
--
  procedure hyperlink
    ( p_col pls_integer
    , p_row pls_integer
    , p_url varchar2
    , p_value varchar2 := null
    , p_sheet pls_integer := null
    );
--
  procedure comment
    ( p_col pls_integer
    , p_row pls_integer
    , p_text varchar2
    , p_author varchar2 := null
    , p_width pls_integer := 150  -- pixels
    , p_height pls_integer := 100  -- pixels
    , p_sheet pls_integer := null
    );
--
  procedure mergecells
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_sheet pls_integer := null
    );
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    );
--
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_defined_name varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    );
--
  procedure defined_name
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_name varchar2
    , p_sheet pls_integer := null
    , p_localsheet pls_integer := null
    );
--
  procedure set_column_width
    ( p_col pls_integer
    , p_width number
    , p_sheet pls_integer := null
    );
--
  procedure set_column
    ( p_col pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    );
--
  procedure set_row
    ( p_row pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    );
--
  procedure set_row_height
    ( p_row pls_integer
    , p_height number
    , p_sheet pls_integer := null
    );
--
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1
    , p_sheet pls_integer := null
    );
--
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1
    , p_sheet pls_integer := null
    );
--
  procedure freeze_pane
    ( p_col pls_integer
    , p_row pls_integer
    , p_sheet pls_integer := null
    );
--
  procedure set_autofilter
    ( p_column_start pls_integer := null
    , p_column_end pls_integer := null
    , p_row_start pls_integer := null
    , p_row_end pls_integer := null
    , p_sheet pls_integer := null
    );

  procedure set_tablestyle(p_column_start  pls_integer := null,
                           p_column_end    pls_integer := null,
                           p_row_start     pls_integer := null,
                           p_row_end       pls_integer := null,
                           p_tb_style_name varchar2,
                           p_sheet         pls_integer := null);

  procedure set_autofitcolumn(
    p_value varchar2,
    p_col   pls_integer,
    p_sheet pls_integer
  );
--
  function finish
  return blob;
--
  procedure save
    ( p_directory varchar2
    , p_filename varchar2
    );
--
  procedure query2sheet
    ( p_sql VARCHAR2
    , p_column_headers boolean := true
    , p_directory varchar2 := null
    , p_filename varchar2 := null
    , p_sheet pls_integer := NULL
    );
--
  procedure query2sheet2
    ( p_sql VARCHAR2
    , p_row NUMBER
    , p_column_headers boolean := true
    , p_directory varchar2 := null
    , p_filename varchar2 := null
    , p_sheet pls_integer := NULL
    );

  procedure query2sheet3
    ( p_sql VARCHAR2
    , p_row            NUMBER := 1
    , p_col            number := 1
    , p_column_headers boolean := true
    , p_set_autofilter boolean := false
    , p_tablestyle varchar2 := null
    , p_autofitcolumns boolean := false
    , p_directory varchar2 := null
    , p_filename varchar2 := null
    , p_sheet pls_integer := NULL
    , p_sheet_name     varchar2 := null
    , p_alignment t_alignment := cast(null as t_alignment)
    );

  function get_last_row (
    p_sheet pls_integer
  ) return number;

--
/* Example
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 5, 1, 5 );
  as_xlsx.cell( 3, 1, 3 );
  as_xlsx.cell( 2, 2, 45 );
  as_xlsx.cell( 3, 2, 'Anton Scheffer', p_alignment => as_xlsx.get_alignment( p_wraptext => true ) );
  as_xlsx.cell( 1, 4, sysdate, p_fontId => as_xlsx.get_font( 'Calibri', p_rgb => 'FFFF0000' ) );
  as_xlsx.cell( 2, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( 'dd/mm/yyyy h:mm' ) );
  as_xlsx.cell( 3, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( as_xlsx.orafmt2excel( 'dd/mon/yyyy' ) ) );
  as_xlsx.cell( 5, 5, 75, p_borderId => as_xlsx.get_border( 'double', 'double', 'double', 'double' ) );
  as_xlsx.cell( 2, 3, 33 );
  as_xlsx.hyperlink( 1, 6, 'http://www.amis.nl', 'Amis site' );
  as_xlsx.cell( 1, 7, 'Some merged cells', p_alignment => as_xlsx.get_alignment( p_horizontal => 'center' ) );
  as_xlsx.mergecells( 1, 7, 3, 7 );
  for i in 1 .. 5
  loop
    as_xlsx.comment( 3, i + 3, 'Row ' || (i+3), 'Anton' );
  end loop;
  as_xlsx.new_sheet;
  as_xlsx.set_row( 1, p_fillId => as_xlsx.get_fill( 'solid', 'FFFF0000' ) ) ;
  for i in 1 .. 5
  loop
    as_xlsx.cell( 1, i, i );
    as_xlsx.cell( 2, i, i * 3 );
    as_xlsx.cell( 3, i, 'x ' || i * 3 );
  end loop;
  as_xlsx.query2sheet( 'select rownum, x.*
, case when mod( rownum, 2 ) = 0 then rownum * 3 end demo
, case when mod( rownum, 2 ) = 1 then ''demo '' || rownum end demo2 from dual x connect by rownum <= 5' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
--
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 15, p_sheet => 2 );
  as_xlsx.cell( 2, 7, 13, p_sheet => 2 );
  as_xlsx.cell( 2, 8, 17, p_sheet => 2 );
  as_xlsx.list_validation( 6, 3, 1, 6, 1, 8, p_show_error => true, p_sheet => 1 );
  as_xlsx.defined_name( 2, 6, 2, 8, 'Anton', 2 );
  as_xlsx.list_validation
    ( 6, 1, 'Anton'
    , p_style => 'information'
    , p_title => 'valid values are'
    , p_prompt => '13, 15 and 17'
    , p_show_error => true
    , p_error_title => 'Are you sure?'
    , p_error_txt => 'Valid values are: 13, 15 and 17'
    , p_sheet => 1 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
--
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.set_autofilter( 1,1, p_row_start => 5, p_row_end => 8 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 5 );
  as_xlsx.cell( 2, 7, 3 );
  as_xlsx.cell( 2, 8, 7 );
  as_xlsx.set_autofilter( 2,2, p_row_start => 5, p_row_end => 8 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
--
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  for c in 1 .. 10
  loop
    as_xlsx.cell( c, 1, 'COL' || c );
    as_xlsx.cell( c, 2, 'val' || c );
    as_xlsx.cell( c, 3, c );
  end loop;
  as_xlsx.freeze_rows( 1 );
  as_xlsx.new_sheet;
  for r in 1 .. 10
  loop
    as_xlsx.cell( 1, r, 'ROW' || r );
    as_xlsx.cell( 2, r, 'val' || r );
    as_xlsx.cell( 3, r, r );
  end loop;
  as_xlsx.freeze_cols( 3 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 3, 3, 'Start freeze' );
  as_xlsx.freeze_pane( 3,3 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
*/

/*

-- Exemplo de relat�rio com auto filtro
begin
  as_xlsx.query2sheet3
    ( p_sql => 'select ''Linha '' || rownum as "Filtro" from dual connect by level <= 10'
    , p_row => 4
    , p_col => 4
    , p_set_autofilter => true
    , p_directory => 'MY_DIR'
    , p_filename => 'autofilter3.xlsx'
   );
end;

--

-- Exemplo de relat�rio com estilo de tabela
declare
  v_estilo_tabela_clara  varchar2(20) := 'TableStyleLight';
  v_estilo_tabela_media  varchar2(20) := 'TableStyleMedium';
  v_estilo_tabela_escura varchar2(20) := 'TableStyleDark';
  v_estilo_tabela        varchar2(20);
  v_num_planilha         pls_integer;
  v_nome_planilha        varchar2(30);
  v_template_qry         varchar2(200);
  v_qry                  varchar2(200);
  v_row                  pls_integer := 1;
  v_col                  pls_integer := 1;
  
  procedure gera_excel is
  begin
    as_xlsx.query2sheet3
    ( p_sql => v_qry
    , p_row => v_row, p_col => v_col
    , p_autofitcolumns => true
    , p_tablestyle => v_estilo_tabela
    , p_sheet => v_num_planilha, p_sheet_name => v_nome_planilha
    );
  end;
begin
  v_template_qry := q'[select '#TABLE_STYLE#' as Estilo from dual union all
                    select 'Linha ' || rownum from dual connect by level <= 3]';
  v_num_planilha := 1; 
  v_nome_planilha := 'Estilos de Tabelas Claras';
  for a in 1..21 loop
    v_estilo_tabela := v_estilo_tabela_clara || a;
    v_qry := replace(v_template_qry,'#TABLE_STYLE#',v_estilo_tabela);
    gera_excel;
    if mod(a,5) = 0 then
      v_col := 1;
      v_row := v_row + 6;
    else
      v_col := v_col + 2;
    end if;
  end loop;
  
  v_col := 1; v_row := 1;
  v_num_planilha := 2;
  v_nome_planilha := 'Estilos de Tabelas M�dias';
  for b in 1..28 loop
    v_estilo_tabela := v_estilo_tabela_media || b;
    v_qry := replace(v_template_qry,'#TABLE_STYLE#',v_estilo_tabela);
    gera_excel;
    if mod(b,5) = 0 then
      v_col := 1;
      v_row := v_row + 6;
    else
      v_col := v_col + 2;
    end if;
  end loop;
  
  v_col := 1; v_row := 1;
  v_num_planilha := 3;
  v_nome_planilha := 'Estilos de Tabelas Escuras';
  for c in 1..11 loop
    v_estilo_tabela := v_estilo_tabela_escura || c;
    v_qry := replace(v_template_qry,'#TABLE_STYLE#',v_estilo_tabela);
    gera_excel;
    if mod(c,5) = 0 then
      v_col := 1;
      v_row := v_row + 6;
    else
      v_col := v_col + 2;
    end if;
  end loop;
  
  as_xlsx.save('MY_DIR','as_xlsx_table_style.xlsx');
end;

--
-- Exemplo de alinhamento utilizando querie
declare
    t_alignment as_xlsx.t_alignment;
    tmp_sql varchar2(200);
    v_sql varchar2(4000);
begin
  tmp_sql := 
  'select ' || chr(39) || 'Testando alinhamento no excel' || chr(39) || ' Campo1, ' ||
               chr(39) || 'Texto Gigante para testar a quebra de linha no excel ' ||
               'Texto Gigante para testar a quebra de linha no excel' ||chr(39) || ' Campo2 , ' ||
               '123456.78 Campo3 ' ||
  'from dual';
  for i in 1..5 loop
    v_sql := v_sql || tmp_sql || case when i < 5 then ' union all ' end;
  end loop;
    
  t_alignment(1).vertical := 'top';
  t_alignment(1).horizontal := 'center';
  t_alignment(1).wrapText := false;
    
  t_alignment(2).vertical := 'center';
  t_alignment(2).horizontal := 'right';
  t_alignment(2).wrapText := true;
    
  t_alignment(3).vertical := 'bottom';
  t_alignment(3).horizontal := 'right';
  t_alignment(3).wrapText := false;
    
  as_xlsx.query2sheet3
    ( p_sql => v_sql
    , p_tablestyle => 'TableStyleMedium9'
    , p_alignment => t_alignment
    );
    
  as_xlsx.set_column_width(1,50);
  as_xlsx.set_column_width(2,40);
  as_xlsx.set_column_width(3,25);
    
  as_xlsx.save('MY_DIR','WrapText.xlsx');
end;

--
-- Exemplo de auto dimensionamento da largura das colunas
begin
  as_xlsx.query2sheet3
    ( p_sql => 'select 1 c1, ''abcdefghijklmno'' c2 from dual'
    , p_directory => 'MY_DIR'
    , p_filename => 'TESTE_AUTOFITCOLUMNS.xlsx'
    , p_tablestyle => 'TableStyleMedium8'
    , p_autofitcolumns => true
    );
end;

--
Exemplo utilizando f�rmulas (2 exemplos)

Exemplo 1 (formula):
declare
  v_ultima_linha pls_integer;
  v_cor0         varchar2(8);
  v_cor1         varchar2(8);
  v_cor2         varchar2(8);
begin
  as_xlsx.query2sheet3(p_sql        => 'select ''ABC'' item, 50 qtde, 200 valor from dual 
                                  union select ''DEF'', 60, 300 from dual
                                  union select ''GHI'', 10, 15 from dual',
                        p_tablestyle => 'TableStyleMedium9');

  v_ultima_linha := as_xlsx.get_last_row(1);
  
  as_xlsx.cell(p_col     => 2,
                p_row     => v_ultima_linha + 1,
                p_value   => 'SUM(B2:B' || v_ultima_linha || ')',
                p_formula => true);
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 1,
                p_value   => 'SUM(C2:C' || v_ultima_linha || ')',
                p_formula => true);

  v_cor0 := 'E26B0A';
  v_cor1 := 'ff99CCFF';
  v_cor2 := 'ffC5D9F1';
  
  as_xlsx.mergecells(p_tl_col => 2,
                      p_tl_row => v_ultima_linha + 3,
                      p_br_col => 3,
                      p_br_row => v_ultima_linha + 3);
  
  as_xlsx.cell(p_col      => 2,
                p_row      => v_ultima_linha + 3,
                p_value    => 'Totais',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor0));
  as_xlsx.cell(p_col      => 3,
                p_row      => v_ultima_linha + 3,
                p_value    => ' ',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor0));
  as_xlsx.cell(p_col    => 2,
                p_row    => v_ultima_linha + 4,
                p_value  => 'Total de Itens',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 4,
                p_value   => 'COUNTA(A2:A' || v_ultima_linha || ')',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);
  as_xlsx.cell(p_col   => 2,
                p_row   => v_ultima_linha + 5,
                p_value => 'Total Qtde',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 5,
                p_value   => 'SUM(B2:B' || v_ultima_linha || ')',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);
  as_xlsx.cell(p_col   => 2,
                p_row   => v_ultima_linha + 6,
                p_value => 'Total Valor',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 6,
                p_value   => 'C2+C3+C4',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);
  as_xlsx.cell(p_col   => 2,
                p_row   => v_ultima_linha + 7,
                p_value => 'Qtde M�nima',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 7,
                p_value   => 'MIN(B2:B' || v_ultima_linha || ')',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);
  as_xlsx.cell(p_col   => 2,
                p_row   => v_ultima_linha + 8,
                p_value => 'Valor M�ximo',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 8,
                p_value   => 'MAX(C2:C' || v_ultima_linha || ')',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);
  as_xlsx.cell(p_col   => 2,
                p_row   => v_ultima_linha + 9,
                p_value => 'M�dia Qtde',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor1));
  as_xlsx.cell(p_col     => 3,
                p_row     => v_ultima_linha + 9,
                p_value   => 'AVERAGE(B2:B' || v_ultima_linha || ')',
                p_borderId => as_xlsx.get_border('thin',
                                                  'thin',
                                                  'thin',
                                                  'thin'),
                p_fillId => as_xlsx.get_fill('solid', v_cor2),
                p_formula => true);

  as_xlsx.save(p_directory => 'MY_DIR',
                p_filename  => 'formula_tabela.xlsx');
end;
--
Exemplo 2 (formula):
declare
  v_ultima_linha pls_integer;
begin
  as_xlsx.query2sheet3(
    p_sql => 'select 1 numero, 2 numero2, '' '' total from dual union 
              select 3,7, '' '' from dual union
              select 4,1, '' '' from dual union
              select 10,-3, '' '' from dual',
    p_tablestyle => 'TableStyleMedium9');
  
  v_ultima_linha := as_xlsx.get_last_row(1);
  
  for c_row in 2..v_ultima_linha loop
    as_xlsx.cell(
      p_col => 3,
      p_row => c_row,
      p_value => 'A'||c_row||'+B'||c_row,
      p_formula => true
    );
  end loop;
  
  as_xlsx.save(
    p_directory => 'MY_DIR', 
    p_filename => 'formula.xlsx');
end;

*/
end;
/
create or replace package body as_xlsx is
  --
  /* Altera��es:
  Satoru Fukugawa - SH 4372/2019 - Altera��o na formata��o da data para 1900 ao inv�s de 1904
  */
  --
  c_LOCAL_FILE_HEADER        constant raw(4) := hextoraw('504B0304'); -- Local file header signature
  c_END_OF_CENTRAL_DIRECTORY constant raw(4) := hextoraw('504B0506'); -- End of central directory signature
  --
  type tp_XF_fmt is record(
    numFmtId  pls_integer,
    fontId    pls_integer,
    fillId    pls_integer,
    borderId  pls_integer,
    alignment tp_alignment);
  type tp_col_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_row_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_widths is table of number index by pls_integer;
  type tp_heights is table of number index by pls_integer;
  type tp_cell is record(
    value number,
    style varchar2(50),
    value2 varchar2(32767 char),
    formula boolean);
  type tp_cells is table of tp_cell index by pls_integer;
  type tp_rows is table of tp_cells index by pls_integer;
  type tp_autofitcol is record(
    width number
  );
  type tp_autofitcols is table of tp_autofitcol index by pls_integer; -- indice � o nro. da coluna
  type tp_autofilter is record(
    column_start pls_integer,
    column_end   pls_integer,
    row_start    pls_integer,
    row_end      pls_integer);
  type tp_autofilters is table of tp_autofilter index by pls_integer;
  type tp_tablestyle is record(
    column_start  pls_integer,
    column_end    pls_integer,
    row_start     pls_integer,
    row_end       pls_integer,
    tb_style_name varchar2(100));
  type tp_tablestyles is table of tp_tablestyle index by pls_integer;
  type tp_hyperlink is record(
    cell varchar2(10),
    url  varchar2(1000));
  type tp_hyperlinks is table of tp_hyperlink index by pls_integer;
  subtype tp_author is varchar2(32767 char);
  type tp_authors is table of pls_integer index by tp_author;
  authors tp_authors;
  type tp_comment is record(
    text   varchar2(32767 char),
    author tp_author,
    row    pls_integer,
    column pls_integer,
    width  pls_integer,
    height pls_integer);
  type tp_comments is table of tp_comment index by pls_integer;
  type tp_mergecells is table of varchar2(21) index by pls_integer;
  type tp_validation is record(
    type             varchar2(10),
    errorstyle       varchar2(32),
    showinputmessage boolean,
    prompt           varchar2(32767 char),
    title            varchar2(32767 char),
    error_title      varchar2(32767 char),
    error_txt        varchar2(32767 char),
    showerrormessage boolean,
    formula1         varchar2(32767 char),
    formula2         varchar2(32767 char),
    allowBlank       boolean,
    sqref            varchar2(32767 char));
  type tp_validations is table of tp_validation index by pls_integer;
  type tp_sheet is record(
    rows        tp_rows,
    widths      tp_widths,
    heights     tp_heights,
    name        varchar2(100),
    freeze_rows pls_integer,
    freeze_cols pls_integer,
    autofilters tp_autofilters,
    tablestyles tp_tablestyles,
    autofitcols tp_autofitcols,
    hyperlinks  tp_hyperlinks,
    col_fmts    tp_col_fmts,
    row_fmts    tp_row_fmts,
    comments    tp_comments,
    mergecells  tp_mergecells,
    validations tp_validations,
    showGridLines boolean);
  type tp_sheets is table of tp_sheet index by pls_integer;
  type tp_numFmt is record(
    numFmtId   pls_integer,
    formatCode varchar2(100));
  type tp_numFmts is table of tp_numFmt index by pls_integer;
  type tp_fill is record(
    patternType varchar2(30),
    fgRGB       varchar2(8));
  type tp_fills is table of tp_fill index by pls_integer;
  type tp_cellXfs is table of tp_xf_fmt index by pls_integer;
  type tp_font is record(
    name      varchar2(100),
    family    pls_integer,
    fontsize  number,
    theme     pls_integer,
    RGB       varchar2(8),
    underline boolean,
    italic    boolean,
    bold      boolean);
  type tp_fonts is table of tp_font index by pls_integer;
  type tp_border is record(
    top    varchar2(17),
    bottom varchar2(17),
    left   varchar2(17),
    right  varchar2(17));
  type tp_borders is table of tp_border index by pls_integer;
  type tp_numFmtIndexes is table of pls_integer index by pls_integer;
  type tp_strings is table of pls_integer index by varchar2(32767 char);
  type tp_str_ind is table of varchar2(32767 char) index by pls_integer;
  type tp_defined_name is record(
    name  varchar2(32767 char),
    ref   varchar2(32767 char),
    sheet pls_integer);
  type tp_defined_names is table of tp_defined_name index by pls_integer;
  type tp_book is record(
    sheets        tp_sheets,
    strings       tp_strings,
    str_ind       tp_str_ind,
    str_cnt       pls_integer := 0,
    fonts         tp_fonts,
    fills         tp_fills,
    borders       tp_borders,
    numFmts       tp_numFmts,
    cellXfs       tp_cellXfs,
    numFmtIndexes tp_numFmtIndexes,
    defined_names tp_defined_names);
  workbook tp_book;

  has_formula boolean := false;
  --
  procedure blob2file(p_blob      blob,
                      p_directory varchar2 := 'MY_DIR',
                      p_filename  varchar2 := 'my.xlsx') is
    t_fh  utl_file.file_type;
    t_len pls_integer := 32767;
  begin
    t_fh := utl_file.fopen(p_directory, p_filename, 'wb');
    for i in 0 .. trunc((dbms_lob.getlength(p_blob) - 1) / t_len) loop
      utl_file.put_raw(t_fh, dbms_lob.substr(p_blob, t_len, i * t_len + 1));
    end loop;
    utl_file.fclose(t_fh);
  end;
  --
  function raw2num(p_raw raw, p_len integer, p_pos integer) return number is
  begin
    return utl_raw.cast_to_binary_integer(utl_raw.substr(p_raw,
                                                         p_pos,
                                                         p_len),
                                          utl_raw.little_endian);
  end;
  --
  function little_endian(p_big number, p_bytes pls_integer := 4) return raw is
  begin
    return utl_raw.substr(utl_raw.cast_from_binary_integer(p_big,
                                                           utl_raw.little_endian),
                          1,
                          p_bytes);
  end;
  --
  function blob2num(p_blob blob, p_len integer, p_pos integer) return number is
  begin
    return utl_raw.cast_to_binary_integer(dbms_lob.substr(p_blob,
                                                          p_len,
                                                          p_pos),
                                          utl_raw.little_endian);
  end;
  --
  procedure add1file(p_zipped_blob in out blob,
                     p_name        varchar2,
                     p_content     blob) is
    t_now        date;
    t_blob       blob;
    t_len        integer;
    t_clen       integer;
    t_crc32      raw(4) := hextoraw('00000000');
    t_compressed boolean := false;
    t_name       raw(32767);
  begin
    t_now := sysdate;
    t_len := nvl(dbms_lob.getlength(p_content), 0);
    if t_len > 0 then
      t_blob       := utl_compress.lz_compress(p_content);
      t_clen       := dbms_lob.getlength(t_blob) - 18;
      t_compressed := t_clen < t_len;
      t_crc32      := dbms_lob.substr(t_blob, 4, t_clen + 11);
    end if;
    if not t_compressed then
      t_clen := t_len;
      t_blob := p_content;
    end if;
    if p_zipped_blob is null then
      dbms_lob.createtemporary(p_zipped_blob, true);
    end if;
    t_name := utl_i18n.string_to_raw(p_name, 'AL32UTF8');
    dbms_lob.append(p_zipped_blob,
                    utl_raw.concat(c_LOCAL_FILE_HEADER -- Local file header signature
                                  ,
                                   hextoraw('1400') -- version 2.0
                                  ,
                                   case when
                                   t_name =
                                   utl_i18n.string_to_raw(p_name, 'US8PC437') then
                                   hextoraw('0000') -- no General purpose bits
                                   else hextoraw('0008') -- set Language encoding flag (EFS)
                                   end,
                                   case when t_compressed then
                                   hextoraw('0800') -- deflate
                                   else hextoraw('0000') -- stored
                                   end,
                                   little_endian(to_number(to_char(t_now,
                                                                   'ss')) / 2 +
                                                 to_number(to_char(t_now,
                                                                   'mi')) * 32 +
                                                 to_number(to_char(t_now,
                                                                   'hh24')) * 2048,
                                                 2) -- File last modification time
                                  ,
                                   little_endian(to_number(to_char(t_now,
                                                                   'dd')) +
                                                 to_number(to_char(t_now,
                                                                   'mm')) * 32 +
                                                 (to_number(to_char(t_now,
                                                                    'yyyy')) - 1980) * 512,
                                                 2) -- File last modification date
                                  ,
                                   t_crc32 -- CRC-32
                                  ,
                                   little_endian(t_clen) -- compressed size
                                  ,
                                   little_endian(t_len) -- uncompressed size
                                  ,
                                   little_endian(utl_raw.length(t_name), 2) -- File name length
                                  ,
                                   hextoraw('0000') -- Extra field length
                                  ,
                                   t_name -- File name
                                   ));
    if t_compressed then
      dbms_lob.copy(p_zipped_blob,
                    t_blob,
                    t_clen,
                    dbms_lob.getlength(p_zipped_blob) + 1,
                    11); -- compressed content
    elsif t_clen > 0 then
      dbms_lob.copy(p_zipped_blob,
                    t_blob,
                    t_clen,
                    dbms_lob.getlength(p_zipped_blob) + 1,
                    1); --  content
    end if;
    if dbms_lob.istemporary(t_blob) = 1 then
      dbms_lob.freetemporary(t_blob);
    end if;
  end;
  --
  procedure finish_zip(p_zipped_blob in out blob) is
    t_cnt             pls_integer := 0;
    t_offs            integer;
    t_offs_dir_header integer;
    t_offs_end_header integer;
    t_comment         raw(32767) := utl_raw.cast_to_raw('Implementation by Anton Scheffer');
  begin
    t_offs_dir_header := dbms_lob.getlength(p_zipped_blob);
    t_offs            := 1;
    while dbms_lob.substr(p_zipped_blob,
                          utl_raw.length(c_LOCAL_FILE_HEADER),
                          t_offs) = c_LOCAL_FILE_HEADER loop
      t_cnt := t_cnt + 1;
      dbms_lob.append(p_zipped_blob,
                      utl_raw.concat(hextoraw('504B0102') -- Central directory file header signature
                                    ,
                                     hextoraw('1400') -- version 2.0
                                    ,
                                     dbms_lob.substr(p_zipped_blob,
                                                     26,
                                                     t_offs + 4),
                                     hextoraw('0000') -- File comment length
                                    ,
                                     hextoraw('0000') -- Disk number where file starts
                                    ,
                                     hextoraw('0000') -- Internal file attributes =>
                                     --     0000 binary file
                                     --     0100 (ascii)text file
                                    ,
                                     case when dbms_lob.substr(p_zipped_blob,
                                                     1,
                                                     t_offs + 30 +
                                                     blob2num(p_zipped_blob,
                                                              2,
                                                              t_offs + 26) - 1) in
                                     (hextoraw('2F') -- /
                                    ,
                                      hextoraw('5C') -- \
                                      ) then hextoraw('10000000') -- a directory/folder
                                     else hextoraw('2000B681') -- a file
                                     end -- External file attributes
                                    ,
                                     little_endian(t_offs - 1) -- Relative offset of local file header
                                    ,
                                     dbms_lob.substr(p_zipped_blob,
                                                     blob2num(p_zipped_blob,
                                                              2,
                                                              t_offs + 26),
                                                     t_offs + 30) -- File name
                                     ));
      t_offs := t_offs + 30 + blob2num(p_zipped_blob, 4, t_offs + 18) -- compressed size
                + blob2num(p_zipped_blob, 2, t_offs + 26) -- File name length
                + blob2num(p_zipped_blob, 2, t_offs + 28); -- Extra field length
    end loop;
    t_offs_end_header := dbms_lob.getlength(p_zipped_blob);
    dbms_lob.append(p_zipped_blob,
                    utl_raw.concat(c_END_OF_CENTRAL_DIRECTORY -- End of central directory signature
                                  ,
                                   hextoraw('0000') -- Number of this disk
                                  ,
                                   hextoraw('0000') -- Disk where central directory starts
                                  ,
                                   little_endian(t_cnt, 2) -- Number of central directory records on this disk
                                  ,
                                   little_endian(t_cnt, 2) -- Total number of central directory records
                                  ,
                                   little_endian(t_offs_end_header -
                                                 t_offs_dir_header) -- Size of central directory
                                  ,
                                   little_endian(t_offs_dir_header) -- Offset of start of central directory, relative to start of archive
                                  ,
                                   little_endian(nvl(utl_raw.length(t_comment),
                                                     0),
                                                 2) -- ZIP file comment length
                                  ,
                                   t_comment));
  end;
  --
  function alfan_col(p_col pls_integer) return varchar2 is
  begin
    return case when p_col > 702 then chr(64 + trunc((p_col - 27) / 676)) || chr(65 +
                                                                                 mod(trunc((p_col - 1) / 26) - 1,
                                                                                     26)) || chr(65 +
                                                                                                 mod(p_col - 1,
                                                                                                     26)) when p_col > 26 then chr(64 +
                                                                                                                                   trunc((p_col - 1) / 26)) || chr(65 +
                                                                                                                                                                   mod(p_col - 1,
                                                                                                                                                                       26)) else chr(64 +
                                                                                                                                                                                     p_col) end;
  end;
  --
  function col_alfan(p_col varchar2) return pls_integer is
  begin
    return ascii(substr(p_col, -1)) - 64 + nvl((ascii(substr(p_col, -2, 1)) - 64) * 26,
                                               0) + nvl((ascii(substr(p_col,
                                                                      -3,
                                                                      1)) - 64) * 676,
                                                        0);
  end;
  --
  procedure clear_workbook is
    t_row_ind pls_integer;
    t_clear_rows tp_rows;
  begin
    for s in 1 .. workbook.sheets.count() loop
      workbook.sheets( s ).rows := t_clear_rows;
      /*t_row_ind := workbook.sheets(s).rows.first();
      while t_row_ind is not null loop
        workbook.sheets(s).rows(t_row_ind).delete();
        t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
      end loop;*/
      workbook.sheets(s).rows.delete();
      workbook.sheets(s).widths.delete();
      workbook.sheets(s).heights.delete();
      workbook.sheets(s).autofilters.delete();
      workbook.sheets(s).tablestyles.delete();
      workbook.sheets(s).autofitcols.delete();
      workbook.sheets(s).hyperlinks.delete();
      workbook.sheets(s).col_fmts.delete();
      workbook.sheets(s).row_fmts.delete();
      workbook.sheets(s).comments.delete();
      workbook.sheets(s).mergecells.delete();
      workbook.sheets(s).validations.delete();
    end loop;
    workbook.strings.delete();
    workbook.str_ind.delete();
    workbook.fonts.delete();
    workbook.fills.delete();
    workbook.borders.delete();
    workbook.numFmts.delete();
    workbook.cellXfs.delete();
    workbook.defined_names.delete();
    workbook := null;
  end;
  --
  procedure new_sheet(p_sheetname varchar2 := null,p_showGridLines boolean := true) is
    t_nr  pls_integer := workbook.sheets.count() + 1;
    t_ind pls_integer;
  begin
    workbook.sheets(t_nr).name := nvl(dbms_xmlgen.convert(translate(p_sheetname,
                                                                    'a/\[]*:?',
                                                                    'a')),
                                      'Sheet' || t_nr);
    workbook.sheets(t_nr).showGridLines := p_showGridLines;
    if workbook.strings.count() = 0 then
      workbook.str_cnt := 0;
    end if;
    if workbook.fonts.count() = 0 then
      t_ind := get_font('Calibri');
    end if;
    if workbook.fills.count() = 0 then
      t_ind := get_fill('none');
      t_ind := get_fill('gray125');
    end if;
    if workbook.borders.count() = 0 then
      t_ind := get_border('', '', '', '');
    end if;
  end;
  --
  procedure new_sheet2(p_sheet_n pls_integer := null, p_sheetname varchar2 := null) is
    t_nr  pls_integer := workbook.sheets.count() + 1;
    t_ind pls_integer;
  begin
    if p_sheet_n is not null then
      t_nr := p_sheet_n;
    end if;

    workbook.sheets(t_nr).name := nvl(dbms_xmlgen.convert(translate(p_sheetname,
                                                                    'a/\[]*:?',
                                                                    'a')),
                                      'Sheet' || t_nr);
    if workbook.strings.count() = 0 then
      workbook.str_cnt := 0;
    end if;
    if workbook.fonts.count() = 0 then
      t_ind := get_font('Calibri');
    end if;
    if workbook.fills.count() = 0 then
      t_ind := get_fill('none');
      t_ind := get_fill('gray125');
    end if;
    if workbook.borders.count() = 0 then
      t_ind := get_border('', '', '', '');
    end if;
  end;
  --
  procedure set_col_width(p_sheet  pls_integer,
                          p_col    pls_integer,
                          p_format varchar2) is
    t_width  number;
    t_nr_chr pls_integer;
  begin
    if p_format is null then
      return;
    end if;
    if instr(p_format, ';') > 0 then
      t_nr_chr := length(translate(substr(p_format,
                                          1,
                                          instr(p_format, ';') - 1),
                                   'a\"',
                                   'a'));
    else
      t_nr_chr := length(translate(p_format, 'a\"', 'a'));
    end if;
    t_width := trunc((t_nr_chr * 7 + 5) / 7 * 256) / 256; -- assume default 11 point Calibri
    if workbook.sheets(p_sheet).widths.exists(p_col) then
      workbook.sheets(p_sheet).widths(p_col) := greatest(workbook.sheets(p_sheet)
                                                         .widths(p_col),
                                                         t_width);
    else
      workbook.sheets(p_sheet).widths(p_col) := greatest(t_width, 8.43);
    end if;
  end;
  --
  function OraFmt2Excel(p_format varchar2 := null) return varchar2 is
    t_format varchar2(1000) := substr(p_format, 1, 1000);
  begin
    t_format := replace(replace(t_format, 'hh24', 'hh'), 'hh12', 'hh');
    t_format := replace(t_format, 'mi', 'mm');
    t_format := replace(replace(replace(t_format, 'AM', '~~'), 'PM', '~~'),
                        '~~',
                        'AM/PM');
    t_format := replace(replace(replace(t_format, 'am', '~~'), 'pm', '~~'),
                        '~~',
                        'AM/PM');
    t_format := replace(replace(t_format, 'day', 'DAY'), 'DAY', 'dddd');
    t_format := replace(replace(t_format, 'dy', 'DY'), 'DAY', 'ddd');
    t_format := replace(replace(t_format, 'RR', 'RR'), 'RR', 'YY');
    t_format := replace(replace(t_format, 'month', 'MONTH'),
                        'MONTH',
                        'mmmm');
    t_format := replace(replace(t_format, 'mon', 'MON'), 'MON', 'mmm');
    return t_format;
  end;
  --
  function get_numFmt(p_format varchar2 := null) return pls_integer is
    t_cnt      pls_integer;
    t_numFmtId pls_integer;
  begin
    if p_format is null then
      return 0;
    end if;
    t_cnt := workbook.numFmts.count();
    for i in 1 .. t_cnt loop
      if workbook.numFmts(i).formatCode = p_format then
        t_numFmtId := workbook.numFmts(i).numFmtId;
        exit;
      end if;
    end loop;
    if t_numFmtId is null then
      t_numFmtId := case
                      when t_cnt = 0 then
                       164
                      else
                       workbook.numFmts(t_cnt).numFmtId + 1
                    end;
      t_cnt := t_cnt + 1;
      workbook.numFmts(t_cnt).numFmtId := t_numFmtId;
      workbook.numFmts(t_cnt).formatCode := p_format;
      workbook.numFmtIndexes(t_numFmtId) := t_cnt;
    end if;
    return t_numFmtId;
  end;
  --
  function get_font(p_name      varchar2,
                    p_family    pls_integer := 2,
                    p_fontsize  number := 11,
                    p_theme     pls_integer := 1,
                    p_underline boolean := false,
                    p_italic    boolean := false,
                    p_bold      boolean := false,
                    p_rgb       varchar2 := null -- this is a hex ALPHA Red Green Blue value
                    ) return pls_integer is
    t_ind pls_integer;
  begin
    if workbook.fonts.count() > 0 then
      for f in 0 .. workbook.fonts.count() - 1 loop
        if (workbook.fonts(f)
           .name = p_name and workbook.fonts(f).family = p_family and workbook.fonts(f)
           .fontsize = p_fontsize and workbook.fonts(f).theme = p_theme and workbook.fonts(f)
           .underline = p_underline and workbook.fonts(f).italic = p_italic and workbook.fonts(f)
           .bold = p_bold and
            (workbook.fonts(f)
            .rgb = p_rgb or (workbook.fonts(f).rgb is null and p_rgb is null))) then
          return f;
        end if;
      end loop;
    end if;
    t_ind := workbook.fonts.count();
    workbook.fonts(t_ind).name := p_name;
    workbook.fonts(t_ind).family := p_family;
    workbook.fonts(t_ind).fontsize := p_fontsize;
    workbook.fonts(t_ind).theme := p_theme;
    workbook.fonts(t_ind).underline := p_underline;
    workbook.fonts(t_ind).italic := p_italic;
    workbook.fonts(t_ind).bold := p_bold;
    workbook.fonts(t_ind).rgb := p_rgb;
    return t_ind;
  end;
  --
  function get_fill(p_patternType varchar2, p_fgRGB varchar2 := null)
    return pls_integer is
    t_ind pls_integer;
  begin
    if workbook.fills.count() > 0 then
      for f in 0 .. workbook.fills.count() - 1 loop
        if (workbook.fills(f)
           .patternType = p_patternType and
            nvl(workbook.fills(f).fgRGB, 'x') = nvl(upper(p_fgRGB), 'x')) then
          return f;
        end if;
      end loop;
    end if;
    t_ind := workbook.fills.count();
    workbook.fills(t_ind).patternType := p_patternType;
    workbook.fills(t_ind).fgRGB := upper(p_fgRGB);
    return t_ind;
  end;
  --
  function get_border(p_top    varchar2 := 'thin',
                      p_bottom varchar2 := 'thin',
                      p_left   varchar2 := 'thin',
                      p_right  varchar2 := 'thin') return pls_integer is
    t_ind pls_integer;
  begin
    if workbook.borders.count() > 0 then
      for b in 0 .. workbook.borders.count() - 1 loop
        if (nvl(workbook.borders(b).top, 'x') = nvl(p_top, 'x') and
           nvl(workbook.borders(b).bottom, 'x') = nvl(p_bottom, 'x') and
           nvl(workbook.borders(b).left, 'x') = nvl(p_left, 'x') and
           nvl(workbook.borders(b).right, 'x') = nvl(p_right, 'x')) then
          return b;
        end if;
      end loop;
    end if;
    t_ind := workbook.borders.count();
    workbook.borders(t_ind).top := p_top;
    workbook.borders(t_ind).bottom := p_bottom;
    workbook.borders(t_ind).left := p_left;
    workbook.borders(t_ind).right := p_right;
    return t_ind;
  end;
  --
  function get_alignment(p_vertical   varchar2 := null,
                         p_horizontal varchar2 := null,
                         p_wrapText   boolean := null) return tp_alignment is
    t_rv tp_alignment;
  begin
    t_rv.vertical   := p_vertical;
    t_rv.horizontal := p_horizontal;
    t_rv.wrapText   := p_wrapText;
    return t_rv;
  end;
  --
  function get_XfId(p_sheet     pls_integer,
                    p_col       pls_integer,
                    p_row       pls_integer,
                    p_numFmtId  pls_integer := null,
                    p_fontId    pls_integer := null,
                    p_fillId    pls_integer := null,
                    p_borderId  pls_integer := null,
                    p_alignment tp_alignment := null) return varchar2 is
    t_cnt    pls_integer;
    t_XfId   pls_integer;
    t_XF     tp_XF_fmt;
    t_col_XF tp_XF_fmt;
    t_row_XF tp_XF_fmt;
  begin
    if workbook.sheets(p_sheet).col_fmts.exists(p_col) then
      t_col_XF := workbook.sheets(p_sheet).col_fmts(p_col);
    end if;
    if workbook.sheets(p_sheet).row_fmts.exists(p_row) then
      t_row_XF := workbook.sheets(p_sheet).row_fmts(p_row);
    end if;
    t_XF.numFmtId  := coalesce(p_numFmtId,
                               t_col_XF.numFmtId,
                               t_row_XF.numFmtId,
                               0);
    t_XF.fontId    := coalesce(p_fontId,
                               t_col_XF.fontId,
                               t_row_XF.fontId,
                               0);
    t_XF.fillId    := coalesce(p_fillId,
                               t_col_XF.fillId,
                               t_row_XF.fillId,
                               0);
    t_XF.borderId  := coalesce(p_borderId,
                               t_col_XF.borderId,
                               t_row_XF.borderId,
                               0);
    t_XF.alignment := coalesce(p_alignment,
                               t_col_XF.alignment,
                               t_row_XF.alignment);
    if (t_XF.numFmtId + t_XF.fontId + t_XF.fillId + t_XF.borderId = 0 and
       t_XF.alignment.vertical is null and
       t_XF.alignment.horizontal is null and
       not nvl(t_XF.alignment.wrapText, false)) then
      return '';
    end if;
    if t_XF.numFmtId > 0 then
      set_col_width(p_sheet,
                    p_col,
                    workbook.numFmts(workbook.numFmtIndexes(t_XF.numFmtId))
                    .formatCode);
    end if;
    t_cnt := workbook.cellXfs.count();
    for i in 1 .. t_cnt loop
      if (workbook.cellXfs(i)
         .numFmtId = t_XF.numFmtId and workbook.cellXfs(i)
         .fontId = t_XF.fontId and workbook.cellXfs(i).fillId = t_XF.fillId and workbook.cellXfs(i)
         .borderId = t_XF.borderId and
          nvl(workbook.cellXfs(i).alignment.vertical, 'x') =
          nvl(t_XF.alignment.vertical, 'x') and
          nvl(workbook.cellXfs(i).alignment.horizontal, 'x') =
          nvl(t_XF.alignment.horizontal, 'x') and
          nvl(workbook.cellXfs(i).alignment.wrapText, false) =
          nvl(t_XF.alignment.wrapText, false)) then
        t_XfId := i;
        exit;
      end if;
    end loop;
    if t_XfId is null then
      t_cnt := t_cnt + 1;
      t_XfId := t_cnt;
      workbook.cellXfs(t_cnt) := t_XF;
    end if;
    return 's="' || t_XfId || '"';
  end;
  --
  procedure cell(p_col       pls_integer,
                 p_row       pls_integer,
                 p_value     number,
                 p_numFmtId  pls_integer := null,
                 p_fontId    pls_integer := null,
                 p_fillId    pls_integer := null,
                 p_borderId  pls_integer := null,
                 p_alignment tp_alignment := null,
                 p_sheet     pls_integer := null,
                 p_formula boolean := false) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := p_value;
    workbook.sheets(t_sheet).rows(p_row)(p_col).value2 := to_char(p_value);
    workbook.sheets(t_sheet).rows(p_row)(p_col).formula := p_formula;
    if p_formula then
      has_formula := true;
    end if;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := null;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId(t_sheet,
                                                                  p_col,
                                                                  p_row,
                                                                  p_numFmtId,
                                                                  p_fontId,
                                                                  p_fillId,
                                                                  p_borderId,
                                                                  p_alignment);
  end;
  --
  function add_string(p_string varchar2) return pls_integer is
    t_cnt pls_integer;
  begin
    if workbook.strings.exists(p_string) then
      t_cnt := workbook.strings(p_string);
    else
      t_cnt := workbook.strings.count();
      workbook.str_ind(t_cnt) := p_string;
      workbook.strings(nvl(p_string, '')) := t_cnt;
    end if;
    workbook.str_cnt := workbook.str_cnt + 1;
    return t_cnt;
  end;
  --
  procedure cell(p_col       pls_integer,
                 p_row       pls_integer,
                 p_value     varchar2,
                 p_numFmtId  pls_integer := null,
                 p_fontId    pls_integer := null,
                 p_fillId    pls_integer := null,
                 p_borderId  pls_integer := null,
                 p_alignment tp_alignment := null,
                 p_sheet     pls_integer := null,
                 p_formula boolean := false) is
    t_sheet     pls_integer := nvl(p_sheet, workbook.sheets.count());
    t_alignment tp_alignment := p_alignment;
  begin
    if not p_formula then
      workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(p_value);
    end if;
    workbook.sheets(t_sheet).rows(p_row)(p_col).value2 := p_value;
    workbook.sheets(t_sheet).rows(p_row)(p_col).formula := p_formula;
    if p_formula then
      has_formula := true;
    end if;
    if t_alignment.wrapText is null and instr(p_value, chr(13)) > 0 then
      t_alignment.wrapText := true;
    end if;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' ||
                                                         get_XfId(t_sheet,
                                                                  p_col,
                                                                  p_row,
                                                                  p_numFmtId,
                                                                  p_fontId,
                                                                  p_fillId,
                                                                  p_borderId,
                                                                  t_alignment);
  end;
  --
  procedure cell(p_col       pls_integer,
                 p_row       pls_integer,
                 p_value     date,
                 p_numFmtId  pls_integer := null,
                 p_fontId    pls_integer := null,
                 p_fillId    pls_integer := null,
                 p_borderId  pls_integer := null,
                 p_alignment tp_alignment := null,
                 p_sheet     pls_integer := null,
                 p_formula boolean := false) is
    t_numFmtId pls_integer := p_numFmtId;
    t_sheet    pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := p_value -
                                                         to_date('01-01-1900',
                                                                 'DD-MM-YYYY') +
                                                         2;
    workbook.sheets(t_sheet).rows(p_row)(p_col).value2 := p_value;
    workbook.sheets(t_sheet).rows(p_row)(p_col).formula := p_formula;
    if p_formula then
      has_formula := true;
    end if;
    if t_numFmtId is null and
       not (workbook.sheets(t_sheet).col_fmts.exists(p_col) and workbook.sheets(t_sheet).col_fmts(p_col)
        .numFmtId is not null) and
       not (workbook.sheets(t_sheet).row_fmts.exists(p_row) and workbook.sheets(t_sheet).row_fmts(p_row)
        .numFmtId is not null) then
      t_numFmtId := get_numFmt('dd/mm/yyyy');
    end if;
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := get_XfId(t_sheet,
                                                                  p_col,
                                                                  p_row,
                                                                  t_numFmtId,
                                                                  p_fontId,
                                                                  p_fillId,
                                                                  p_borderId,
                                                                  p_alignment);
  end;
  --
  procedure hyperlink(p_col   pls_integer,
                      p_row   pls_integer,
                      p_url   varchar2,
                      p_value varchar2 := null,
                      p_sheet pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).rows(p_row)(p_col).value := add_string(nvl(p_value,
                                                                        p_url));
    workbook.sheets(t_sheet).rows(p_row)(p_col).style := 't="s" ' ||
                                                         get_XfId(t_sheet,
                                                                  p_col,
                                                                  p_row,
                                                                  '',
                                                                  get_font('Calibri',
                                                                           p_theme     => 10,
                                                                           p_underline => true));
    t_ind := workbook.sheets(t_sheet).hyperlinks.count() + 1;
    workbook.sheets(t_sheet).hyperlinks(t_ind).cell := alfan_col(p_col) ||
                                                       p_row;
    workbook.sheets(t_sheet).hyperlinks(t_ind).url := p_url;
  end;
  --
  procedure comment(p_col    pls_integer,
                    p_row    pls_integer,
                    p_text   varchar2,
                    p_author varchar2 := null,
                    p_width  pls_integer := 150,
                    p_height pls_integer := 100,
                    p_sheet  pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := workbook.sheets(t_sheet).comments.count() + 1;
    workbook.sheets(t_sheet).comments(t_ind).row := p_row;
    workbook.sheets(t_sheet).comments(t_ind).column := p_col;
    workbook.sheets(t_sheet).comments(t_ind).text := dbms_xmlgen.convert(p_text);
    workbook.sheets(t_sheet).comments(t_ind).author := dbms_xmlgen.convert(p_author);
    workbook.sheets(t_sheet).comments(t_ind).width := p_width;
    workbook.sheets(t_sheet).comments(t_ind).height := p_height;
  end;
  --
  procedure mergecells(p_tl_col pls_integer -- top left
                      ,
                       p_tl_row pls_integer,
                       p_br_col pls_integer -- bottom right
                      ,
                       p_br_row pls_integer,
                       p_sheet  pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := workbook.sheets(t_sheet).mergecells.count() + 1;
    workbook.sheets(t_sheet).mergecells(t_ind) := alfan_col(p_tl_col) ||
                                                  p_tl_row || ':' ||
                                                  alfan_col(p_br_col) ||
                                                  p_br_row;
  end;
  --
  procedure add_validation(p_type        varchar2,
                           p_sqref       varchar2,
                           p_style       varchar2 := 'stop' -- stop, warning, information
                          ,
                           p_formula1    varchar2 := null,
                           p_formula2    varchar2 := null,
                           p_title       varchar2 := null,
                           p_prompt      varchar := null,
                           p_show_error  boolean := false,
                           p_error_title varchar2 := null,
                           p_error_txt   varchar2 := null,
                           p_sheet       pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := workbook.sheets(t_sheet).validations.count() + 1;
    workbook.sheets(t_sheet).validations(t_ind).type := p_type;
    workbook.sheets(t_sheet).validations(t_ind).errorstyle := p_style;
    workbook.sheets(t_sheet).validations(t_ind).sqref := p_sqref;
    workbook.sheets(t_sheet).validations(t_ind).formula1 := p_formula1;
    workbook.sheets(t_sheet).validations(t_ind).error_title := p_error_title;
    workbook.sheets(t_sheet).validations(t_ind).error_txt := p_error_txt;
    workbook.sheets(t_sheet).validations(t_ind).title := p_title;
    workbook.sheets(t_sheet).validations(t_ind).prompt := p_prompt;
    workbook.sheets(t_sheet).validations(t_ind).showerrormessage := p_show_error;
  end;
  --
  procedure list_validation(p_sqref_col   pls_integer,
                            p_sqref_row   pls_integer,
                            p_tl_col      pls_integer -- top left
                           ,
                            p_tl_row      pls_integer,
                            p_br_col      pls_integer -- bottom right
                           ,
                            p_br_row      pls_integer,
                            p_style       varchar2 := 'stop' -- stop, warning, information
                           ,
                            p_title       varchar2 := null,
                            p_prompt      varchar := null,
                            p_show_error  boolean := false,
                            p_error_title varchar2 := null,
                            p_error_txt   varchar2 := null,
                            p_sheet       pls_integer := null) is
  begin
    add_validation('list',
                   alfan_col(p_sqref_col) || p_sqref_row,
                   p_style => lower(p_style),
                   p_formula1 => '$' || alfan_col(p_tl_col) || '$' ||
                                 p_tl_row || ':$' || alfan_col(p_br_col) || '$' ||
                                 p_br_row,
                   p_title => p_title,
                   p_prompt => p_prompt,
                   p_show_error => p_show_error,
                   p_error_title => p_error_title,
                   p_error_txt => p_error_txt,
                   p_sheet => p_sheet);
  end;
  --
  procedure list_validation(p_sqref_col    pls_integer,
                            p_sqref_row    pls_integer,
                            p_defined_name varchar2,
                            p_style        varchar2 := 'stop' -- stop, warning, information
                           ,
                            p_title        varchar2 := null,
                            p_prompt       varchar := null,
                            p_show_error   boolean := false,
                            p_error_title  varchar2 := null,
                            p_error_txt    varchar2 := null,
                            p_sheet        pls_integer := null) is
  begin
    add_validation('list',
                   alfan_col(p_sqref_col) || p_sqref_row,
                   p_style => lower(p_style),
                   p_formula1 => p_defined_name,
                   p_title => p_title,
                   p_prompt => p_prompt,
                   p_show_error => p_show_error,
                   p_error_title => p_error_title,
                   p_error_txt => p_error_txt,
                   p_sheet => p_sheet);
  end;
  --
  procedure defined_name(p_tl_col     pls_integer -- top left
                        ,
                         p_tl_row     pls_integer,
                         p_br_col     pls_integer -- bottom right
                        ,
                         p_br_row     pls_integer,
                         p_name       varchar2,
                         p_sheet      pls_integer := null,
                         p_localsheet pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := workbook.defined_names.count() + 1;
    workbook.defined_names(t_ind).name := p_name;
    workbook.defined_names(t_ind).ref := 'Sheet' || t_sheet || '!$' ||
                                         alfan_col(p_tl_col) || '$' ||
                                         p_tl_row || ':$' ||
                                         alfan_col(p_br_col) || '$' ||
                                         p_br_row;
    workbook.defined_names(t_ind).sheet := p_localsheet;
  end;
  --
  procedure set_column_width(p_col   pls_integer,
                             p_width number,
                             p_sheet pls_integer := null) is
  begin
    workbook.sheets(nvl(p_sheet, workbook.sheets.count())).widths(p_col) := p_width;
  end;
  --
  procedure set_column(p_col       pls_integer,
                       p_numFmtId  pls_integer := null,
                       p_fontId    pls_integer := null,
                       p_fillId    pls_integer := null,
                       p_borderId  pls_integer := null,
                       p_alignment tp_alignment := null,
                       p_sheet     pls_integer := null) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).col_fmts(p_col).numFmtId := p_numFmtId;
    workbook.sheets(t_sheet).col_fmts(p_col).fontId := p_fontId;
    workbook.sheets(t_sheet).col_fmts(p_col).fillId := p_fillId;
    workbook.sheets(t_sheet).col_fmts(p_col).borderId := p_borderId;
    workbook.sheets(t_sheet).col_fmts(p_col).alignment := p_alignment;
  end;
  --
  procedure set_row_height
    ( p_row pls_integer
    , p_height number
    , p_sheet pls_integer := null
    ) is
  begin
    workbook.sheets(nvl(p_sheet, workbook.sheets.count())).heights(p_row) := p_height;
  end;
  --
  procedure set_row(p_row       pls_integer,
                    p_numFmtId  pls_integer := null,
                    p_fontId    pls_integer := null,
                    p_fillId    pls_integer := null,
                    p_borderId  pls_integer := null,
                    p_alignment tp_alignment := null,
                    p_sheet     pls_integer := null) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).row_fmts(p_row).numFmtId := p_numFmtId;
    workbook.sheets(t_sheet).row_fmts(p_row).fontId := p_fontId;
    workbook.sheets(t_sheet).row_fmts(p_row).fillId := p_fillId;
    workbook.sheets(t_sheet).row_fmts(p_row).borderId := p_borderId;
    workbook.sheets(t_sheet).row_fmts(p_row).alignment := p_alignment;
  end;
  --
  procedure freeze_rows(p_nr_rows pls_integer := 1,
                        p_sheet   pls_integer := null) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).freeze_cols := null;
    workbook.sheets(t_sheet).freeze_rows := p_nr_rows;
  end;
  --
  procedure freeze_cols(p_nr_cols pls_integer := 1,
                        p_sheet   pls_integer := null) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).freeze_rows := null;
    workbook.sheets(t_sheet).freeze_cols := p_nr_cols;
  end;
  --
  procedure freeze_pane(p_col   pls_integer,
                        p_row   pls_integer,
                        p_sheet pls_integer := null) is
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    workbook.sheets(t_sheet).freeze_rows := p_row;
    workbook.sheets(t_sheet).freeze_cols := p_col;
  end;
  --
  procedure set_autofilter(p_column_start pls_integer := null,
                           p_column_end   pls_integer := null,
                           p_row_start    pls_integer := null,
                           p_row_end      pls_integer := null,
                           p_sheet        pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := 1;
    workbook.sheets(t_sheet).autofilters(t_ind).column_start := p_column_start;
    workbook.sheets(t_sheet).autofilters(t_ind).column_end := p_column_end;
    workbook.sheets(t_sheet).autofilters(t_ind).row_start := p_row_start;
    workbook.sheets(t_sheet).autofilters(t_ind).row_end := p_row_end;
    defined_name(p_column_start,
                 p_row_start,
                 p_column_end,
                 p_row_end,
                 '_xlnm._FilterDatabase',
                 t_sheet,
                 t_sheet - 1);
  end;

  procedure set_tablestyle(p_column_start  pls_integer := null,
                           p_column_end    pls_integer := null,
                           p_row_start     pls_integer := null,
                           p_row_end       pls_integer := null,
                           p_tb_style_name varchar2,
                           p_sheet         pls_integer := null) is
    t_ind   pls_integer;
    t_sheet pls_integer := nvl(p_sheet, workbook.sheets.count());
  begin
    t_ind := workbook.sheets(t_sheet).tablestyles.count() + 1;
    workbook.sheets(t_sheet).tablestyles(t_ind).column_start := p_column_start;
    workbook.sheets(t_sheet).tablestyles(t_ind).column_end := p_column_end;
    workbook.sheets(t_sheet).tablestyles(t_ind).row_start := p_row_start;
    workbook.sheets(t_sheet).tablestyles(t_ind).row_end := p_row_end;
    workbook.sheets(t_sheet).tablestyles(t_ind).tb_style_name := p_tb_style_name;
  end;

  procedure set_autofitcolumn(
    p_value varchar2,
    p_col   pls_integer,
    p_sheet pls_integer
  ) is
    t_nr_chr number;
    tmp_width number;
    t_width number;
    t_pixels number;
  begin
    t_nr_chr := length(p_value) + 4;-- Foi adicionado mais 4 caracteres para
                                    -- compensar o "bot�o com seta para baixo"
                                    -- do combobox do filtro
    tmp_width := trunc((t_nr_chr * 7 + 5) / 7 * 256) / 256;


    t_width := 0;
    if workbook.sheets(p_sheet).autofitcols.exists(p_col) then
      t_width := nvl(workbook.sheets(p_sheet).autofitcols(p_col).width,0);
    end if;

    if tmp_width > t_width then
      workbook.sheets(p_sheet).autofitcols(p_col).width := tmp_width;
    end if;
  end;
  --
  /*
    procedure add1xml
      ( p_excel in out nocopy blob
      , p_filename varchar2
      , p_xml clob
      )
    is
      t_tmp blob;
      c_step constant number := 24396;
    begin
      dbms_lob.createtemporary( t_tmp, true );
      for i in 0 .. trunc( length( p_xml ) / c_step )
      loop
        dbms_lob.append( t_tmp, utl_i18n.string_to_raw( substr( p_xml, i * c_step + 1, c_step ), 'AL32UTF8' ) );
      end loop;
      add1file( p_excel, p_filename, t_tmp );
      dbms_lob.freetemporary( t_tmp );
    end;
  */
  --
  procedure add1xml(p_excel    in out nocopy blob,
                    p_filename varchar2,
                    p_xml      clob) is
    t_tmp        blob;
    dest_offset  integer := 1;
    src_offset   integer := 1;
    lang_context integer;
    warning      integer;
  begin
    lang_context := dbms_lob.DEFAULT_LANG_CTX;
    dbms_lob.createtemporary(t_tmp, true);
    dbms_lob.converttoblob(t_tmp,
                           p_xml,
                           dbms_lob.lobmaxsize,
                           dest_offset,
                           src_offset,
                           nls_charset_id('AL32UTF8'),
                           lang_context,
                           warning);
    add1file(p_excel, p_filename, t_tmp);
    dbms_lob.freetemporary(t_tmp);
  end;
  --
  function finish return blob is
    t_excel   blob;
    t_xxx     clob;
    t_tmp     varchar2(32767 char);
    t_str     varchar2(32767 char);
    t_c       number;
    t_h       number;
    t_w       number;
    t_cw      number;
    t_cell    varchar2(1000 char);
    t_row_ind pls_integer;
    t_col_min pls_integer;
    t_col_max pls_integer;
    t_col_ind pls_integer;
    t_len     pls_integer;
    ts        timestamp := systimestamp;
    t_cnt_tablestyle pls_integer := 0;
    t_tmp_tablestyle pls_integer := 0;
    lc_rows tp_rows;

  begin

    for x in 1..workbook.sheets.count() loop
      if workbook.sheets(x).tablestyles.count() > 0 then
        t_cnt_tablestyle := t_cnt_tablestyle + workbook.sheets(x).tablestyles.count();
      end if;
    end loop;
    t_tmp_tablestyle := t_cnt_tablestyle;

    dbms_lob.createtemporary(t_excel, true);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    for s in 1 .. workbook.sheets.count() loop
      t_xxx := t_xxx || '
<Override PartName="/xl/worksheets/sheet' || s ||
               '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    end loop;
    t_xxx := t_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
   if t_cnt_tablestyle > 0 then
      for x in 1..t_cnt_tablestyle loop
        t_xxx := t_xxx ||
        '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"' ||
        ' PartName="/xl/tables/table' || x || '.xml"/>';
      end loop;
    end if;

    if has_formula then
      t_xxx := t_xxx ||
      '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml" PartName="/xl/calcChain.xml"/>';
    end if;

    t_xxx := t_xxx || '
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    for s in 1 .. workbook.sheets.count() loop
      if workbook.sheets(s).comments.count() > 0 then
        t_xxx := t_xxx || '
<Override PartName="/xl/comments' || s ||
                 '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>';
      end if;
    end loop;

    t_xxx := t_xxx || '
</Types>';
    add1xml(t_excel, '[Content_Types].xml', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context('userenv', 'os_user') ||
             '</dc:creator>
<cp:lastModifiedBy>' || sys_context('userenv', 'os_user') ||
             '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' ||
             to_char(current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM') ||
             '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' ||
             to_char(current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM') ||
             '</dcterms:modified>
</cp:coreProperties>';
    add1xml(t_excel, 'docProps/core.xml', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count() || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count() ||
             '" baseType="lpstr">';
    for s in 1 .. workbook.sheets.count() loop
      t_xxx := t_xxx || '
<vt:lpstr>' || workbook.sheets(s).name || '</vt:lpstr>';
    end loop;
    t_xxx := t_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    add1xml(t_excel, 'docProps/app.xml', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    add1xml(t_excel, '_rels/.rels', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    if workbook.numFmts.count() > 0 then
      t_xxx := t_xxx || '<numFmts count="' || workbook.numFmts.count() || '">';
      for n in 1 .. workbook.numFmts.count() loop
        t_xxx := t_xxx || '<numFmt numFmtId="' || workbook.numFmts(n)
                .numFmtId || '" formatCode="' || workbook.numFmts(n)
                .formatCode || '"/>';
      end loop;
      t_xxx := t_xxx || '</numFmts>';
    end if;
    t_xxx := t_xxx || '<fonts count="' || workbook.fonts.count() ||
             '" x14ac:knownFonts="1">';
    for f in 0 .. workbook.fonts.count() - 1 loop
      t_xxx := t_xxx || '<font>' || case
                 when workbook.fonts(f).bold then
                  '<b/>'
               end || case
                 when workbook.fonts(f).italic then
                  '<i/>'
               end || case
                 when workbook.fonts(f).underline then
                  '<u/>'
               end || '<sz val="' ||
               to_char(workbook.fonts(f).fontsize,
                       'TM9',
                       'NLS_NUMERIC_CHARACTERS=.,') || '"/>
<color ' || case
                 when workbook.fonts(f).rgb is not null then
                  'rgb="' || workbook.fonts(f).rgb
                 else
                  'theme="' || workbook.fonts(f).theme
               end || '"/>
<name val="' || workbook.fonts(f).name || '"/>
<family val="' || workbook.fonts(f).family || '"/>
<scheme val="none"/>
</font>';
    end loop;
    t_xxx := t_xxx || '</fonts>
<fills count="' || workbook.fills.count() || '">';
    for f in 0 .. workbook.fills.count() - 1 loop
      t_xxx := t_xxx || '<fill><patternFill patternType="' || workbook.fills(f)
              .patternType || '">' || case
                 when workbook.fills(f).fgRGB is not null then
                  '<fgColor rgb="' || workbook.fills(f).fgRGB || '"/>'
               end || '</patternFill></fill>';
    end loop;
    t_xxx := t_xxx || '</fills>
<borders count="' || workbook.borders.count() || '">';
    for b in 0 .. workbook.borders.count() - 1 loop
      t_xxx := t_xxx || '<border>' || case
                 when workbook.borders(b).left is null then
                  '<left/>'
                 else
                  '<left style="' || workbook.borders(b).left || '"/>'
               end || case
                 when workbook.borders(b).right is null then
                  '<right/>'
                 else
                  '<right style="' || workbook.borders(b).right || '"/>'
               end || case
                 when workbook.borders(b).top is null then
                  '<top/>'
                 else
                  '<top style="' || workbook.borders(b).top || '"/>'
               end || case
                 when workbook.borders(b).bottom is null then
                  '<bottom/>'
                 else
                  '<bottom style="' || workbook.borders(b).bottom || '"/>'
               end || '</border>';
    end loop;
    t_xxx := t_xxx || '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || (workbook.cellXfs.count() + 1) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    for x in 1 .. workbook.cellXfs.count() loop
      t_xxx := t_xxx || '<xf numFmtId="' || workbook.cellXfs(x).numFmtId ||
               '" fontId="' || workbook.cellXfs(x).fontId || '" fillId="' || workbook.cellXfs(x)
              .fillId || '" borderId="' || workbook.cellXfs(x).borderId || '">';
      if (workbook.cellXfs(x).alignment.horizontal is not null or workbook.cellXfs(x)
         .alignment.vertical is not null or workbook.cellXfs(x)
         .alignment.wrapText) then
        t_xxx := t_xxx || '<alignment' || case
                   when workbook.cellXfs(x).alignment.horizontal is not null then
                    ' horizontal="' || workbook.cellXfs(x).alignment.horizontal || '"'
                 end || case
                   when workbook.cellXfs(x).alignment.vertical is not null then
                    ' vertical="' || workbook.cellXfs(x).alignment.vertical || '"'
                 end || case
                   when workbook.cellXfs(x).alignment.wrapText then
                    ' wrapText="true"'
                 end || '/>';
      end if;
      t_xxx := t_xxx || '</xf>';
    end loop;
    t_xxx := t_xxx || '</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>';
    add1xml(t_excel, 'xl/styles.xml', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="false" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    for s in 1 .. workbook.sheets.count() loop
      t_xxx := t_xxx || '
<sheet name="' || workbook.sheets(s).name || '" sheetId="' || s ||
               '" r:id="rId' || (9 + s) || '"/>';
    end loop;
    t_xxx := t_xxx || '</sheets>';
    if workbook.defined_names.count() > 0 then
      t_xxx := t_xxx || '<definedNames>';
      for s in 1 .. workbook.defined_names.count() loop
        t_xxx := t_xxx || '
<definedName name="' || workbook.defined_names(s).name || '"' || case
                   when workbook.defined_names(s).sheet is not null then
                    ' localSheetId="' || to_char(workbook.defined_names(s).sheet) || '"'
                 end || '>' || workbook.defined_names(s).ref ||
                 '</definedName>';
      end loop;
      t_xxx := t_xxx || '</definedNames>';
    end if;
    t_xxx := t_xxx || '<calcPr calcId="144525"/></workbook>';
    add1xml(t_excel, 'xl/workbook.xml', t_xxx);
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';
    add1xml(t_excel, 'xl/theme/theme1.xml', t_xxx);

    for s in 1 .. workbook.sheets.count() loop
      t_col_min := 16384;
      t_col_max := 1;
      /*t_row_ind := workbook.sheets(s).rows.first();
      while t_row_ind is not null loop
        t_col_min := least(t_col_min,
                           workbook.sheets(s).rows(t_row_ind).first());
        t_col_max := greatest(t_col_max,
                              workbook.sheets(s).rows(t_row_ind).last());
        t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
      end loop;*/
      lc_rows := workbook.sheets( s ).rows;
      t_row_ind := lc_rows.first();
      loop
        exit when t_row_ind is null;
        t_col_min := least( t_col_min, lc_rows( t_row_ind ).first () );
        t_col_max := greatest( t_col_max, lc_rows( t_row_ind ).last() );
        t_row_ind := lc_rows.next( t_row_ind );
      end loop;

      t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
<dimension ref="' || alfan_col(t_col_min) || workbook.sheets(s)
              .rows.first() || ':' || alfan_col(t_col_max) || workbook.sheets(s)
              .rows.last() || '"/>
<sheetViews>
<sheetView' ||
case
  when workbook.sheets(s).showGridLines = true then
    ' showGridLines="1"'
  else
    ' showGridLines="0"'
end ||
case
                 when s = 1 then
                  ' tabSelected="1"'
               end || ' workbookViewId="0">';
      if workbook.sheets(s)
       .freeze_rows > 0 and workbook.sheets(s).freeze_cols > 0 then
        t_xxx := t_xxx || ('<pane xSplit="' || workbook.sheets(s)
                 .freeze_cols || '" ' || 'ySplit="' || workbook.sheets(s)
                 .freeze_rows || '" ' || 'topLeftCell="' ||
                 alfan_col(workbook.sheets(s).freeze_cols + 1) ||
                 (workbook.sheets(s).freeze_rows + 1) || '" ' ||
                 'activePane="bottomLeft" state="frozen"/>');
      else
        if workbook.sheets(s).freeze_rows > 0 then
          t_xxx := t_xxx || '<pane ySplit="' || workbook.sheets(s)
                  .freeze_rows || '" topLeftCell="A' ||
                   (workbook.sheets(s).freeze_rows + 1) ||
                   '" activePane="bottomLeft" state="frozen"/>';
        end if;
        if workbook.sheets(s).freeze_cols > 0 then
          t_xxx := t_xxx || '<pane xSplit="' || workbook.sheets(s)
                  .freeze_cols || '" topLeftCell="' ||
                   alfan_col(workbook.sheets(s).freeze_cols + 1) ||
                   '1" activePane="bottomLeft" state="frozen"/>';
        end if;
      end if;
      t_xxx := t_xxx ||
               '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>';
      if workbook.sheets(s).widths.count() > 0 then
        t_xxx     := t_xxx || '<cols>';
        t_col_ind := workbook.sheets(s).widths.first();
        while t_col_ind is not null loop
          t_xxx     := t_xxx || '<col min="' || t_col_ind || '" max="' ||
                       t_col_ind || '" width="' ||
                       to_char(workbook.sheets(s).widths(t_col_ind),
                               'TM9',
                               'NLS_NUMERIC_CHARACTERS=.,') ||
                       '" customWidth="1"/>';
          t_col_ind := workbook.sheets(s).widths.next(t_col_ind);
        end loop;
        t_xxx := t_xxx || '</cols>';
      end if;
      t_xxx     := t_xxx || '<sheetData>';
      --t_row_ind := workbook.sheets(s).rows.first();
      t_row_ind := lc_rows.first();
      t_tmp     := null;
      --while t_row_ind is not null loop
      while t_row_ind is not null loop
        if workbook.sheets(s).heights.exists(t_row_ind) then
          t_tmp     := t_tmp || '<row r="' || t_row_ind ||
                       '" customHeight="1" ht="' ||
                       workbook.sheets(s).heights(t_row_ind) ||
                       '" spans="' || t_col_min || ':' || t_col_max || '">';
        else
          t_tmp     := t_tmp || '<row r="' || t_row_ind || '" spans="' ||
                       t_col_min || ':' || t_col_max || '">';
        end if;

        t_len     := length(t_tmp);
        --t_col_ind := workbook.sheets(s).rows(t_row_ind).first();
        t_col_ind := lc_rows( t_row_ind ).first();

        while t_col_ind is not null loop
          if lc_rows(t_row_ind)(t_col_ind).formula then
            t_cell := '<c r="' || alfan_col(t_col_ind) || t_row_ind || '"' || ' ' ||
                      workbook.sheets(s).rows(t_row_ind)(t_col_ind).style || '><f>' ||
                      lc_rows(t_row_ind)(t_col_ind).value2
                      || '</f></c>';
          else
            t_cell := '<c r="' || alfan_col(t_col_ind) || t_row_ind || '"' || ' ' || workbook.sheets(s)
                     .rows(t_row_ind)(t_col_ind)
                     .style || '><v>' ||
                      --to_char(workbook.sheets(s).rows(t_row_ind)(t_col_ind)
                      to_char( lc_rows( t_row_ind )( t_col_ind )
                              .value,
                              'TM9',
                              'NLS_NUMERIC_CHARACTERS=.,') || '</v></c>';
          end if;
          if t_len > 32000 then
            dbms_lob.writeappend(t_xxx, t_len, t_tmp);
            t_tmp := null;
            t_len := 0;
          end if;
          t_tmp     := t_tmp || t_cell;
          t_len     := t_len + length(t_cell);
          --t_col_ind := workbook.sheets(s).rows(t_row_ind).next(t_col_ind);
          t_col_ind := lc_rows( t_row_ind ).next( t_col_ind );
        end loop;
        t_tmp     := t_tmp || '</row>';
        --t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
        t_row_ind := lc_rows.next( t_row_ind );
      end loop;
      t_tmp := t_tmp || '</sheetData>';
      t_len := length(t_tmp);
      dbms_lob.writeappend(t_xxx, t_len, t_tmp);
      for a in 1 .. workbook.sheets(s).autofilters.count() loop
        declare
          col_ini_alfa varchar2(3);
          col_fim_alfa varchar2(3);
          lin_ini_alfa varchar2(8);
          lin_fim_alfa varchar2(8);
        begin
          col_ini_alfa := alfan_col(nvl(workbook.sheets(s).autofilters(a)
                                        .column_start,
                                        t_col_min));
          col_fim_alfa := alfan_col(coalesce(workbook.sheets(s).autofilters(a)
                                             .column_end,
                                             workbook.sheets(s).autofilters(a)
                                             .column_start,
                                             t_col_max));
          lin_ini_alfa := nvl(workbook.sheets(s).autofilters(a).row_start,
                              workbook.sheets(s).rows.first());
          lin_fim_alfa := nvl(workbook.sheets(s).autofilters(a).row_end,
                              workbook.sheets(s).rows.last());

          t_xxx := t_xxx || '<autoFilter ref="' || col_ini_alfa ||
                   lin_ini_alfa || ':' || col_fim_alfa || lin_fim_alfa ||
                   '"/>';

          /*t_xxx := t_xxx || '<autoFilter ref="' ||
          alfan_col( nvl( workbook.sheets( s ).autofilters( a ).column_start, t_col_min ) ) ||
          nvl( workbook.sheets( s ).autofilters( a ).row_start, workbook.sheets( s ).rows.first() ) || ':' ||
          alfan_col( coalesce( workbook.sheets( s ).autofilters( a ).column_end, workbook.sheets( s ).autofilters( a ).column_start, t_col_max ) ) ||
          nvl( workbook.sheets( s ).autofilters( a ).row_end, workbook.sheets( s ).rows.last() ) || '"/>';*/
        end;
      end loop;
      if workbook.sheets(s).mergecells.count() > 0 then
        t_xxx := t_xxx || '<mergeCells count="' ||
                 to_char(workbook.sheets(s).mergecells.count()) || '">';
        for m in 1 .. workbook.sheets(s).mergecells.count() loop
          t_xxx := t_xxx || '<mergeCell ref="' || workbook.sheets(s)
                  .mergecells(m) || '"/>';
        end loop;
        t_xxx := t_xxx || '</mergeCells>';
      end if;
      --
      if workbook.sheets(s).validations.count() > 0 then
        t_xxx := t_xxx || '<dataValidations count="' ||
                 to_char(workbook.sheets(s).validations.count()) || '">';
        for m in 1 .. workbook.sheets(s).validations.count() loop
          t_xxx := t_xxx || '<dataValidation' || ' type="' || workbook.sheets(s).validations(m).type || '"' ||
                   ' errorStyle="' || workbook.sheets(s).validations(m)
                  .errorstyle || '"' || ' allowBlank="' || case
                     when nvl(workbook.sheets(s).validations(m).allowBlank, true) then
                      '1'
                     else
                      '0'
                   end || '"' || ' sqref="' || workbook.sheets(s).validations(m)
                  .sqref || '"';
          if workbook.sheets(s).validations(m).prompt is not null then
            t_xxx := t_xxx || ' showInputMessage="1" prompt="' || workbook.sheets(s).validations(m)
                    .prompt || '"';
            if workbook.sheets(s).validations(m).title is not null then
              t_xxx := t_xxx || ' promptTitle="' || workbook.sheets(s).validations(m)
                      .title || '"';
            end if;
          end if;
          if workbook.sheets(s).validations(m).showerrormessage then
            t_xxx := t_xxx || ' showErrorMessage="1"';
            if workbook.sheets(s).validations(m).error_title is not null then
              t_xxx := t_xxx || ' errorTitle="' || workbook.sheets(s).validations(m)
                      .error_title || '"';
            end if;
            if workbook.sheets(s).validations(m).error_txt is not null then
              t_xxx := t_xxx || ' error="' || workbook.sheets(s).validations(m)
                      .error_txt || '"';
            end if;
          end if;
          t_xxx := t_xxx || '>';
          if workbook.sheets(s).validations(m).formula1 is not null then
            t_xxx := t_xxx || '<formula1>' || workbook.sheets(s).validations(m)
                    .formula1 || '</formula1>';
          end if;
          if workbook.sheets(s).validations(m).formula2 is not null then
            t_xxx := t_xxx || '<formula2>' || workbook.sheets(s).validations(m)
                    .formula2 || '</formula2>';
          end if;
          t_xxx := t_xxx || '</dataValidation>';
        end loop;
        t_xxx := t_xxx || '</dataValidations>';
      end if;
      --
      if workbook.sheets(s).hyperlinks.count() > 0 then
        t_xxx := t_xxx || '<hyperlinks>';
        for h in 1 .. workbook.sheets(s).hyperlinks.count() loop
          t_xxx := t_xxx || '<hyperlink ref="' || workbook.sheets(s).hyperlinks(h).cell ||
                   '" r:id="rId' || h || '"/>';
        end loop;
        t_xxx := t_xxx || '</hyperlinks>';
      end if;
      t_xxx := t_xxx ||
               '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      if workbook.sheets(s).comments.count() > 0 then
        t_xxx := t_xxx || '<legacyDrawing r:id="rId' ||
                 (workbook.sheets(s).hyperlinks.count() + 1) || '"/>';
      end if;

      if workbook.sheets(s).tablestyles.count() > 0 then
        --begin
        t_xxx :=
          t_xxx || '<tableParts count="' ||
          to_char(workbook.sheets(s).tablestyles.count()) ||
          '">';
        /*exception
          when others then
            dbms_output.put_line(sqlerrm);
            dbms_output.put_line('tb_style_count: '||workbook.sheets(s).tablestyles.count());
        end;*/
        for i in 1..workbook.sheets(s).tablestyles.count() loop
          t_xxx := t_xxx ||
                   '<tablePart r:id="rId' || to_char(i) || '"/>';
        end loop;
        t_xxx := t_xxx || '</tableParts>';
      end if;

      --
      t_xxx := t_xxx || '</worksheet>';
      add1xml(t_excel, 'xl/worksheets/sheet' || s || '.xml', t_xxx);

      if workbook.sheets(s).hyperlinks.count() > 0 or
         workbook.sheets(s).comments.count() > 0 or
         workbook.sheets(s).tablestyles.count() > 0
      then
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        if workbook.sheets(s).comments.count() > 0 then
          t_xxx := t_xxx || '<Relationship Id="rId' ||
                   (workbook.sheets(s).hyperlinks.count() + 2) ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' || s ||
                   '.xml"/>';
          t_xxx := t_xxx || '<Relationship Id="rId' ||
                   (workbook.sheets(s).hyperlinks.count() + 1) ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || s ||
                   '.vml"/>';
        end if;
        for h in 1 .. workbook.sheets(s).hyperlinks.count() loop
          t_xxx := t_xxx || '<Relationship Id="rId' || h ||
                   '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets(s).hyperlinks(h).url ||
                   '" TargetMode="External"/>';
        end loop;
        for i in reverse 1..workbook.sheets(s).tablestyles.count() loop
          t_xxx := t_xxx ||
          '<Relationship Id="rId' || i || '"' ||
          ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"' ||
          ' Target="../tables/table' || t_tmp_tablestyle || '.xml"/>';
          t_tmp_tablestyle := t_tmp_tablestyle - 1;
        end loop;
        t_xxx := t_xxx || '</Relationships>';
        add1xml(t_excel,
                'xl/worksheets/_rels/sheet' || s || '.xml.rels',
                t_xxx);
      end if;
      --
      if workbook.sheets(s).comments.count() > 0 then
        declare
          cnt        pls_integer;
          author_ind tp_author;
          --          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        begin
          authors.delete();
          for c in 1 .. workbook.sheets(s).comments.count() loop
            authors(workbook.sheets(s).comments(c).author) := 0;
          end loop;
          t_xxx      := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
          cnt        := 0;
          author_ind := authors.first();
          while author_ind is not null or
                authors.next(author_ind) is not null loop
            authors(author_ind) := cnt;
            t_xxx := t_xxx || '<author>' || author_ind || '</author>';
            cnt := cnt + 1;
            author_ind := authors.next(author_ind);
          end loop;
        end;
        t_xxx := t_xxx || '</authors><commentList>';
        for c in 1 .. workbook.sheets(s).comments.count() loop
          t_xxx := t_xxx || '<comment ref="' ||
                   alfan_col(workbook.sheets(s).comments(c).column) ||
                   to_char(workbook.sheets(s).comments(c)
                           .row || '" authorId="' ||
                            authors(workbook.sheets(s).comments(c).author)) || '">
<text>';
          if workbook.sheets(s).comments(c).author is not null then
            t_xxx := t_xxx ||
                     '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' || workbook.sheets(s).comments(c)
                    .author || ':</t></r>';
          end if;
          t_xxx := t_xxx ||
                   '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' || case
                     when workbook.sheets(s).comments(c).author is not null then
                      '
'
                   end || workbook.sheets(s).comments(c).text ||
                   '</t></r></text></comment>';
        end loop;
        t_xxx := t_xxx || '</commentList></comments>';
        add1xml(t_excel, 'xl/comments' || s || '.xml', t_xxx);
        t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
        for c in 1 .. workbook.sheets(s).comments.count() loop
          t_xxx := t_xxx || '<v:shape id="_x0000_s' || to_char(c) ||
                   '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' ||
                   to_char(c) ||
                   ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>';
          t_w   := workbook.sheets(s).comments(c).width;
          t_c   := 1;
          loop
            if workbook.sheets(s)
             .widths.exists(workbook.sheets(s).comments(c).column + t_c) then
              t_cw := 256 * workbook.sheets(s)
                     .widths(workbook.sheets(s).comments(c).column + t_c);
              t_cw := trunc((t_cw + 18) / 256 * 7); -- assume default 11 point Calibri
            else
              t_cw := 64;
            end if;
            exit when t_w < t_cw;
            t_c := t_c + 1;
            t_w := t_w - t_cw;
          end loop;
          t_h   := workbook.sheets(s).comments(c).height;
          t_xxx := t_xxx || to_char('<x:Anchor>' || workbook.sheets(s).comments(c)
                                    .column || ',15,' || workbook.sheets(s).comments(c).row ||
                                    ',30,' || (workbook.sheets(s).comments(c)
                                    .column + t_c - 1) || ',' || round(t_w) || ',' ||
                                    (workbook.sheets(s).comments(c)
                                    .row + 1 + trunc(t_h / 20)) || ',' ||
                                    mod(t_h, 20) || '</x:Anchor>');
          t_xxx := t_xxx ||
                   to_char('<x:AutoFill>False</x:AutoFill><x:Row>' ||
                           (workbook.sheets(s).comments(c).row - 1) ||
                           '</x:Row><x:Column>' ||
                           (workbook.sheets(s).comments(c).column - 1) ||
                           '</x:Column></x:ClientData></v:shape>');
        end loop;
        t_xxx := t_xxx || '</xml>';
        add1xml(t_excel, 'xl/drawings/vmlDrawing' || s || '.vml', t_xxx);
      end if;
      --
    end loop;

    for s in 1..workbook.sheets.count() loop

      if workbook.sheets(s).tablestyles.count() > 0 then
        for a in 1 .. workbook.sheets(s).tablestyles.count() loop
          declare
            col_ini_alfa varchar2(3);
            col_fim_alfa varchar2(3);
            lin_ini_alfa varchar2(8);
            lin_fim_alfa varchar2(8);
            cnt_col      pls_integer;
            col_start    pls_integer;
            col_end      pls_integer;
            row_start    pls_integer;
            row_end      pls_integer;
            tmp_id       pls_integer;
          begin
            col_start := workbook.sheets(s).tablestyles(a).column_start;
            col_end   := workbook.sheets(s).tablestyles(a).column_end;
            row_start := workbook.sheets(s).tablestyles(a).row_start;
            row_end := workbook.sheets(s).tablestyles(a).row_end;

            t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            col_ini_alfa :=
              alfan_col(col_start);
            col_fim_alfa :=
              alfan_col(coalesce(col_end,
                                 col_start));

            if row_start = row_end then
              row_end := row_end + 1;
            end if;

            lin_ini_alfa := row_start;
            lin_fim_alfa := row_end;

            if workbook.sheets(s).rows.exists(row_end + 1) and has_formula then
              t_xxx := t_xxx ||
                       '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' ||
                       ' id="' || t_cnt_tablestyle || '"' ||
                       ' name="Tabela' || t_cnt_tablestyle || '"' ||
                       ' displayName="Tabela' || t_cnt_tablestyle || '"' ||
                       ' ref="' || col_ini_alfa || lin_ini_alfa || ':' ||
                       col_fim_alfa || to_char(row_end+1) || '"' ||
                       ' totalsRowCount="1">';
            else
              t_xxx := t_xxx ||
                       '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' ||
                       ' id="' || t_cnt_tablestyle || '"' ||
                       ' name="Tabela' || t_cnt_tablestyle || '"' ||
                       ' displayName="Tabela' || t_cnt_tablestyle || '"' ||
                       ' ref="' || col_ini_alfa || lin_ini_alfa || ':' ||
                       col_fim_alfa || lin_fim_alfa || '"' ||
                       ' totalsRowShown="0">';
            end if;

            t_xxx := t_xxx || '<autoFilter ref="' || col_ini_alfa ||
                     lin_ini_alfa || ':' || col_fim_alfa || lin_fim_alfa ||
                     '"/>';

            cnt_col := col_end - col_start + 1;
            t_xxx := t_xxx || '<tableColumns count="' || cnt_col || '">';

            tmp_id := 1;

            for col in col_start..col_end loop

              dbms_output.put_line(row_end);
              if workbook.sheets(s).rows.exists(row_end + 1) and
                 workbook.sheets(s).rows(row_end+1).exists(col) then
                if workbook.sheets(s).rows(row_end+1)(col).formula then
                  t_xxx :=
                    t_xxx || '<tableColumn id="' || tmp_id || '" name="' ||
                    to_char(workbook.sheets(s).rows(row_start)(col).value2) ||
                    '" totalsRowFunction="custom">' ||
                    '<totalsRowFormula>' ||
                    workbook.sheets(s).rows(row_end+1)(col).value2 ||
                    '</totalsRowFormula>' ||
                    '</tableColumn>';
                else
                  t_xxx :=
                    t_xxx || '<tableColumn id="' || tmp_id || '" name="' ||
                    to_char(workbook.sheets(s).rows(row_start)(col).value2) ||
                    '"/>';
                end if;
              else
                t_xxx :=
                  t_xxx || '<tableColumn id="' || tmp_id || '" name="' ||
                  to_char(workbook.sheets(s).rows(row_start)(col).value2) ||
                  '"/>';
              end if;
              tmp_id := tmp_id + 1;
            end loop;

            t_xxx := t_xxx || '</tableColumns>';
            t_xxx :=
              t_xxx || '<tableStyleInfo name="' ||
              workbook.sheets(s).tablestyles(a).tb_style_name ||
              '" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>';
            t_xxx := t_xxx || '</table>';

            add1xml(t_excel,
                    'xl/tables/table' || t_cnt_tablestyle || '.xml',
                    t_xxx);
            t_cnt_tablestyle := t_cnt_tablestyle - 1;
          end;
        end loop;
      end if;
    end loop;

    if has_formula then
      t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ||
               '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

      /*xxxxxxxxxxxxxxxxxxxxxxx*/
      t_cell := '';
      for s in 1 .. workbook.sheets.count() loop
        lc_rows := workbook.sheets( s ).rows;

        t_row_ind := lc_rows.first();
        t_tmp     := null;
        while t_row_ind is not null loop
          t_len     := length(t_tmp);
          t_col_ind := lc_rows( t_row_ind ).first();

          while t_col_ind is not null loop
            if lc_rows(t_row_ind)(t_col_ind).formula then
              t_cell := '<c r="' ||
                        alfan_col(t_col_ind) || t_row_ind||
                        '" i="1" l="1"/>';
            end if;
            if t_len > 32000 then
              dbms_lob.writeappend(t_xxx, t_len, t_tmp);
              t_tmp := null;
              t_len := 0;
            end if;
            t_tmp     := t_tmp || t_cell;
            t_len     := t_len + length(t_cell);
            --t_col_ind := workbook.sheets(s).rows(t_row_ind).next(t_col_ind);
            t_col_ind := lc_rows( t_row_ind ).next( t_col_ind );
          end loop;
          --t_row_ind := workbook.sheets(s).rows.next(t_row_ind);
          t_row_ind := lc_rows.next( t_row_ind );
        end loop;
        t_len := length(t_tmp);
        dbms_lob.writeappend(t_xxx, t_len, t_tmp);
      end loop;
      /*xxxxxxxxxxxxxxxxxxxxxxx*/

      t_xxx := t_xxx || '</calcChain>';
      add1xml(t_excel, 'xl/calcChain.xml', t_xxx);
    end if;

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';

    if has_formula then
      t_xxx := t_xxx ||
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>';
    end if;

    for s in 1 .. workbook.sheets.count() loop
      t_xxx := t_xxx || '
<Relationship Id="rId' || (9 + s) ||
               '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s ||
               '.xml"/>';
    end loop;

    t_xxx := t_xxx || '</Relationships>';
    add1xml(t_excel, 'xl/_rels/workbook.xml.rels', t_xxx);

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' ||
             workbook.str_cnt || '" uniqueCount="' ||
             workbook.strings.count() || '">';
    t_tmp := null;
    for i in 0 .. workbook.str_ind.count() - 1 loop
      t_str := '<si><t>' ||
               dbms_xmlgen.convert(substr(workbook.str_ind(i), 1, 32000)) ||
               '</t></si>';
      if length(t_tmp) + length(t_str) > 32000 then
        t_xxx := t_xxx || t_tmp;
        t_tmp := null;
      end if;
      t_tmp := t_tmp || t_str;
    end loop;
    t_xxx := t_xxx || t_tmp || '</sst>';
    add1xml(t_excel, 'xl/sharedStrings.xml', t_xxx);
    finish_zip(t_excel);
    clear_workbook;
    return t_excel;
  end;
  --
  procedure save(p_directory varchar2, p_filename varchar2) is
  begin
    blob2file(finish, p_directory, p_filename);
  end;
  --
  procedure query2sheet(p_sql            VARCHAR2,
                        p_column_headers boolean := true,
                        p_directory      varchar2 := null,
                        p_filename       varchar2 := null,
                        p_sheet          pls_integer := NULL) is
    t_sheet     pls_integer;
    t_c         integer;
    t_col_cnt   integer;
    t_desc_tab  dbms_sql.desc_tab2;
    d_tab       dbms_sql.date_table;
    n_tab       dbms_sql.number_table;
    v_tab       dbms_sql.varchar2_table;
    t_bulk_size pls_integer := 200;
    t_r         integer;
    t_cur_row   pls_integer;
  begin
    if p_sheet is null then
      new_sheet;
    end if;
    t_c := dbms_sql.open_cursor;
    dbms_sql.parse(t_c, p_sql, dbms_sql.native);
    dbms_sql.describe_columns2(t_c, t_col_cnt, t_desc_tab);
    for c in 1 .. t_col_cnt loop
      if p_column_headers then
        cell(c, 1, t_desc_tab(c).col_name, p_sheet => t_sheet);
      end if;
      --      dbms_output.put_line( t_desc_tab( c ).col_name || ' ' || t_desc_tab( c ).col_type );
      case
        when t_desc_tab(c).col_type in (2, 100, 101) then
          dbms_sql.define_array(t_c, c, n_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
          dbms_sql.define_array(t_c, c, d_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
          dbms_sql.define_array(t_c, c, v_tab, t_bulk_size, 1);
        else
          null;
      end case;
    end loop;
    --
    t_cur_row := case
                   when p_column_headers then
                    2
                   else
                    1
                 end;
    t_sheet   := nvl(p_sheet, workbook.sheets.count());
    --
    t_r := dbms_sql.execute(t_c);
    loop
      t_r := dbms_sql.fetch_rows(t_c);
      if t_r > 0 then
        for c in 1 .. t_col_cnt loop
          case
            when t_desc_tab(c).col_type in (2, 100, 101) then
              dbms_sql.column_value(t_c, c, n_tab);
              for i in 0 .. t_r - 1 loop
                if n_tab(i + n_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       n_tab(i + n_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              n_tab.delete;
            when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
              dbms_sql.column_value(t_c, c, d_tab);
              for i in 0 .. t_r - 1 loop
                if d_tab(i + d_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       d_tab(i + d_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              d_tab.delete;
            when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
              dbms_sql.column_value(t_c, c, v_tab);
              for i in 0 .. t_r - 1 loop
                if v_tab(i + v_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       v_tab(i + v_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              v_tab.delete;
            else
              null;
          end case;
        end loop;
      end if;
      exit when t_r != t_bulk_size;
      t_cur_row := t_cur_row + t_r;
    end loop;
    dbms_sql.close_cursor(t_c);
    if (p_directory is not null and p_filename is not null) then
      save(p_directory, p_filename);
    end if;
  exception
    when others then
      if dbms_sql.is_open(t_c) then
        dbms_sql.close_cursor(t_c);
      end if;
  end;
  --
  -- MAMJ - query2sheet2.prc � utilizada para iniciar o detalhe da planilha acima da linha 2,
  -- ou seja, quando h� mais de linha de cabe�alho, passando o parametro p_row com
  -- o n�mero da linha a ser iniciada.
  -- A query2sheet.prc inicia o detalhe da planilha na linha 1 ou 2.
  procedure query2sheet2(p_sql            VARCHAR2,
                         p_row            NUMBER,
                         p_column_headers boolean := true,
                         p_directory      varchar2 := null,
                         p_filename       varchar2 := null,
                         p_sheet          pls_integer := NULL) is
    t_sheet     pls_integer;
    t_c         integer;
    t_col_cnt   integer;
    t_desc_tab  dbms_sql.desc_tab2;
    d_tab       dbms_sql.date_table;
    n_tab       dbms_sql.number_table;
    v_tab       dbms_sql.varchar2_table;
    t_bulk_size pls_integer := 200;
    t_r         integer;
    t_cur_row   pls_integer;
  begin
    if p_sheet is null then
      new_sheet;
    end if;
    t_c := dbms_sql.open_cursor;
    dbms_sql.parse(t_c, p_sql, dbms_sql.native);
    dbms_sql.describe_columns2(t_c, t_col_cnt, t_desc_tab);
    for c in 1 .. t_col_cnt loop
      if p_column_headers then
        cell(c, 1, t_desc_tab(c).col_name, p_sheet => t_sheet);
      end if;
      --      dbms_output.put_line( t_desc_tab( c ).col_name || ' ' || t_desc_tab( c ).col_type );
      case
        when t_desc_tab(c).col_type in (2, 100, 101) then
          dbms_sql.define_array(t_c, c, n_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
          dbms_sql.define_array(t_c, c, d_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
          dbms_sql.define_array(t_c, c, v_tab, t_bulk_size, 1);
        else
          null;
      end case;
    end loop;
    --
    -- MAMJ --
    IF p_row IS NULL OR p_row = 0 THEN
      t_cur_row := case
                     when p_column_headers then
                      2
                     else
                      1
                   end;
    ELSE
      t_cur_row := p_row;
    END IF;
    -- fim --

    --t_cur_row := case when p_column_headers then 2 else 1 end;  -- MAMJ
    t_sheet := nvl(p_sheet, workbook.sheets.count());
    --
    t_r := dbms_sql.execute(t_c);
    loop
      t_r := dbms_sql.fetch_rows(t_c);
      if t_r > 0 then
        for c in 1 .. t_col_cnt loop
          case
            when t_desc_tab(c).col_type in (2, 100, 101) then
              dbms_sql.column_value(t_c, c, n_tab);
              for i in 0 .. t_r - 1 loop
                if n_tab(i + n_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       n_tab(i + n_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              n_tab.delete;
            when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
              dbms_sql.column_value(t_c, c, d_tab);
              for i in 0 .. t_r - 1 loop
                if d_tab(i + d_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       d_tab(i + d_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              d_tab.delete;
            when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
              dbms_sql.column_value(t_c, c, v_tab);
              for i in 0 .. t_r - 1 loop
                if v_tab(i + v_tab.first()) is not null then
                  cell(c,
                       t_cur_row + i,
                       v_tab(i + v_tab.first()),
                       p_sheet => t_sheet);
                end if;
              end loop;
              v_tab.delete;
            else
              null;
          end case;
        end loop;
      end if;
      exit when t_r != t_bulk_size;
      t_cur_row := t_cur_row + t_r;
    end loop;
    dbms_sql.close_cursor(t_c);
    if (p_directory is not null and p_filename is not null) then
      save(p_directory, p_filename);
    end if;
  exception
    when others then
      if dbms_sql.is_open(t_c) then
        dbms_sql.close_cursor(t_c);
      end if;
  end;

  -- inicia o detalhe da planilha a partir de determinada linha e coluna
  procedure query2sheet3(p_sql            VARCHAR2,
                         p_row            NUMBER := 1,
                         p_col            number := 1,
                         p_column_headers boolean := true,
                         p_set_autofilter boolean := false,
                         p_tablestyle     varchar2 := null,
                         p_autofitcolumns boolean := false,
                         p_directory      varchar2 := null,
                         p_filename       varchar2 := null,
                         p_sheet          pls_integer := NULL,
                         p_sheet_name     varchar2 := null,
                         p_alignment t_alignment := cast(null as t_alignment)) is
    t_sheet     pls_integer;
    t_c         integer;
    t_col_cnt   integer;
    t_desc_tab  dbms_sql.desc_tab2;
    d_tab       dbms_sql.date_table;
    n_tab       dbms_sql.number_table;
    v_tab       dbms_sql.varchar2_table;
    t_bulk_size pls_integer := 200;
    t_r         integer;
    t_cur_row   pls_integer;
    t_width     number;
    tmp_width   number;
    t_autofitcol tp_autofitcols;
    t_nr_chr     pls_integer;

    type t_cabecalhos is table of pls_integer index by varchar2(255);
    r_cabecalhos t_cabecalhos;
  begin
    /*if p_sheet is null then
      new_sheet;
    end if;*/
    new_sheet2(p_sheet, p_sheet_name);

    t_sheet := nvl(p_sheet, workbook.sheets.count());

    t_c := dbms_sql.open_cursor;
    dbms_sql.parse(t_c, p_sql, dbms_sql.native);
    dbms_sql.describe_columns2(t_c, t_col_cnt, t_desc_tab);

    r_cabecalhos.delete;

    for c in 1 .. t_col_cnt loop
      if p_column_headers then

        /* Quando existir tablestyle, o Excel n�o permite Cabe�alhos com o mesmo
           nome. Dessa forma, caso haja cabe�alhos duplicados, o sistema ir�
           alterar o nome do cabe�alho incluindo um ponto no final
        */

        if r_cabecalhos.exists(to_char(t_desc_tab(c).col_name)) then
          r_cabecalhos(to_char(t_desc_tab(c).col_name)) :=
            r_cabecalhos(to_char(t_desc_tab(c).col_name)) + 1;

          t_desc_tab(c).col_name :=
              to_char(t_desc_tab(c).col_name) || '_' ||
              r_cabecalhos(to_char(t_desc_tab(c).col_name));

        else
          r_cabecalhos(t_desc_tab(c).col_name) := 1;
        end if;
        /**/

        if p_alignment.exists(c) then
          cell((c + p_col - 1),
             p_row,
             t_desc_tab(c).col_name,
             p_sheet => t_sheet,
             p_alignment => p_alignment(c));
        else
          cell((c + p_col - 1),
             p_row,
             t_desc_tab(c).col_name,
             p_sheet => t_sheet);
        end if;

        if p_autofitcolumns then
          set_autofitcolumn(
            t_desc_tab(c).col_name,
            (c + p_col - 1),
            t_sheet);
        end if;
      end if;
      --      dbms_output.put_line( t_desc_tab( c ).col_name || ' ' || t_desc_tab( c ).col_type );
      case
        when t_desc_tab(c).col_type in (2, 100, 101) then
          dbms_sql.define_array(t_c, c, n_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
          dbms_sql.define_array(t_c, c, d_tab, t_bulk_size, 1);
        when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
          dbms_sql.define_array(t_c, c, v_tab, t_bulk_size, 1);
        else
          null;
      end case;
    end loop;
    --
    -- MAMJ --
    IF p_row IS NULL OR p_row = 0 THEN
      t_cur_row := case
                     when p_column_headers then
                      2
                     else
                      1
                   end;
    ELSE
      t_cur_row := p_row + 1;
    END IF;
    -- fim --

    --t_cur_row := case when p_column_headers then 2 else 1 end;  -- MAMJ

    --
    t_r := dbms_sql.execute(t_c);
    loop
      t_r := dbms_sql.fetch_rows(t_c);
      if t_r > 0 then
        for c in 1 .. t_col_cnt loop
          case
            when t_desc_tab(c).col_type in (2, 100, 101) then
              dbms_sql.column_value(t_c, c, n_tab);
              for i in 0 .. t_r - 1 loop
                if n_tab(i + n_tab.first()) is not null then
                  if p_alignment.exists(c) then
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         n_tab(i + n_tab.first()),
                         p_sheet => t_sheet,
                         p_alignment => p_alignment(c));
                  else
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         n_tab(i + n_tab.first()),
                         p_sheet => t_sheet);
                  end if;

                  if p_autofitcolumns then
                    set_autofitcolumn(
                      n_tab(i + n_tab.first()),
                      (c + p_col - 1),
                      t_sheet);
                  end if;
                end if;
              end loop;
              n_tab.delete;
            when t_desc_tab(c).col_type in (12, 178, 179, 180, 181, 231) then
              dbms_sql.column_value(t_c, c, d_tab);
              for i in 0 .. t_r - 1 loop
                if d_tab(i + d_tab.first()) is not null then
                  if p_alignment.exists(c) then
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         d_tab(i + d_tab.first()),
                         p_sheet => t_sheet,
                         p_alignment => p_alignment(c));
                  else
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         d_tab(i + d_tab.first()),
                         p_sheet => t_sheet);
                  end if;

                  if p_autofitcolumns then
                    set_autofitcolumn(
                      d_tab(i + d_tab.first()),
                      (c + p_col - 1),
                      t_sheet);
                  end if;
                end if;
              end loop;
              d_tab.delete;
            when t_desc_tab(c).col_type in (1, 8, 9, 96, 112) then
              dbms_sql.column_value(t_c, c, v_tab);
              for i in 0 .. t_r - 1 loop
                if v_tab(i + v_tab.first()) is not null then
                  if p_alignment.exists(c) then
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         v_tab(i + v_tab.first()),
                         p_sheet => t_sheet,
                         p_alignment => p_alignment(c));
                  else
                    cell((c + p_col - 1),
                         t_cur_row + i,
                         v_tab(i + v_tab.first()),
                         p_sheet => t_sheet);
                  end if;

                  if p_autofitcolumns then
                    set_autofitcolumn(
                      v_tab(i + v_tab.first()),
                      (c + p_col - 1),
                      t_sheet);
                  end if;
                end if;
              end loop;
              v_tab.delete;
            else
              null;
          end case;
        end loop;
      end if;
      exit when t_r != t_bulk_size;
      t_cur_row := t_cur_row + t_r;
    end loop;

    if p_autofitcolumns then
      for col in 1..t_col_cnt loop
        t_width := workbook.sheets(t_sheet).autofitcols((col + p_col - 1)).width;
        set_column_width((col + p_col - 1),t_width,t_sheet);
      end loop;
    end if;

    if p_set_autofilter then
      as_xlsx.set_autofilter(p_col,
                             (p_col + t_col_cnt - 1),
                             p_row_start => p_row,
                             p_row_end => p_row + workbook.sheets(t_sheet).rows.last() - 1);
      --as_xlsx.set_autofilter(1,1,10,12);
    end if;

    if p_tablestyle is not null then
      set_tablestyle(p_col,
                     (p_col + t_col_cnt - 1),
                     p_row,
                     workbook.sheets(t_sheet).rows.last(),
                     p_tablestyle,
                     t_sheet);
    end if;

    dbms_sql.close_cursor(t_c);

    if (p_directory is not null and p_filename is not null) then
      save(p_directory, p_filename);
    end if;
  exception
    when others then
      if dbms_sql.is_open(t_c) then
        dbms_sql.close_cursor(t_c);
      end if;
  end;

  function get_last_row (
    p_sheet pls_integer
  ) return number is
  begin
    return workbook.sheets(p_sheet).rows.last();
  end;

end;
/
