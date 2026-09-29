*&---------------------------------------------------------------------*
*&  Include    /SBXC/ZCKPRLT01_TOP                                     *
*&---------------------------------------------------------------------*

PROGRAM  /sbxc/zckprlt01      MESSAGE-ID /sbxc/zsp_cockpit.

TYPE-POOLS: rsds, icon.

** Tabelas
TABLES: /sbxc/zckp_ctrl,         " Tabela de controlo
        /sbxc/zckp_tab00,        " Tabela de processos
        /sbxc/zckp_tab03,        " Opções processamento (botões)
        /sbxc/zckp_invh,         " Invoice header
        t001.                    " Empresas

CONSTANTS: gc_stable TYPE lvc_s_stbl VALUE 'XX',
           lc_error  TYPE symsgty VALUE 'E'.

*----------------------------------------------------------------------*
*  Ecrã de selecção                                                    *
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-002.
SELECT-OPTIONS: p_proc  FOR /sbxc/zckp_tab00-processo,
                p_ano   FOR /sbxc/zckp_ctrl-ano,
                p_seqno FOR /sbxc/zckp_ctrl-seqno,
                p_bukrs FOR t001-bukrs.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-003.
SELECT-OPTIONS: p_data    FOR /sbxc/zckp_ctrl-data_in,
                p_hora    FOR /sbxc/zckp_ctrl-hora_in,
                p_user    FOR /sbxc/zckp_ctrl-user_in,
                p_stout  FOR /sbxc/zckp_ctrl-status1,
                p_mot_nc FOR /sbxc/zckp_invh-mot_n_contab.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-004.
SELECT-OPTIONS: p_refdoc FOR /sbxc/zckp_invh-ref_doc_no,
                p_vendor FOR /sbxc/zckp_invh-vendor,
                p_docfi  FOR /sbxc/zckp_invh-doc_fi.
SELECTION-SCREEN END OF BLOCK b3.
SELECTION-SCREEN: SKIP 2.

*----------------------------------------------------------------------*
*  Variáveis                                                           *
*----------------------------------------------------------------------*

DATA: v_uname TYPE sy-uname.

DATA: p_status TYPE RANGE OF /sbxc/zckp_tab10-status.

TYPES: BEGIN OF fm_display_type,
         processo    TYPE /sbxc/zckp_tab00-processo,
         fm_cab_disp TYPE /sbxc/zckp_tab00-fm_cab_disp,
         fm_lin_disp TYPE /sbxc/zckp_tab00-fm_lin_disp,
         fm_dblclk   TYPE /sbxc/zckp_tab00-fm_dblclk,
       END OF fm_display_type.

TYPES: BEGIN OF gt_botoes,
         est_botao TYPE /sbxc/zckp_tab03-est_botao,
         alvhl     TYPE /sbxc/zckp_tab03-alvhl,
         num       TYPE /sbxc/zckp_tab03-num,
         function  TYPE /sbxc/zckp_tab03-function, "user command
         icon      TYPE /sbxc/zckp_tab03-icon,
         butn_type TYPE /sbxc/zckp_tab03-butn_type,
         disabled  TYPE /sbxc/zckp_tab03-disabled,
         fm        TYPE /sbxc/zckp_tab03-fm,       " módulo de função
         num_pai   TYPE /sbxc/zckp_tab03-num_pai,
         tooltip   TYPE /sbxc/zckp_tab3t-quickinfo,
       END OF gt_botoes.

DATA: BEGIN OF key,
        processo TYPE /sbxc/zckp_tab00-processo,
        ano      TYPE /sbxc/zckp_ctrl-ano,
        seqno    TYPE /sbxc/zckp_ctrl-seqno,
      END OF key.

DATA: str_cab                TYPE /sbxc/zckp_tab00-est_cab,  " Cabeçalho
      str_lin                TYPE /sbxc/zckp_tab00-est_lin,  " Linha
      est_botao              TYPE /sbxc/zckp_tab03-est_botao, " Opções (botões)
      it_botoes_s            TYPE TABLE OF gt_botoes, "Status
      it_botoes              TYPE TABLE OF gt_botoes, "cabeçalho
      it_botoes_l            TYPE TABLE OF gt_botoes, "linha
      it_fm_display          TYPE TABLE OF fm_display_type,
      wa_fm_display          TYPE fm_display_type,
      r_click(1),
      option(10),
      low(10),
      line                   TYPE rsdswhere,
      line_tmp               TYPE rsdswhere,
      gt_proc                TYPE TABLE OF /sbxc/zckp_tab00,  " Dados do processo
      gt_status              TYPE TABLE OF /sbxc/zckp_tab_status,  " Dados dos status
      wa_status              TYPE /sbxc/zckp_tab_status,
      ds_clauses             TYPE rsds_where,
      t_fieldcat_cab         TYPE lvc_t_fcat,
      t_fieldcat_lin         TYPE lvc_t_fcat,
      t_fieldcat_status      TYPE lvc_t_fcat,
      wa_fieldcat            LIKE LINE OF t_fieldcat_cab,
      r_dyn_table_cab        TYPE REF TO data,
      r_wa_dyn_table_cab     TYPE REF TO data,
      r_dyn_table_cab_tmp    TYPE REF TO data,
      r_wa_dyn_table_cab_tmp TYPE REF TO data,
      r_dyn_table_lin        TYPE REF TO data,
      r_wa_dyn_table_lin     TYPE REF TO data,
      r_dyn_table_lin2       TYPE REF TO data,
      r_wa_dyn_table_lin2    TYPE REF TO data,
      pos                    TYPE sy-index,
      campo(40),
      wa_edit                TYPE lvc_s_styl,
      it_edit                TYPE lvc_t_styl,
      wa_color               TYPE lvc_s_scol,
      it_color               TYPE lvc_t_scol,
      wa_cor                 TYPE lvc_s_colo,
      w_cor                  TYPE  /sbxc/zckp_color,
      wa_zckp_ctrl           TYPE /sbxc/zckp_ctrl,
      f_cor                  TYPE TABLE OF /sbxc/zckp_color,
      lt_ui_funct            TYPE ui_functions,
      lt_ui_funct_status     TYPE ui_functions,
      lt_ui_fun_lin          TYPE ui_functions,
      ls_vari                TYPE disvariant,
      lin_or_cab(3),
      idx_lin                TYPE sy-index,
      erro_status(1),
      lt_filter              TYPE lvc_t_filt,
      ls_filter              TYPE lvc_s_filt,
      field                  TYPE lvc_s_fcat,
      conteudo(100),
      origem(3),
      refresh_table          TYPE c,
      status1_tmp            TYPE /sbxc/zckp_ctrl-status1,
      fname                  TYPE rs38l-name,
      l_status_final         TYPE /sbxc/zckp_status1,
      gd_repid               TYPE sy-repid,
      save_ok                TYPE sy-ucomm,
      ok_code                TYPE sy-ucomm,
      gd_changes             TYPE c,
      funcao_pre             TYPE rs38l-name,
      funcao                 TYPE rs38l-name,
      funcao_pos             TYPE rs38l-name,
      gs_f4                  TYPE lvc_s_f4,
      gt_f4                  TYPE lvc_t_f4,
      gt_dd03l               TYPE TABLE OF dd03l,
      gt_tab04               TYPE TABLE OF /sbxc/zckp_tab04,
      gt_tab05               TYPE TABLE OF /sbxc/zckp_tab05,
      gt_tab10               TYPE TABLE OF /sbxc/zckp_tab10,
      gt_tab10t              TYPE TABLE OF /sbxc/zckptab10t, "Status - Textos
      gs_tab10               TYPE /sbxc/zckp_tab10,
      linha_cab_sel_ind      TYPE TABLE OF lvc_s_row,
      linha_cab_sel          TYPE TABLE OF lvc_s_roid,
      it_ctrl                TYPE TABLE OF /sbxc/zckp_ctrl.

DATA: lt_zckp_tab3t TYPE TABLE OF /sbxc/zckp_tab3t.

DATA: lt_tab12 TYPE TABLE OF /sbxc/zckp_tab12,
      lt_hmail TYPE TABLE OF /sbxc/zckp_hmail,
      lt_tab09 TYPE TABLE OF /sbxc/zckp_tab09.

FIELD-SYMBOLS:
  <t_cab_table>     TYPE STANDARD TABLE,
  <t_cab_table_tmp> TYPE STANDARD TABLE,
  <wa_cab_table>    TYPE any,
  <t_lin_table>     TYPE STANDARD TABLE,
  <wa_lin_table>    TYPE any,
  <t_lin_table2>    TYPE STANDARD TABLE,
  <wa_lin_table2>   TYPE any,
  <cabecalho>       TYPE any,
  <fs_style>        TYPE lvc_t_styl,
  <fs_color>        TYPE lvc_t_scol,
  <cab>             TYPE any,
  <linha>           TYPE any,
  <valor>           TYPE any.

* Variaveis para ALV
DATA:
  go_docking         TYPE REF TO cl_gui_docking_container,
  go_splitter        TYPE REF TO cl_gui_splitter_container,
  go_splitter1       TYPE REF TO cl_gui_splitter_container,
  go_splitter2       TYPE REF TO cl_gui_splitter_container,
  go_splitter3       TYPE REF TO cl_gui_splitter_container,
  go_cell_main       TYPE REF TO cl_gui_container,
  go_cell_att        TYPE REF TO cl_gui_container,
  go_cell_top        TYPE REF TO cl_gui_container,
  go_cell_top1       TYPE REF TO cl_gui_container,
  go_cell_top2       TYPE REF TO cl_gui_container,
  go_cell_top3       TYPE REF TO cl_gui_container,
  go_cell_bottom     TYPE REF TO cl_gui_container,
  go_grid_proc       TYPE REF TO cl_gui_alv_grid,
  go_grid_status     TYPE REF TO cl_gui_alv_grid,
  go_grid_lin        TYPE REF TO cl_gui_alv_grid,
  go_grid_cab        TYPE REF TO cl_gui_alv_grid,
  go_attviewer       TYPE REF TO cl_gui_html_viewer,
  comp_container     TYPE i,
  comp_container_fil TYPE i,
  disp_doc_active(1) VALUE '',
  disp_fil_active(1) VALUE 'X',
  gs_layout          TYPE lvc_s_layo.

CLASS lcl_event_toolbar DEFINITION DEFERRED.

DATA: gs_grid2_toolbar        TYPE stb_button,
      go_event_toolbar        TYPE REF TO lcl_event_toolbar,

** botoes ALV linhas
      gs_grid3_toolbar        TYPE stb_button,
      go_event_toolbar_lin    TYPE REF TO lcl_event_toolbar,

** botoes ALV Status
      gs_grid4_toolbar        TYPE stb_button,
      go_event_toolbar_status TYPE REF TO lcl_event_toolbar.

**** AFG-ID-2500006 *****

DATA  : lo_cols       TYPE REF TO cl_salv_columns,
          lo_salv_table TYPE REF TO cl_salv_table,
          lo_column     TYPE REF TO cl_salv_column,
          col_name(30),
          col_desc(20).


  DATA : gr_struct_typ       TYPE REF TO  cl_abap_datadescr,
         gr_dyntable_typ     TYPE REF TO  cl_abap_tabledescr,
         gr_struct_typ_tmp   TYPE REF TO  cl_abap_datadescr,
         gr_dyntable_typ_tmp TYPE REF TO  cl_abap_tabledescr,
         ls_component        TYPE cl_abap_structdescr=>component,
         gt_component        TYPE         cl_abap_structdescr=>component_table,
         lo_element  TYPE REF TO cl_abap_elemdescr.

  DATA: lo_struct TYPE REF TO cl_abap_structdescr,
        ls_data   TYPE /sbxc/zckp_ctrl.


Data:
lt_comp     TYPE cl_abap_structdescr=>component_table,
lt_comp_lin     TYPE cl_abap_structdescr=>component_table.
