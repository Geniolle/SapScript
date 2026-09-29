*&---------------------------------------------------------------------*
*&  Include           ZCKPRLT01_F01                                    *
*&---------------------------------------------------------------------*
**-------------------------------------------------------------------*
**  MODIFICAÇÕES
*---------------------------------------------------------------------*
*& Autor        : A. Garrido (SBX)
*& Data         : 18/03/2025
*& Referência   : 2500006 - Dump Cockpit de Faturas
*& Transporte   : S4DK951531
*& Ch. pesquisa : AFG-ID-2500006
*& Objetivo     : Alteração da utilização do Create Dynamic Table para funcionalidade RTTS
*
*&---------------------------------------------------------------------*



*&---------------------------------------------------------------------*
*&      Form  VALIDA_PROJECTOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM valida_processos .
*1) Validar q processos têm autorização e se todos os processos
*   seleccionados têm as mesmas estruturas e opções

  DATA: wa_botoes TYPE gt_botoes,
        gs_proc   TYPE /sbxc/zckp_tab00.

  REFRESH: gt_proc.
  CLEAR: str_cab, str_lin, est_botao.

*1.1) validar q proj tem autorização
  SELECT * INTO @DATA(ls_tab00) FROM  /sbxc/zckp_tab00
    WHERE  processo IN @p_proc
    ORDER BY PRIMARY KEY.
    MOVE-CORRESPONDING ls_tab00 TO /sbxc/zckp_tab00.

    CALL FUNCTION '/SBXC/ZCKP_VALIDA_PROCESSO'
      EXPORTING
        processo        = /sbxc/zckp_tab00-processo
      IMPORTING
        str_proc        = gs_proc
      EXCEPTIONS
        sem_autorizacao = 1
        nao_existe      = 2
        OTHERS          = 3.
    IF sy-subrc EQ 0.
      MOVE-CORRESPONDING /sbxc/zckp_tab00 TO gs_proc.
      APPEND gs_proc TO gt_proc.
    ENDIF.
  ENDSELECT.

*1.2) Validar se estruturas são iguais
  LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc>).
    IF sy-tabix = 1.
      str_cab = <fs_proc>-est_cab.
      str_lin = <fs_proc>-est_lin.
      est_botao = <fs_proc>-est_botao.
    ENDIF.

    IF str_cab <> <fs_proc>-est_cab OR
       str_lin <> <fs_proc>-est_lin OR
       est_botao <> <fs_proc>-est_botao.
      MESSAGE e000.              " cancelar cockpit
    ENDIF.
  ENDLOOP.


*1.3) Validar se existe alguma

  IF str_cab EQ space.
* Sem autorização para o processo selecionado.
    MESSAGE e004.              " cancelar cockpit. Sem autorização.
  ENDIF.

  "Selecionar textos dos botões
  SELECT * INTO TABLE lt_zckp_tab3t
  FROM /sbxc/zckp_tab3t CLIENT SPECIFIED
   WHERE mandt = sy-mandt
     AND spras = sy-langu.


*1.4) Validar opções
  REFRESH: it_botoes_s, it_botoes.
  LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc5>).
*    Colocar a estrutura de botões seleccionada.
*    Se diferentes devido aos processos seleccionados dá erro.

    est_botao = <fs_proc5>-est_botao.

    SELECT * INTO @DATA(ls_tab03) FROM /sbxc/zckp_tab03 WHERE est_botao = @est_botao.
      MOVE-CORRESPONDING ls_tab03 TO /sbxc/zckp_tab03.
      MOVE-CORRESPONDING /sbxc/zckp_tab03 TO wa_botoes.

      READ TABLE lt_zckp_tab3t ASSIGNING FIELD-SYMBOL(<fs_3t>) WITH KEY est_botao  = /sbxc/zckp_tab03-est_botao
                                                                            num    = /sbxc/zckp_tab03-num
                                                                            alvhl  = ls_tab03-alvhl.
      IF sy-subrc EQ 0.
        wa_botoes-tooltip = <fs_3t>-quickinfo.
      ENDIF.

      IF ls_tab03-alvhl = 'S'. "status

* verifica autorizações para os botões
        LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc1>).
          AUTHORITY-CHECK OBJECT 'ZCKP:BTNID' FOR USER v_uname
          ID '/SBXC/CBTN' FIELD /sbxc/zckp_tab03-function
          ID '/SBXC/CPRO' FIELD <fs_proc1>-processo.
          IF sy-subrc NE 0.
            wa_botoes-disabled = 'X'.
            EXIT.                                       "#EC CI_NOORDER
          ENDIF.
        ENDLOOP.
        APPEND wa_botoes TO it_botoes_s.
        CLEAR wa_botoes.
      ELSEIF ls_tab03-alvhl = 'H'. "estrutura de cabeçalho

        LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc2>).
          AUTHORITY-CHECK OBJECT 'ZCKP:BTNID' FOR USER v_uname
          ID '/SBXC/CBTN' FIELD /sbxc/zckp_tab03-function
          ID '/SBXC/CPRO' FIELD <fs_proc2>-processo.
          IF sy-subrc NE 0.
            wa_botoes-disabled = 'X'.
            EXIT.                                       "#EC CI_NOORDER
          ENDIF.
        ENDLOOP.
        APPEND wa_botoes TO it_botoes.
        CLEAR wa_botoes.
      ELSEIF ls_tab03-alvhl = 'L'. " estrutura de linha

        LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc3>).
          AUTHORITY-CHECK OBJECT 'ZCKP:BTNID' FOR USER v_uname
          ID '/SBXC/CBTN' FIELD /sbxc/zckp_tab03-function
          ID '/SBXC/CPRO' FIELD <fs_proc3>-processo.
          IF sy-subrc NE 0.
            wa_botoes-disabled = 'X'.
            EXIT.                                       "#EC CI_NOORDER
          ENDIF.
        ENDLOOP.
        APPEND wa_botoes TO it_botoes_l.
      ENDIF.
    ENDSELECT.

  ENDLOOP.

  SORT it_botoes BY est_botao num.
  DELETE ADJACENT DUPLICATES FROM it_botoes COMPARING est_botao num.

  SORT it_botoes BY function.
* Ler processos - Funções display

  LOOP AT gt_proc ASSIGNING FIELD-SYMBOL(<fs_proc4>).
    wa_fm_display-processo = <fs_proc4>-processo.
    wa_fm_display-fm_cab_disp = <fs_proc4>-fm_cab_disp.
    wa_fm_display-fm_lin_disp = <fs_proc4>-fm_lin_disp.
    wa_fm_display-fm_dblclk = <fs_proc4>-fm_dblclk .
    APPEND wa_fm_display TO it_fm_display.
  ENDLOOP.

ENDFORM.                    " VALIDA_PROJECTOS

*&---------------------------------------------------------------------*
*&      Form  INICIALIZA_CONTROL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM inicializa_control .

  CONSTANTS: c_90 TYPE i VALUE 90,
             c_70 TYPE i VALUE 70,
             c_20 TYPE i VALUE 20,
             c_60 TYPE i VALUE 60.

* Create docking container
  CREATE OBJECT go_docking
    EXPORTING
      parent = cl_gui_container=>screen0
      ratio  = c_90
    EXCEPTIONS
      OTHERS = 6.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

* ALV PROCESSOS

  CREATE OBJECT go_splitter ##SUBRC_OK
    EXPORTING
      parent            = go_docking
      rows              = 1
      columns           = 2
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3.
  CALL METHOD go_splitter->get_container
    EXPORTING
      row       = 1
      column    = 1
    RECEIVING
      container = go_cell_main.

  CALL METHOD go_splitter->get_container
    EXPORTING
      row       = 1
      column    = 2
    RECEIVING
      container = go_cell_att.

  CREATE OBJECT go_attviewer
    EXPORTING
      parent = go_cell_att.
* Create splitter
  CREATE OBJECT go_splitter1 ##SUBRC_OK
    EXPORTING
      parent            = go_cell_main
      rows              = 2
      columns           = 1
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3.

* separador horizontal
  go_splitter1->set_row_height( id = 1 height = c_70 ).

* Get cell container
  CALL METHOD go_splitter1->get_container
    EXPORTING
      row       = 1
      column    = 1
    RECEIVING
      container = go_cell_top.

  CALL METHOD go_splitter1->get_container
    EXPORTING
      row       = 2
      column    = 1
    RECEIVING
      container = go_cell_bottom.

  CREATE OBJECT go_splitter2 ##SUBRC_OK
    EXPORTING
      parent            = go_cell_top
      rows              = 1
      columns           = 2
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3.


* separador vertical (de processos e status)
  go_splitter2->set_column_width(
    EXPORTING
      id                = 1
      width             = c_20
*  IMPORTING
*    result            =
  EXCEPTIONS
    cntl_error        = 1
    cntl_system_error = 2
    OTHERS            = 3 ) ##SUBRC_OK.


  CALL METHOD go_splitter2->get_container
    EXPORTING
      row       = 1
      column    = 1
    RECEIVING
      container = go_cell_top1.

  CALL METHOD go_splitter2->get_container
    EXPORTING
      row       = 1
      column    = 2
    RECEIVING
      container = go_cell_top2.

  CREATE OBJECT go_splitter3
    EXPORTING
      parent            = go_cell_top1
      rows              = 2
      columns           = 1
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.

  go_splitter3->set_row_height( id = 2 height = c_60 ).

  CALL METHOD go_splitter3->get_container
    EXPORTING
      row       = 1
      column    = 1
    RECEIVING
      container = go_cell_top1.

  CALL METHOD go_splitter3->get_container
    EXPORTING
      row       = 2
      column    = 1
    RECEIVING
      container = go_cell_top3.


  CREATE OBJECT go_grid_proc
    EXPORTING
      i_parent = go_cell_top1
    EXCEPTIONS
      OTHERS   = 5 ##SUBRC_OK.


  CREATE OBJECT go_grid_cab
    EXPORTING
      i_parent = go_cell_top2
    EXCEPTIONS
      OTHERS   = 5 ##SUBRC_OK.


  CREATE OBJECT go_grid_lin
    EXPORTING
      i_parent = go_cell_bottom
    EXCEPTIONS
      OTHERS   = 5 ##SUBRC_OK.

  CREATE OBJECT go_grid_status
    EXPORTING
      i_parent = go_cell_top3
    EXCEPTIONS
      OTHERS   = 5 ##SUBRC_OK.
  IF disp_doc_active IS INITIAL.
    PERFORM esconde_documento.
  ENDIF.
ENDFORM.                    " INICIALIZA_CONTROL
*&---------------------------------------------------------------------*
*&      Form  esconde_documento
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM esconde_documento.
  CALL METHOD go_splitter->set_column_width
    EXPORTING
      id                = 2
      width             = 0
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
*  IF sy-subrc <> 0.
** Implement suitable error handling here
*  ENDIF.
  CALL METHOD go_splitter->set_column_sash
    EXPORTING
      id                = 2
      type              = cl_gui_splitter_container=>type_sashvisible
      value             = cl_gui_splitter_container=>false
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter->set_column_sash
    EXPORTING
      id                = 2
      type              = cl_gui_splitter_container=>type_movable
      value             = cl_gui_splitter_container=>false
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CLEAR disp_doc_active.
ENDFORM.                    " ESCONDE_PESQUISA
*&---------------------------------------------------------------------*
*&      Form  esconde_filtros
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM esconde_filtros.
  CALL METHOD go_splitter2->get_column_width
    EXPORTING
      id                = 1
    IMPORTING
      result            = comp_container_fil
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter2->set_column_width
    EXPORTING
      id                = 1
      width             = 0
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
*  IF sy-subrc <> 0.
** Implement suitable error handling here
*  ENDIF.
  CALL METHOD go_splitter2->set_column_sash
    EXPORTING
      id                = 1
      type              = cl_gui_splitter_container=>type_sashvisible
      value             = cl_gui_splitter_container=>false
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter2->set_column_sash
    EXPORTING
      id                = 1
      type              = cl_gui_splitter_container=>type_movable
      value             = cl_gui_splitter_container=>false
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CLEAR disp_fil_active.
ENDFORM.                    " ESCONDE_PESQUISA
*&---------------------------------------------------------------------*
*&      Form  exibe_documento
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM exibe_documento.

  CONSTANTS: c_50 TYPE i VALUE 50.
  IF comp_container IS INITIAL.
    comp_container = c_50.
  ENDIF.
  CALL METHOD go_splitter->set_column_width
    EXPORTING
      id                = 2
      width             = comp_container
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter->set_column_sash
    EXPORTING
      id                = 2
      type              = cl_gui_splitter_container=>type_sashvisible
      value             = cl_gui_splitter_container=>true
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter->set_column_sash
    EXPORTING
      id                = 2
      type              = cl_gui_splitter_container=>type_movable
      value             = cl_gui_splitter_container=>true
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  disp_doc_active = 'X'.
  PERFORM show_document USING <cab>.
  PERFORM esconde_filtros.
ENDFORM.                    " EXIBE_PESQUISA
*&---------------------------------------------------------------------*
*&      Form  exibe_filtros
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM exibe_filtros .

  CONSTANTS: c_20 TYPE i VALUE 20.

  IF comp_container_fil IS INITIAL.
    comp_container_fil = c_20.
  ENDIF.
  CALL METHOD go_splitter2->set_column_width
    EXPORTING
      id                = 1
      width             = comp_container_fil
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter2->set_column_sash
    EXPORTING
      id                = 1
      type              = cl_gui_splitter_container=>type_sashvisible
      value             = cl_gui_splitter_container=>true
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  CALL METHOD go_splitter2->set_column_sash
    EXPORTING
      id                = 1
      type              = cl_gui_splitter_container=>type_movable
      value             = cl_gui_splitter_container=>true
    EXCEPTIONS
      cntl_error        = 1
      cntl_system_error = 2
      OTHERS            = 3 ##SUBRC_OK.
  PERFORM esconde_documento.
  disp_fil_active = 'X'.
ENDFORM.                    " EXIBE_PESQUISA
*&---------------------------------------------------------------------*
*&      Form  BUILD_ALVS_INIT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM build_alvs_init .

* Processos
  gs_layout-grid_title = TEXT-pro. "'Processos'.
  gs_layout-no_rowmark = 'X'.
  gs_layout-no_toolbar = 'X'.
  gs_layout-cwidth_opt = 'X'.


  CALL METHOD go_grid_proc->set_table_for_first_display
    EXPORTING
      i_structure_name = '/SBXC/ZCKP_TAB00'               " Estrutura
      is_layout        = gs_layout
    CHANGING
      it_outtab        = gt_proc
    EXCEPTIONS
      OTHERS           = 4 ##SUBRC_OK.


* Cabeçalhos
  gs_layout-grid_title = TEXT-cab. "'Cabeçalhos'.
  gs_layout-no_toolbar = ''.
  gs_layout-no_rowmark = ' '.
  gs_layout-sel_mode = 'C'.

*Área de cabeçalho colocar campos da tabela de controlo + os da
*estrutura definida para cabeçalho
  PERFORM estrutura_cab.

  PERFORM toolbar_excluding CHANGING lt_ui_funct.

  DELETE t_fieldcat_cab WHERE fieldname = 'CELLSTYLE'.
  DELETE t_fieldcat_cab WHERE fieldname = 'CELL_COLOR'.


  gs_layout-stylefname = 'CELLSTYLE'.
  gs_layout-ctab_fname  = 'CELL_COLOR'.
  gs_layout-zebra = 'X'.


*Codigo IVA
*Register the field for which the custom F4 has to be displayed

  gs_f4-fieldname  = 'TAX_CODE_SAP'.
  gs_f4-register   = 'X'.
  gs_f4-getbefore  = space.
  gs_f4-chngeafter = space.
  APPEND gs_f4 TO gt_f4.


  CALL METHOD go_grid_lin->register_f4_for_fields
    EXPORTING
      it_f4 = gt_f4.


  ls_vari-report       = sy-repid.
  ls_vari-username     = v_uname.
  ls_vari-handle       = 'CAB'.
* ALV de Cabeçalho
  PERFORM domain_xware_bnk.
  CALL METHOD go_grid_cab->set_table_for_first_display
    EXPORTING
      is_layout            = gs_layout
      it_toolbar_excluding = lt_ui_funct
      is_variant           = ls_vari
      i_save               = 'A'
    CHANGING
      it_fieldcatalog      = t_fieldcat_cab[]
      it_outtab            = <t_cab_table>
    EXCEPTIONS
      OTHERS               = 4 ##SUBRC_OK.

  CALL METHOD go_grid_cab->set_ready_for_input
    EXPORTING
      i_ready_for_input = 1.

** SALV Declarations.
*  TRY.
*      cl_salv_table=>factory(
*      IMPORTING
*      r_salv_table   = lo_salv_table
*      CHANGING
*      t_table        = <t_cab_table>
*      ).
*    CATCH cx_salv_msg .
*  ENDTRY.
*
** get columns object
*  lo_cols = lo_salv_table->get_columns( ).
*
**…Individual Column Names
*  LOOP AT lt_comp INTO ls_component.
*    TRY.
*        col_name = LS_COMPONENT-name.
*        lo_column = lo_cols->get_column( col_name ). " < <
*        col_desc = LS_COMPONENT-type->get_relative_name( ).
*        lo_column->set_medium_text( col_desc ).
*
*        lo_column->set_output_length( CONV #( LS_COMPONENT-type->length ) ).
*      CATCH cx_salv_not_found.                          "#EC NO_HANDLER
*    ENDTRY.
*  ENDLOOP.
** display table
*
*  lo_salv_table->display( ).

**********************************************************************
* Status
  CLEAR gs_layout.
  gs_layout-grid_title = TEXT-sta. "'Status'.
  gs_layout-no_toolbar = ''.
  gs_layout-no_rowmark = ' '.
  gs_layout-cwidth_opt = 'X'.
  gs_layout-sel_mode = 'C'.
  gs_layout-box_fname = 'BOX'.


  PERFORM get_fieldcat_status.
  PERFORM get_tab_status.
  PERFORM toolbar_excluding_status CHANGING lt_ui_funct_status.
* alv status
  CALL METHOD go_grid_status->set_table_for_first_display
    EXPORTING
      i_structure_name     = '/SBXC/ZCKP_TAB_STATUS'         " Estrutura
      is_layout            = gs_layout
      it_toolbar_excluding = lt_ui_funct_status
    CHANGING
      it_fieldcatalog      = t_fieldcat_status[]
      it_outtab            = gt_status
    EXCEPTIONS
      OTHERS               = 4 ##SUBRC_OK.

  <t_cab_table_tmp>[] = <t_cab_table>[].
  CLEAR gs_layout.
**********************************************************************


* Área de linhas colocar os definidos para linha ( se existir )
  PERFORM estrutura_lin.
  ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
              <fs_style>.
**  PERFORM estilo_campos USING 'LIN'.

  gs_layout-grid_title = TEXT-lin. "'Linhas'.
  gs_layout-ctab_fname  = 'CELL_COLOR'.

  gs_layout-cwidth_opt = 'X'.
  gs_layout-no_toolbar = ''.
  gs_layout-no_rowmark = ' '.
  gs_layout-sel_mode = 'C'.

  gs_layout-edit = 'X'.


  DELETE t_fieldcat_lin WHERE fieldname = 'CELLSTYLE'.
  DELETE t_fieldcat_lin WHERE fieldname = 'CELL_COLOR'.

  PERFORM toolbar_excluding_lin CHANGING  lt_ui_fun_lin.
* alv linhas
  ls_vari-handle       = 'LIN'.
  CALL METHOD go_grid_lin->set_table_for_first_display
    EXPORTING
      is_layout            = gs_layout
      it_toolbar_excluding = lt_ui_fun_lin
      is_variant           = ls_vari
      i_save               = 'A'
* Fim inserção
    CHANGING
      it_outtab            = <t_lin_table>
      it_fieldcatalog      = t_fieldcat_lin[].

  CALL METHOD go_grid_cab->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_enter.

  CALL METHOD go_grid_cab->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_modified.


  CALL METHOD go_grid_lin->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_enter.

  CALL METHOD go_grid_lin->register_edit_event
    EXPORTING
      i_event_id = cl_gui_alv_grid=>mc_evt_modified.

**Hotspot STATUS
  SET HANDLER lcl_event_receiver=>handle_hotspot_click FOR go_grid_status.

** Menu contexto
  SET HANDLER lcl_event_receiver=>handle_context_menu FOR go_grid_cab.
  SET HANDLER lcl_event_receiver=>handle_context_menu FOR go_grid_lin.

** Double Click
  SET HANDLER lcl_event_receiver=>handle_double_click FOR go_grid_proc.
  SET HANDLER lcl_event_receiver=>handle_double_click FOR go_grid_lin.
  SET HANDLER lcl_event_receiver=>handle_double_click FOR go_grid_cab.

** Data_changed
  SET HANDLER lcl_event_receiver=>handle_data_changed FOR go_grid_lin.
  SET HANDLER lcl_event_receiver=>handle_data_changed FOR go_grid_cab.


  SET HANDLER lcl_event_receiver=>on_f4 FOR go_grid_lin.

  CREATE OBJECT go_event_toolbar_status.
  SET HANDLER go_event_toolbar_status->handle_user_command
              go_event_toolbar_status->handle_menu_button
              go_event_toolbar_status->handle_toolbar_status FOR go_grid_status.


  CREATE OBJECT go_event_toolbar.
  SET HANDLER go_event_toolbar->handle_user_command
              go_event_toolbar->handle_menu_button
              go_event_toolbar->handle_toolbar
              FOR go_grid_cab.


  CREATE OBJECT go_event_toolbar_lin.
  SET HANDLER go_event_toolbar_lin->handle_user_command
              go_event_toolbar_lin->handle_menu_button
              go_event_toolbar_lin->handle_toolbar_lin FOR go_grid_lin.


  CALL METHOD go_grid_status->set_toolbar_interactive.
  CALL METHOD cl_gui_control=>set_focus
    EXPORTING
      control = go_grid_status.

  CALL METHOD go_grid_lin->set_toolbar_interactive.
  CALL METHOD cl_gui_control=>set_focus
    EXPORTING
      control = go_grid_lin.

  CALL METHOD go_grid_cab->set_toolbar_interactive.
  CALL METHOD cl_gui_control=>set_focus
    EXPORTING
      control = go_grid_cab.

  gd_repid  = syst-repid.
  CALL METHOD go_docking->link
    EXPORTING
      repid  = gd_repid
      dynnr  = '0100'
    EXCEPTIONS
      OTHERS = 4 ##SUBRC_OK.

ENDFORM.                    " BUILD_ALVS_INIT
*&---------------------------------------------------------------------*
*&      Form  ESTRUTURA_CAB
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM estrutura_cab.
  FIELD-SYMBOLS: <mensagem> TYPE any,
                 <buk>      TYPE any.

  DATA:auth_emp TYPE sy-subrc.


****************************************************
* Tabela de controlo
  SELECT        * INTO @DATA(wa_dd03l) FROM  dd03l
         WHERE  tabname    = '/SBXC/ZCKP_CTRL'
          AND comptype NE 'S'
         ORDER BY position.

    CHECK wa_dd03l-fieldname  =  'PROCESSO'   OR
          wa_dd03l-fieldname  =  'ANO'        OR
          wa_dd03l-fieldname  =  'SEQNO'      OR
          wa_dd03l-fieldname  =  'DATA_IN'    OR
          wa_dd03l-fieldname  =  'HORA_IN'    OR
          wa_dd03l-fieldname  =  'ICON_STATUS' OR
          wa_dd03l-fieldname  =  'STATUS1'    OR
          wa_dd03l-fieldname  =  'STATUS_OUT'    OR
          wa_dd03l-fieldname  =  'AUT_MAN'    OR
          wa_dd03l-fieldname  =  'ID_EXTERNO' OR
          wa_dd03l-fieldname =   'PROCESSO_ANT'  OR
          wa_dd03l-fieldname =   'ANO_ANT'       OR
          wa_dd03l-fieldname =   'SEQNO_ANT'     OR
          wa_dd03l-fieldname =   'PROCESSO_POST' OR
          wa_dd03l-fieldname =   'ANO_POST'      OR
          wa_dd03l-fieldname =   'SEQNO_POST' OR
          wa_dd03l-fieldname =   'USER_CHG_ST1'. "ODC - 09_03_2022

    CLEAR wa_fieldcat.
    PERFORM texto_campo USING wa_dd03l-tabname
                              wa_dd03l-fieldname
                        CHANGING wa_fieldcat-coltext
                           wa_fieldcat-seltext
                           wa_fieldcat-scrtext_l
                           wa_fieldcat-scrtext_m
                           wa_fieldcat-scrtext_s.


    wa_fieldcat-fieldname = wa_dd03l-fieldname.
    wa_fieldcat-inttype   = wa_dd03l-inttype.
    wa_fieldcat-outputlen = wa_dd03l-intlen.
    wa_fieldcat-ref_table = wa_dd03l-tabname.
    wa_fieldcat-ref_field = wa_dd03l-fieldname.
    wa_fieldcat-rollname = wa_dd03l-rollname.

    pos = pos + 1.
    wa_fieldcat-col_pos = pos.
    APPEND wa_fieldcat TO t_fieldcat_cab.



  ENDSELECT.



*Campos da estrutura definida para cabeçalho
  IF  /sbxc/zckp_tab00-est_cab NE space.
    SELECT        *  FROM  dd03l INTO @DATA(ls_dd03l)
           WHERE    tabname  = @/sbxc/zckp_tab00-est_cab
      AND comptype NE 'S'
      ORDER BY position.

      CHECK ls_dd03l-fieldname NE 'MANDT'.

      CHECK ls_dd03l-fieldname  NE 'PROCESSO'   AND
            ls_dd03l-fieldname  NE  'ANO'       AND
            ls_dd03l-fieldname  NE  'SEQNO'     AND
            ls_dd03l-fieldname  NE  'DATA_IN'   AND
            ls_dd03l-fieldname  NE  'HORA_IN'   AND
            ls_dd03l-fieldname  NE  'STATUS1'   AND
            ls_dd03l-fieldname  NE  'STATUS_OUT'   AND
            ls_dd03l-fieldname  NE  'AUT_MAN'   AND
            ls_dd03l-fieldname  NE  'ID_EXTERNO'.

      CLEAR wa_fieldcat.
      PERFORM texto_campo USING ls_dd03l-tabname
                                ls_dd03l-fieldname
                          CHANGING wa_fieldcat-coltext
                             wa_fieldcat-seltext
                             wa_fieldcat-scrtext_l
                             wa_fieldcat-scrtext_m
                             wa_fieldcat-scrtext_s.


      wa_fieldcat-fieldname = ls_dd03l-fieldname.
      wa_fieldcat-datatype = ls_dd03l-datatype.
      wa_fieldcat-inttype   = ls_dd03l-inttype.
      wa_fieldcat-intlen = ls_dd03l-intlen.
      wa_fieldcat-ref_table = ls_dd03l-tabname.
      wa_fieldcat-ref_field = ls_dd03l-fieldname.
      wa_fieldcat-rollname = ls_dd03l-rollname.
      pos = pos + 1.
      wa_fieldcat-col_pos = pos.
      wa_fieldcat-edit = 'X'.

      IF wa_fieldcat-fieldname EQ 'STCD1'.
        wa_fieldcat-scrtext_l = TEXT-055.
        wa_fieldcat-scrtext_m = TEXT-055.
        wa_fieldcat-scrtext_s = TEXT-056.
        wa_fieldcat-coltext = TEXT-056.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'COMP_VAT'.
        wa_fieldcat-scrtext_l = TEXT-057.
        wa_fieldcat-scrtext_m = TEXT-057.
        wa_fieldcat-scrtext_s = TEXT-058.
        wa_fieldcat-coltext = TEXT-058.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'NOTAS'.
        wa_fieldcat-just = 'X'.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'ENVIADO_SAPHETY'.
        wa_fieldcat-scrtext_l = TEXT-059.
        wa_fieldcat-scrtext_m = TEXT-059.
        wa_fieldcat-scrtext_s = TEXT-059.
        wa_fieldcat-coltext = TEXT-059.
      ENDIF.
      APPEND wa_fieldcat TO t_fieldcat_cab.
    ENDSELECT.

  ENDIF.

* Campos para controlo de edição
  pos = pos + 1.
  wa_fieldcat-col_pos = pos.
  wa_fieldcat-coltext = ' '.
  wa_fieldcat-seltext = ' '.
  wa_fieldcat-scrtext_l = ' '.
  wa_fieldcat-scrtext_m = ' '.
  wa_fieldcat-scrtext_s = ' '.
  wa_fieldcat-fieldname = 'CELLSTYLE'.
  wa_fieldcat-ref_table = '/SBXC/ZCKP_ALV'.
  wa_fieldcat-ref_field = 'CELL_STYLE'.
  APPEND wa_fieldcat TO t_fieldcat_cab.
  CLEAR wa_fieldcat.


* Campos para controlo de cor
  pos = pos + 1.
  wa_fieldcat-col_pos = pos.
  wa_fieldcat-coltext = ' '.
  wa_fieldcat-seltext = ' '.
  wa_fieldcat-scrtext_l = ' '.
  wa_fieldcat-scrtext_m = ' '.
  wa_fieldcat-scrtext_s = ' '.
  wa_fieldcat-fieldname = 'CELL_COLOR'.
  wa_fieldcat-ref_table = '/SBXC/ZCKP_ALV'.
  wa_fieldcat-ref_field = 'CELL_COLOR'.
  APPEND wa_fieldcat TO t_fieldcat_cab.
  CLEAR wa_fieldcat.


* Status
  pos = 6.
  wa_fieldcat-col_pos = pos.
  wa_fieldcat-coltext = ' '.
  wa_fieldcat-seltext = ' '.
  wa_fieldcat-scrtext_l = ' '.
  wa_fieldcat-scrtext_m = ' '.
  wa_fieldcat-scrtext_s = ' '.
  wa_fieldcat-fieldname = 'STATUS_OUT'.
  wa_fieldcat-ref_table = '/SBXC/ZCKP_TAB10'.
  wa_fieldcat-ref_field = 'STATUS_OUT'.
  APPEND wa_fieldcat TO t_fieldcat_cab.
  CLEAR wa_fieldcat.


** Criar tabela a partir do field catalog
*  CALL METHOD cl_alv_table_create=>create_dynamic_table
*    EXPORTING
*      it_fieldcatalog           = t_fieldcat_cab
*    IMPORTING
*      ep_table                  = r_dyn_table_cab
*    EXCEPTIONS
*      generate_subpool_dir_full = 1
*      OTHERS                    = 2.
*
*  IF sy-subrc <> 0.
*    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
*  ENDIF.
*
*
*
*
*
*    ENDSELECT.
*
*  ENDIF.

**** AFG-ID-2500006 *****
  LOOP AT t_fieldcat_cab INTO wa_fieldcat.
    IF wa_fieldcat-fieldname EQ 'CELL_COLOR'.
      ls_component-name = 'CELL_COLOR'.
      ls_component-type ?= cl_abap_datadescr=>describe_by_name( '/SBXC/ZCKP_ALV-CELL_COLOR' ).
    ELSEIF wa_fieldcat-fieldname EQ 'CELLSTYLE'.
      ls_component-name = 'CELLSTYLE'.
      ls_component-type ?= cl_abap_datadescr=>describe_by_name( '/SBXC/ZCKP_ALV-CELL_STYLE' ).
    ELSEIF wa_fieldcat-fieldname EQ 'STATUS_OUT'.
      ls_component-name = 'STATUS_OUT'.
      ls_component-type ?= cl_abap_datadescr=>describe_by_name( '/SBXC/ZCKP_TAB10-STATUS_OUT' ).
    ELSE.
*   Element Description
      lo_element ?= cl_abap_elemdescr=>describe_by_name( wa_fieldcat-rollname ).
      ls_component-type = lo_element.
*   Field name
      LS_COMPONENT-name = wa_fieldcat-fieldname.
    ENDIF.

*   Filling the component table
    APPEND ls_component TO lt_comp.
    CLEAR: ls_component.
  ENDLOOP.

  gr_struct_typ  ?= cl_abap_structdescr=>create( p_components = lt_comp ).
  gr_dyntable_typ = cl_abap_tabledescr=>create( p_line_type = gr_struct_typ ).

  CREATE DATA r_dyn_table_cab TYPE HANDLE gr_dyntable_typ.
  CREATE DATA  r_wa_dyn_table_cab TYPE HANDLE gr_struct_typ.


  gr_struct_typ_tmp  ?= cl_abap_structdescr=>create( p_components = lt_comp ).
  gr_dyntable_typ_tmp = cl_abap_tabledescr=>create( p_line_type = gr_struct_typ ).

  CREATE DATA r_dyn_table_cab_tmp TYPE HANDLE gr_dyntable_typ_tmp.
  CREATE DATA  r_wa_dyn_table_cab_tmp TYPE HANDLE gr_struct_typ_tmp.

  ASSIGN r_dyn_table_cab->* TO <t_cab_table>.
  ASSIGN r_dyn_table_cab_tmp->* TO <t_cab_table_tmp>.


* Apontador para a tabela
*  ASSIGN r_dyn_table_cab->* TO <t_cab_table>.

* Criar workarea para a tabela
*  CREATE DATA r_wa_dyn_table_cab LIKE LINE OF <t_cab_table>.

**********************************************************************
** Criar tabela a partir do field catalog
*  CALL METHOD cl_alv_table_create=>create_dynamic_table
*    EXPORTING
*      it_fieldcatalog           = t_fieldcat_cab
*    IMPORTING
*      ep_table                  = r_dyn_table_cab_tmp
*    EXCEPTIONS
*      generate_subpool_dir_full = 1
*      OTHERS                    = 2.
*
*
*
*  IF sy-subrc <> 0.
*    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*               WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
*  ENDIF.
*
*
** Apontador para a tabela
*  ASSIGN r_dyn_table_cab_tmp->* TO <t_cab_table_tmp>.
** Criar workarea para a tabela
*  CREATE DATA r_wa_dyn_table_cab_tmp LIKE LINE OF <t_cab_table_tmp>.



**********************************************************************
  DATA: it_dd03l_buk TYPE TABLE OF dd03l.

  SELECT  * FROM  dd03l
    INTO TABLE it_dd03l_buk
         WHERE  tabname  = str_cab
         AND    rollname   = 'BUKRS'.

* Ler valores
  REFRESH <t_cab_table>.

  PERFORM ds_clauses CHANGING ds_clauses.

  DATA: it_ctrl_temp TYPE TABLE OF /sbxc/zckp_ctrl." WITH HEADER LINE.
  REFRESH it_ctrl_temp.

  SELECT * FROM /sbxc/zckp_ctrl INTO TABLE it_ctrl_temp
    WHERE processo IN p_proc
          AND ano IN p_ano
          AND seqno IN p_seqno
          AND data_in IN p_data
          AND hora_in IN p_hora
          AND user_in IN p_user
          AND status1 IN p_stout.
  LOOP AT it_ctrl_temp ASSIGNING FIELD-SYMBOL(<fs_ctrl>).
    CLEAR /sbxc/zckp_ctrl.
    MOVE-CORRESPONDING <fs_ctrl> TO /sbxc/zckp_ctrl.

* Ler entrada de cabeçalho se respeitarem selecção

    ASSIGN r_wa_dyn_table_cab->* TO <cabecalho>.

    MOVE-CORRESPONDING /sbxc/zckp_ctrl  TO <cabecalho>.

    SELECT SINGLE * FROM (str_cab)
    INTO CORRESPONDING FIELDS OF <cabecalho>
    WHERE processo = /sbxc/zckp_ctrl-processo
      AND ano = /sbxc/zckp_ctrl-ano
      AND seqno = /sbxc/zckp_ctrl-seqno
      AND comp_code IN p_bukrs
      AND ref_doc_no IN p_refdoc
      AND vendor  IN p_vendor
      AND doc_fi IN p_docfi
      AND mot_n_contab IN p_mot_nc.

    CHECK sy-subrc = 0.
    MOVE-CORRESPONDING /sbxc/zckp_ctrl  TO <cabecalho>.

    CLEAR auth_emp.
    LOOP AT it_dd03l_buk ASSIGNING FIELD-SYMBOL(<fs_buk>).

      CONCATENATE '<CABECALHO>-' <fs_buk>-fieldname INTO campo.
      ASSIGN (campo) TO <buk>.

*** Valida se está no ecrã de selecção
      CHECK  <buk> IN p_bukrs.

*** Valida se tem autorização
      AUTHORITY-CHECK OBJECT 'F_BKPF_BUK'
               ID 'BUKRS' FIELD <buk>
               ID 'ACTVT' FIELD '01'.
      IF sy-subrc NE 0.
        auth_emp = sy-subrc.
      ENDIF.

    ENDLOOP.

    CHECK auth_emp EQ 0.

    CONCATENATE '<CABECALHO>-' 'STATUS_OUT' INTO campo.
    ASSIGN (campo) TO <buk>.
    IF sy-subrc EQ 0.
      READ TABLE gt_tab10 ASSIGNING FIELD-SYMBOL(<fs_tab10>) WITH KEY status = /sbxc/zckp_ctrl-status1.
      <buk> = <fs_tab10>-status_out.
    ENDIF.

    FIELD-SYMBOLS: <status_cab> TYPE any.
    DATA l_status_cab(40).

    l_status_cab = '<CABECALHO>-STATUS1'.
    ASSIGN (l_status_cab) TO <status_cab>.


*Doc_fi************************************************************ini

*Stat 0 ou 5
**********************************************************************
    FIELD-SYMBOLS: <valor_ctrl>   TYPE any,
                   <gross_amount> TYPE any.
    DATA: campo_emp(35).
    IF /sbxc/zckp_ctrl-status1 EQ '0' OR /sbxc/zckp_ctrl-status1 EQ '5'.
      MOVE '<CABECALHO>-VALOR_CONTROLO' TO campo_emp.
      ASSIGN (campo_emp) TO <valor_ctrl>.

      MOVE '<CABECALHO>-GROSS_AMOUNT' TO campo_emp.
      ASSIGN (campo_emp) TO <gross_amount>.

      <valor_ctrl> = <gross_amount> * -1.
    ENDIF.
**********************************************************************

*stat 3 ou 4
    DATA: lv_campo(80), wa_bkpf TYPE bkpf.

    FIELD-SYMBOLS: <doc_fi>     TYPE any, <comp_code> TYPE any, <ano> TYPE any, <ref_doc_no> TYPE any.

    IF /sbxc/zckp_ctrl-status1 EQ '3' OR /sbxc/zckp_ctrl-status1 EQ '4'.
      MOVE '<CABECALHO>-DOC_FI' TO lv_campo.
      ASSIGN (lv_campo) TO <doc_fi>.

      IF <doc_fi> IS ASSIGNED AND <doc_fi> IS INITIAL.

        MOVE '<CABECALHO>-COMP_CODE' TO lv_campo.
        ASSIGN (lv_campo) TO <comp_code>.
        MOVE '<CABECALHO>-ANO' TO lv_campo.
        ASSIGN (lv_campo) TO <ano>.
        MOVE '<CABECALHO>-REF_DOC_NO' TO lv_campo.
        ASSIGN (lv_campo) TO <ref_doc_no>.

        CLEAR wa_bkpf.
        SELECT SINGLE belnr awkey gjahr cpudt FROM bkpf "#EC CI_NOORDER
       INTO (wa_bkpf-belnr, wa_bkpf-awkey, wa_bkpf-gjahr, wa_bkpf-cpudt)
      WHERE bukrs = <comp_code> AND
      gjahr = <ano> AND
      xblnr = <ref_doc_no> ##WARN_OK.

        UPDATE (str_cab) SET
                   doc_fi = wa_bkpf-belnr
                   ano_lanc = wa_bkpf-gjahr
                   data_criacao = wa_bkpf-cpudt
                   doc_lo = wa_bkpf-awkey(10)
             WHERE processo = /sbxc/zckp_ctrl-processo AND
                   ano = /sbxc/zckp_ctrl-ano AND
                   seqno = /sbxc/zckp_ctrl-seqno.

        <doc_fi> = wa_bkpf-belnr.

      ELSEIF /sbxc/zckp_ctrl-status1 EQ '6'.
        UNASSIGN <mensagem>.
        ASSIGN COMPONENT 'MENSAGEM' OF STRUCTURE <cabecalho> TO <mensagem>.
        IF <mensagem> IS ASSIGNED.
          CLEAR <mensagem>.
*          SELECT SINGLE mensagem FROM /sbxc/zckp_tab09  INTO <mensagem>
*                               WHERE processo = /sbxc/zckp_ctrl-processo
*                                 AND ano      = /sbxc/zckp_ctrl-ano
*                                 AND seqno    = /sbxc/zckp_ctrl-seqno.
          READ TABLE lt_tab09 ASSIGNING FIELD-SYMBOL(<fs_09>)
                    WITH KEY processo = /sbxc/zckp_ctrl-processo
                             ano      = /sbxc/zckp_ctrl-ano
                             seqno    = /sbxc/zckp_ctrl-seqno.
          IF sy-subrc EQ 0.
            <mensagem> = <fs_09>-mensagem.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.
    PERFORM funcao_cabecalho_ini.

*Data default como data actual apenas na entrada do cockpit
    FIELD-SYMBOLS: <status1>       TYPE any, <pstng_date> TYPE any, <xware_bnk> TYPE any,
                   <baseline_date> TYPE any, <processo> TYPE any, <seqno> TYPE any,
                   <zterm>         TYPE any, <empresa> TYPE any.

    ASSIGN COMPONENT 'STATUS1'    OF STRUCTURE <cabecalho> TO <status1>.
    ASSIGN COMPONENT 'PSTNG_DATE' OF STRUCTURE <cabecalho> TO <pstng_date>.
    ASSIGN COMPONENT 'XWARE_BNK' OF STRUCTURE <cabecalho> TO <xware_bnk>.
    ASSIGN COMPONENT 'PROCESSO'    OF STRUCTURE <cabecalho> TO <processo>.
    ASSIGN COMPONENT 'SEQNO'    OF STRUCTURE <cabecalho> TO <seqno>.
    ASSIGN COMPONENT 'ZTERM'    OF STRUCTURE <cabecalho> TO <zterm>.
    ASSIGN COMPONENT 'COMP_CODE'    OF STRUCTURE <cabecalho> TO <empresa>.
    IF <status1> = '0' OR <status1> = '5' OR /sbxc/zckp_ctrl-status1 = ' '.
      IF <pstng_date> IS ASSIGNED.
*      IF <pstng_date> IS INITIAL. "ODC - 18_03_2021
*  CCF Ini 23.02.2023 10:03:45
* 130058, Data de Lançamento | Alteração em Massa
* Pretendem que nas empresas 2* a data de lançamento assuma a que colocaram no CKP
*          IF <empresa> IS ASSIGNED AND <empresa>(1) = '2'.
*          ELSE.
        <pstng_date> = sy-datum.
*          ENDIF.
*   CCF Fim 23.02.2023 10:03:45
*       ENDIF.
      ENDIF.
    ENDIF.
    IF <xware_bnk> IS INITIAL.
      <xware_bnk> = '1'.
    ENDIF.
* CCF 02.10.2020 especifico SF calculo data base de pagamento
    DATA: processo      LIKE /sbxc/zckp_invh-processo, seqno LIKE /sbxc/zckp_invh-seqno,
          baseline_date LIKE /sbxc/zckp_invh-baseline_date.

    processo = <processo>.
    seqno = <seqno>.

    CALL FUNCTION '/SBXC/ZCKP_BASELINE_DATE'
      EXPORTING
        processo      = processo
        seqno         = seqno
      CHANGING
        baseline_date = baseline_date.

    ASSIGN COMPONENT 'BASELINE_DATE' OF STRUCTURE <cabecalho> TO <baseline_date>.
    <baseline_date> = baseline_date.

** CCF 16.11.2021
** Se a condição de pagamento tem o campo “Dia fixo” preenchido (T052 - ZFAEL), o sistema faz EOMONTH da “data base”.
*  DATA: zfael    LIKE t052-zfael,
*        lv_ldate TYPE sy-datum.
*
*  SELECT SINGLE zfael INTO zfael FROM t052 WHERE zterm = <ZTERM>.
*  IF zfael = '31'.
*    CALL FUNCTION 'LAST_DAY_OF_MONTHS'
*      EXPORTING
*        day_in            =  baseline_date
*      IMPORTING
*        last_day_of_month = lv_ldate.
**  EXCEPTIONS
**       DAY_IN_NO_DATE    = 1
**       OTHERS            = 2
*   ENDIF   .

* fim 02.10.2021
* Status pode ter sido actualizado
    SELECT SINGLE status1
      INTO /sbxc/zckp_ctrl-status1
      FROM /sbxc/zckp_ctrl
      WHERE processo = /sbxc/zckp_ctrl-processo
      AND ano = /sbxc/zckp_ctrl-ano
      AND seqno = /sbxc/zckp_ctrl-seqno.
    MOVE-CORRESPONDING /sbxc/zckp_ctrl  TO <cabecalho>.
    ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <cabecalho> TO <fs_style>.
    PERFORM estilo_campos USING 'CAB' <status_cab>.
    ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <cabecalho> TO <fs_color>.

    PERFORM cor_campos USING 'CAB' '1'.

    FIELD-SYMBOLS: <vendor> TYPE any, <name1> TYPE any.
    ASSIGN COMPONENT 'VENDOR' OF STRUCTURE <cabecalho> TO <vendor>.
    ASSIGN COMPONENT 'NAME1' OF STRUCTURE <cabecalho> TO <name1>.
    IF <vendor> IS ASSIGNED.
      IF <vendor> IS NOT INITIAL AND <name1> IS INITIAL.
        SELECT SINGLE name1 INTO <name1>
          FROM lfa1
          WHERE lifnr EQ <vendor>.
      ENDIF.
    ENDIF.
    APPEND <cabecalho> TO <t_cab_table> .

  ENDLOOP.

  DATA: otab   TYPE abap_sortorder_tab,
        l_line TYPE abap_sortorder.
  l_line-name = 'PROCESSO'.
  APPEND l_line TO otab.
  l_line-name = 'ANO'.
  APPEND l_line TO otab.
  l_line-name = 'SEQNO'.
  APPEND l_line TO otab.
  TRY.
      SORT <t_cab_table> BY (otab).
    CATCH cx_sy_dyn_table_ill_comp_val.
      MESSAGE i084(/sbxc/zckp_cockpit) DISPLAY LIKE 'E'.
      LEAVE PROGRAM.
  ENDTRY.
ENDFORM.                    " ESTRUTURA_CAB
*&---------------------------------------------------------------------*
*&      Form  FUNCAO_CABECALHO_INI
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM funcao_cabecalho_ini .

  READ TABLE it_fm_display INTO wa_fm_display WITH KEY processo = /sbxc/zckp_ctrl-processo.


  IF wa_fm_display-fm_cab_disp NE space AND
       sy-subrc EQ 0.

    CALL FUNCTION 'FUNCTION_EXISTS' ##FM_SUBRC_OK
      EXPORTING
        funcname           = wa_fm_display-fm_cab_disp
      EXCEPTIONS
        function_not_exist = 1
        OTHERS             = 2.
    REFRESH f_cor.

    IF sy-subrc EQ 0.
      CALL FUNCTION wa_fm_display-fm_cab_disp
        EXPORTING
          est_cab = /sbxc/zckp_tab00-est_cab
          t_tab12 = lt_tab12
          t_hmail = lt_hmail
        TABLES
          cor     = f_cor
        CHANGING
          cab     = <cabecalho>.
    ENDIF.

  ENDIF.
ENDFORM.                    " FUNCAO_CABECALHO_INI
*&---------------------------------------------------------------------*
*&      Form  ESTILO_CAMPOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_1212   text
*----------------------------------------------------------------------*
FORM estilo_campos USING cab_lin TYPE any status TYPE any.
  FIELD-SYMBOLS <barcode_cab> TYPE any.
  DATA l_barcode_cab(30).
  l_barcode_cab = '<CAB>-BARCODE'.
  ASSIGN (l_barcode_cab) TO <barcode_cab>.

  FIELD-SYMBOLS <trat_cab> TYPE any.
  DATA l_trat_cab(30).

  IF <cab> IS ASSIGNED.
    l_trat_cab = '<CAB>-USER_TRATAMENTO'.
    ASSIGN (l_trat_cab) TO <trat_cab>.
  ELSE.
    l_trat_cab = '<CABECALHO>-USER_TRATAMENTO'.
    ASSIGN (l_trat_cab) TO <trat_cab>.
  ENDIF.

  REFRESH it_edit.

  DATA: tabela TYPE dd03l-tabname.

  IF cab_lin = 'CAB'.
    tabela = str_cab.
  ELSE.
    tabela = str_lin.
  ENDIF.


  LOOP AT gt_dd03l ASSIGNING FIELD-SYMBOL(<fs_dd031>) WHERE tabname = tabela.

    wa_edit-fieldname = <fs_dd031>-fieldname.

    READ TABLE gt_tab10 ASSIGNING FIELD-SYMBOL(<fs_tab10>) WITH KEY status = status.
    IF sy-subrc EQ 0 AND <fs_tab10>-edit EQ 'X'.

      READ TABLE gt_tab04 ASSIGNING FIELD-SYMBOL(<fs_tab04>) WITH KEY processo = wa_fm_display-processo
      estrutura  = cab_lin
      campo = <fs_dd031>-fieldname.

      IF sy-subrc = 0.

* Para cada campo editável verificar se utilizador actual pode editar
        READ TABLE gt_tab05 ASSIGNING FIELD-SYMBOL(<fs_tab05>) WITH KEY processo   = <fs_tab04>-processo
                       uname      = v_uname
                       estrutura  = cab_lin
                       campo      = <fs_tab04>-campo.

        IF sy-subrc = 0.
*Caso seja factura electronica (Sem BARCODE) nao permitir modificar quantidades e valores.
          IF <barcode_cab> IS ASSIGNED.
            IF ( wa_edit-fieldname  =  'ITEM_AMOUNT' OR
               wa_edit-fieldname  =  'QUANT_FACT'  OR
               wa_edit-fieldname  =  'QUANTITY'    OR
               wa_edit-fieldname  =  'ITEM_AMOUNT_FORN' )
          AND <barcode_cab> = ''.
              wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
              CLEAR wa_edit-style.
            ELSE.
              IF status = '9'.
                IF wa_edit-fieldname = 'MOT_N_CONTAB'.
                  wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
                ELSE.
                  wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
                ENDIF.
              ELSE.
                wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
              ENDIF.
            ENDIF.
          ELSE.
            IF status = '9'.
              IF wa_edit-fieldname = 'MOT_N_CONTAB'.
                wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
              ELSE.
                wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
              ENDIF.
            ELSE.
              wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
            ENDIF.
          ENDIF.
*          IF ( wa_edit-fieldname  =  'ITEM_AMOUNT' OR
*               wa_edit-fieldname  =  'QUANT_FACT'  OR
*               wa_edit-fieldname  =  'QUANTITY'    OR
*               wa_edit-fieldname  =  'ITEM_AMOUNT_FORN' )
*          AND <barcode_cab> = ''.
*            wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
*            CLEAR wa_edit-style.
*          ELSE.
*            IF status = '9'.
*              IF wa_edit-fieldname = 'MOT_N_CONTAB'.
*                wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
*              ELSE.
*                wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
*              ENDIF.
*            ELSE.
*              wa_edit-style = cl_gui_alv_grid=>mc_style_enabled.
*            ENDIF.
*          ENDIF.
          IF <trat_cab> IS ASSIGNED.
            IF <trat_cab> NE v_uname AND <trat_cab> IS NOT INITIAL.
              wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
            ENDIF.
          ENDIF.

        ELSE.
          wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.

        ENDIF.
      ELSE.
        wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
      ENDIF.

    ELSE.
      wa_edit-style = cl_gui_alv_grid=>mc_style_disabled.
    ENDIF.
    INSERT wa_edit INTO TABLE it_edit.

    IF <fs_style> IS ASSIGNED.
      <fs_style> = it_edit.
    ENDIF.

  ENDLOOP.

ENDFORM.                    " ESTILO_CAMPOS
*&---------------------------------------------------------------------*
*&      Form  COR_CAMPOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_1226   text
*      -->P_1227   text
*----------------------------------------------------------------------*
FORM cor_campos  USING   cab_lin TYPE any index TYPE any.
  REFRESH it_color.

*  DATA: tabela TYPE dd03l-tabname.
*
*  IF cab_lin = 'CAB'.
*    tabela = str_cab.
*  ELSE.
*    tabela = str_lin.
*  ENDIF.

  REFRESH <fs_color>.
  LOOP AT f_cor ASSIGNING FIELD-SYMBOL(<fs_cor>) WHERE tabix = index.  " Verificar
    wa_color-fname = <fs_cor>-fname.

    wa_cor-col = <fs_cor>-col.
    wa_cor-int = <fs_cor>-int.
    wa_cor-inv = <fs_cor>-inv.
    MOVE wa_cor TO wa_color-color.

    INSERT wa_color INTO TABLE it_color.
    <fs_color> = it_color.
  ENDLOOP.

  REFRESH f_cor.

ENDFORM.                    " COR_CAMPOS
*&---------------------------------------------------------------------*
*&      Form  TEXTO_CAMPO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_DD03L_TABNAME  text
*      -->P_DD03L_FIELDNAME  text
*      <--P_WA_FIELDCAT_COLTEXT  text
*      <--P_WA_FIELDCAT_SELTEXT  text
*      <--P_WA_FIELDCAT_SCRTEXT_L  text
*      <--P_WA_FIELDCAT_SCRTEXT_M  text
*      <--P_WA_FIELDCAT_SCRTEXT_S  text
*----------------------------------------------------------------------*
FORM texto_campo  USING    tabname TYPE dd03l-tabname
                           fieldname TYPE dd03l-fieldname
                  CHANGING coltext TYPE lvc_s_fcat-coltext
                           seltext TYPE lvc_s_fcat-seltext
                           scrtext_l TYPE lvc_s_fcat-scrtext_l
                           scrtext_m TYPE lvc_s_fcat-scrtext_m
                           scrtext_s TYPE lvc_s_fcat-scrtext_s.

  DATA: campos TYPE TABLE OF dfies.

  CLEAR: coltext, seltext, scrtext_l, scrtext_m, scrtext_s.

  CALL FUNCTION 'DDIF_FIELDINFO_GET'
    EXPORTING
      tabname        = tabname
      fieldname      = fieldname
    TABLES
      dfies_tab      = campos
    EXCEPTIONS
      not_found      = 1
      internal_error = 2
      OTHERS         = 3.
  IF sy-subrc <> 0.
*    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    RETURN.
  ENDIF.

  LOOP AT campos ASSIGNING FIELD-SYMBOL(<fs_camp>).
    coltext   = <fs_camp>-scrtext_s .
    seltext   = <fs_camp>-scrtext_s .
    scrtext_l = <fs_camp>-scrtext_l .
    scrtext_m = <fs_camp>-scrtext_m .
    scrtext_s = <fs_camp>-scrtext_s .
  ENDLOOP.
ENDFORM.                    " TEXTO_CAMPO
*&---------------------------------------------------------------------*
*&      Form  TOOLBAR_EXCLUDING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_LT_UI_FUNCT  text
*----------------------------------------------------------------------*
FORM toolbar_excluding  CHANGING  pt_funct TYPE ui_functions.

  DATA: ls_ui_funct TYPE ui_func.


  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_cut.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_graph.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_check.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_cut.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_insert_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_delete_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_append_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste_new_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_undo.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_refresh.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_info.
  APPEND ls_ui_funct TO pt_funct.

ENDFORM.                    " TOOLBAR_EXCLUDING
*&---------------------------------------------------------------------*
*&      Form  ESTRUTURA_LIN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM estrutura_lin .

  DATA: wa_dd03l TYPE dd03l.

  SELECT * INTO TABLE @DATA(lt_tab04)
     FROM  /sbxc/zckp_tab04
" Campo com permissão de edição ?
       WHERE  processo   = @/sbxc/zckp_tab00-processo
       AND    estrutura  = 'LIN'.

  SELECT * INTO TABLE @DATA(lt_tab05) FROM  /sbxc/zckp_tab05
      WHERE  processo   = @/sbxc/zckp_tab00-processo
      AND    uname      = @v_uname
      AND    estrutura  = 'LIN'.

  CLEAR pos.
* Campos da estrutura definida para cabeçalho
  IF  /sbxc/zckp_tab00-est_cab NE space.
    SELECT        * FROM  dd03l INTO wa_dd03l

           WHERE  tabname  = /sbxc/zckp_tab00-est_lin
            AND comptype NE 'S'
            AND   fieldname NE  'MANDT'
                  ORDER BY position .

      CLEAR wa_fieldcat.

      PERFORM texto_campo USING wa_dd03l-tabname
                                wa_dd03l-fieldname
                          CHANGING wa_fieldcat-coltext
                             wa_fieldcat-seltext
                             wa_fieldcat-scrtext_l
                             wa_fieldcat-scrtext_m
                             wa_fieldcat-scrtext_s.

      wa_fieldcat-fieldname = wa_dd03l-fieldname.
      wa_fieldcat-inttype   = wa_dd03l-inttype.
      wa_fieldcat-outputlen = wa_dd03l-intlen.
      wa_fieldcat-ref_table = wa_dd03l-tabname.
      wa_fieldcat-ref_field = wa_dd03l-fieldname.
      wa_fieldcat-rollname = wa_dd03l-rollname.
      IF wa_fieldcat-fieldname EQ 'DB_CR_IND'.
        wa_fieldcat-f4availabl = 'X'.
      ENDIF.

      READ TABLE lt_tab04 ASSIGNING FIELD-SYMBOL(<fs_04>) WITH KEY campo = wa_dd03l-fieldname.
*      SELECT SINGLE * FROM  /sbxc/zckp_tab04
*      " Campo com permissão de edição ?
*             WHERE  processo   = /sbxc/zckp_tab00-processo
*             AND    estrutura  = 'LIN'
*             AND    campo = dd03l-fieldname.

      IF sy-subrc = 0.
* Para cada campo editável verificar se utilizador actual pode editar
        READ TABLE lt_tab05 ASSIGNING FIELD-SYMBOL(<fs_05>) WITH KEY campo = <fs_04>-campo.
*        SELECT SINGLE * FROM  /sbxc/zckp_tab05 CLIENT SPECIFIED
*               WHERE  mandt      = sy-mandt
*               AND    processo   = /sbxc/zckp_tab00-processo
*               AND    uname      = v_uname
*               AND    estrutura  = 'LIN'
*               AND    campo      = /sbxc/zckp_tab04-campo.

        IF sy-subrc = 0.
          wa_fieldcat-edit = 'X'.
        ELSE.
          wa_fieldcat-edit = ' '.
        ENDIF.

      ELSE.
        wa_fieldcat-edit = ' '.
      ENDIF.

      IF wa_fieldcat-fieldname  = 'TAX_AMOUNT'.
        wa_fieldcat-decimals_o = '2'.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'TAX_IMPOSTO_SAP'.
        wa_fieldcat-scrtext_l = TEXT-050.
        wa_fieldcat-scrtext_m = TEXT-050.
        wa_fieldcat-scrtext_s = TEXT-050.
        wa_fieldcat-coltext = TEXT-050.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'TAX_AMOUNT'.
        wa_fieldcat-scrtext_l = TEXT-051.
        wa_fieldcat-scrtext_m = TEXT-051.
        wa_fieldcat-scrtext_s = TEXT-051.
        wa_fieldcat-coltext = TEXT-051.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'PO_ITEM'.
        wa_fieldcat-scrtext_l = TEXT-052.
        wa_fieldcat-scrtext_m = TEXT-052.
        wa_fieldcat-scrtext_s = TEXT-052.
        wa_fieldcat-coltext = TEXT-052.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'REF_DOC_ITEM'.
        wa_fieldcat-scrtext_l = TEXT-053.
        wa_fieldcat-scrtext_m = TEXT-053.
        wa_fieldcat-scrtext_s = TEXT-053.
        wa_fieldcat-coltext = TEXT-053.
      ENDIF.

      IF wa_fieldcat-fieldname EQ 'ITEM_SAPHETY'.
        wa_fieldcat-scrtext_l = TEXT-054.
        wa_fieldcat-scrtext_m = TEXT-054.
        wa_fieldcat-scrtext_s = TEXT-054.
        wa_fieldcat-coltext = TEXT-054.
      ENDIF.

      pos = pos + 1.
      wa_fieldcat-col_pos = pos.
      IF wa_dd03l-domname = 'XFELD'.
        wa_fieldcat-checkbox = 'X'.
      ENDIF.
      APPEND wa_fieldcat TO t_fieldcat_lin.

    ENDSELECT.


* Campos para controlo de edição
    wa_fieldcat-coltext = ' '.
    wa_fieldcat-seltext = ' '.
    wa_fieldcat-scrtext_l = ' '.
    wa_fieldcat-scrtext_m = ' '.
    wa_fieldcat-scrtext_s = ' '.
    wa_fieldcat-fieldname = 'CELLSTYLE'.
    wa_fieldcat-ref_table = '/SBXC/ZCKP_ALV'.
    wa_fieldcat-ref_field = 'CELL_STYLE'.
    APPEND wa_fieldcat TO t_fieldcat_lin.
    CLEAR wa_fieldcat.



    wa_fieldcat-coltext = ' '.
    wa_fieldcat-seltext = ' '.
    wa_fieldcat-scrtext_l = ' '.
    wa_fieldcat-scrtext_m = ' '.
    wa_fieldcat-scrtext_s = ' '.
    wa_fieldcat-fieldname = 'CELL_COLOR'.
    wa_fieldcat-ref_table = '/SBXC/ZCKP_ALV'.
    wa_fieldcat-ref_field = 'CELL_COLOR'.


    APPEND wa_fieldcat TO t_fieldcat_lin.
    CLEAR wa_fieldcat.

**** AFG-ID-2500006 *****

** Criar tabela a partir do field catalog
*    CALL METHOD cl_alv_table_create=>create_dynamic_table
*      EXPORTING
*        it_fieldcatalog           = t_fieldcat_lin
*      IMPORTING
*        ep_table                  = r_dyn_table_lin
*      EXCEPTIONS
*        generate_subpool_dir_full = 1
*        OTHERS                    = 2.
*
*
*    IF sy-subrc <> 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
*    ENDIF.

** Apontador para a tabela
*    ASSIGN r_dyn_table_lin->* TO <t_lin_table>.
*
** Criar workarea para a tabela
*    CREATE DATA r_wa_dyn_table_lin LIKE LINE OF <t_lin_table>.
*
** Apontador para a wa
*    ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.



*  perform estrutura_lin.
    ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
                <fs_style>.

    PERFORM estilo_campos USING 'LIN' /sbxc/zckp_ctrl-status1 .

** Tabela de apoio
*    CALL METHOD cl_alv_table_create=>create_dynamic_table
*      EXPORTING
*        it_fieldcatalog           = t_fieldcat_lin
*      IMPORTING
*        ep_table                  = r_dyn_table_lin2
*      EXCEPTIONS
*        generate_subpool_dir_full = 1
*        OTHERS                    = 2.
*
*    IF sy-subrc <> 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
*    ENDIF.
*
** Apontador para a tabela
*    ASSIGN r_dyn_table_lin2->* TO <t_lin_table2>.
*
** Criar workarea para a tabela
*    CREATE DATA r_wa_dyn_table_lin2 LIKE LINE OF <t_lin_table2>.



  LOOP AT t_fieldcat_lin INTO wa_fieldcat.
    IF wa_fieldcat-fieldname EQ 'CELL_COLOR'.
      ls_component-name = 'CELL_COLOR'.
      ls_component-type ?= cl_abap_datadescr=>describe_by_name( '/SBXC/ZCKP_ALV-CELL_COLOR' ).
    ELSEIF wa_fieldcat-fieldname EQ 'CELLSTYLE'.
      ls_component-name = 'CELLSTYLE'.
      ls_component-type ?= cl_abap_datadescr=>describe_by_name( '/SBXC/ZCKP_ALV-CELL_STYLE' ).
    ELSE.
*   Element Description
      lo_element ?= cl_abap_elemdescr=>describe_by_name( wa_fieldcat-rollname ).
      ls_component-type = lo_element.
*   Field name
      LS_COMPONENT-name = wa_fieldcat-fieldname.
    ENDIF.

*   Filling the component table
    APPEND ls_component TO lt_comp_lin.
    CLEAR: ls_component.
  ENDLOOP.

  gr_struct_typ  ?= cl_abap_structdescr=>create( p_components = lt_comp_lin ).
  gr_dyntable_typ = cl_abap_tabledescr=>create( p_line_type = gr_struct_typ ).

  CREATE DATA r_dyn_table_lin TYPE HANDLE gr_dyntable_typ.
  CREATE DATA  r_wa_dyn_table_lin TYPE HANDLE gr_struct_typ.


  gr_struct_typ_tmp  ?= cl_abap_structdescr=>create( p_components = lt_comp_lin ).
  gr_dyntable_typ_tmp = cl_abap_tabledescr=>create( p_line_type = gr_struct_typ ).

  CREATE DATA r_dyn_table_lin2 TYPE HANDLE gr_dyntable_typ_tmp.
  CREATE DATA  r_wa_dyn_table_lin2 TYPE HANDLE gr_struct_typ_tmp.


  ASSIGN r_dyn_table_lin->* TO <t_lin_table>.
  ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
  ASSIGN r_dyn_table_lin2->* TO <t_lin_table2>.

**********************************************************************

  ENDIF.

ENDFORM.                    " ESTRUTURA_LIN
*&---------------------------------------------------------------------*
*&      Form  TOOLBAR_EXCLUDING_LIN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_LT_UI_FUN_LIN  text
*----------------------------------------------------------------------*
FORM toolbar_excluding_lin  CHANGING  pt_funct TYPE ui_functions.

  DATA: ls_ui_funct TYPE ui_func.


  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_cut.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_graph.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_check.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_cut.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_insert_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_delete_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_append_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste_new_row.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_undo.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_refresh.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_subtot.
  APPEND ls_ui_funct TO pt_funct.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_info.
  APPEND ls_ui_funct TO pt_funct.


ENDFORM.                    " TOOLBAR_EXCLUDING_LIN
*&---------------------------------------------------------------------*
*&      Form  DOUBLE_CLICK
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_E_COLUMN  text
*----------------------------------------------------------------------*
FORM double_click USING e_column TYPE any e_row TYPE any.
  DATA field TYPE dd03l-fieldname.
  field = e_column.

  IF wa_fm_display-fm_dblclk NE space.

    CALL FUNCTION 'FUNCTION_EXISTS'
      EXPORTING
        funcname           = wa_fm_display-fm_dblclk
      EXCEPTIONS
        function_not_exist = 1
        OTHERS             = 2.

    IF sy-subrc = 0.

**** Análise 1
      CALL FUNCTION wa_fm_display-fm_dblclk
        EXPORTING
          anexo_on = disp_doc_active
          campo    = field
        IMPORTING
          refresh  = refresh_table
        TABLES
          cor      = f_cor       " Cor para células
          linha    = <t_lin_table>
        CHANGING
          cab      = <cab>.

    ENDIF.

    CALL METHOD go_grid_cab->refresh_table_display(
        is_stable      = gc_stable
        i_soft_refresh = 'X' ).
    CALL METHOD go_grid_lin->refresh_table_display(
        is_stable      = gc_stable
        i_soft_refresh = 'X' ).

  ENDIF.

ENDFORM.                    " DOUBLE_CLICK
*&---------------------------------------------------------------------*
*&      Form  DEFINE_FCODE_TABLES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_STD_FCODES  text
*      -->P_LT_OWN_FCODES  text
*----------------------------------------------------------------------*
FORM define_fcode_tables
           TABLES   lt_std_fcodes TYPE ui_functions
                    lt_own_fcodes TYPE ui_functions.
  APPEND 'MULT' TO lt_own_fcodes.

ENDFORM.                    " DEFINE_FCODE_TABLES
*&---------------------------------------------------------------------*
*&      Form  ADICIONA_OPCOES_H
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_E_OBJECT  text
*----------------------------------------------------------------------*
FORM adiciona_opcoes_h  USING p_e_object TYPE REF TO cl_ctmenu.

  DATA: fcode        TYPE  ui_func,
        text         TYPE	gui_text,
        function_tmp TYPE /sbxc/zckp_tab03-function.

  SORT it_botoes BY function.
  LOOP AT it_botoes ASSIGNING FIELD-SYMBOL(<fs_botoes>).

    CHECK function_tmp NE <fs_botoes>-function.

    CLEAR: fcode, text.
    MOVE: <fs_botoes>-function TO fcode,
          <fs_botoes>-tooltip TO text.

    CALL METHOD p_e_object->add_function
      EXPORTING
        fcode = fcode
        text  = text.

    function_tmp = <fs_botoes>-function.

  ENDLOOP.

ENDFORM.                    " ADICIONA_OPCOES_H
*&---------------------------------------------------------------------*
*&      Form  ADICIONA_OPCOES_L
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_E_OBJECT  text
*----------------------------------------------------------------------*
FORM adiciona_opcoes_l  USING p_e_object TYPE REF TO cl_ctmenu.

  DATA: fcode        TYPE  ui_func,
        text         TYPE	gui_text,
        function_tmp LIKE /sbxc/zckp_tab03-function.

  SORT it_botoes_l BY function.
  LOOP AT it_botoes_l ASSIGNING FIELD-SYMBOL(<fs_botoes>).

    CHECK function_tmp NE <fs_botoes>-function.

    CLEAR: fcode, text.
    MOVE: <fs_botoes>-function TO fcode,
          <fs_botoes>-tooltip TO text.

    CALL METHOD p_e_object->add_function
      EXPORTING
        fcode = fcode
        text  = text.

    function_tmp = <fs_botoes>-function.

  ENDLOOP.

ENDFORM.                    " ADICIONA_OPCOES_L
*&---------------------------------------------------------------------*
*&      Form  GRAVA_ALTERACOES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM grava_alteracoes .

  DATA: var1(40),        var2(40).
  FIELD-SYMBOLS: <f1> TYPE any, <f2> TYPE any.


  ASSIGN r_wa_dyn_table_cab->* TO  <wa_cab_table> .

  LOOP  AT <t_cab_table> ASSIGNING <wa_cab_table>.
    var1 = '<WA_CAB_TABLE>-PROCESSO'.
    ASSIGN (var1) TO <f1>.
    var2 = '<WA_CAB_TABLE>-SEQNO'.
    ASSIGN (var2) TO <f2>.

    SELECT   SINGLE  fm_alteracao FROM  /sbxc/zckp_tab00
           INTO /sbxc/zckp_tab00-fm_alteracao
           WHERE  processo = <f1>.


    CALL FUNCTION 'FUNCTION_EXISTS'
      EXPORTING
        funcname           = /sbxc/zckp_tab00-fm_alteracao
      EXCEPTIONS
        function_not_exist = 1
        OTHERS             = 2.

    CHECK sy-subrc = 0.

    CALL FUNCTION /sbxc/zckp_tab00-fm_alteracao
      EXPORTING
        cab   = <wa_cab_table>
      TABLES
        linha = <t_lin_table>.

  ENDLOOP.

  CLEAR gd_changes.


**********************************************************************
  IF lin_or_cab NE 'STA'.
    CLEAR: <t_cab_table_tmp>, <t_cab_table_tmp>[].
    <t_cab_table_tmp>[] = <t_cab_table>[].
  ENDIF.
**********************************************************************


ENDFORM.                    " GRAVA_ALTERACOES
*&---------------------------------------------------------------------*
*&      Form  FUNCAO_CABECALHO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM funcao_cabecalho .

  READ TABLE it_fm_display INTO wa_fm_display WITH KEY processo = key-processo.


  IF wa_fm_display-fm_cab_disp NE space AND
       sy-subrc EQ 0.

    CALL FUNCTION 'FUNCTION_EXISTS'
      EXPORTING
        funcname           = wa_fm_display-fm_cab_disp
      EXCEPTIONS
        function_not_exist = 1
        OTHERS             = 2.

    IF sy-subrc = 0.
      IF <cab> IS ASSIGNED.
        CALL FUNCTION wa_fm_display-fm_cab_disp
          EXPORTING
            est_cab = /sbxc/zckp_tab00-est_cab
            t_tab12 = lt_tab12
            t_hmail = lt_hmail
          TABLES
            cor     = f_cor       " Cor para células
            linha   = <t_lin_table>
          CHANGING
            cab     = <cab>.
      ENDIF.
    ENDIF.

  ENDIF.

ENDFORM.                    " FUNCAO_CABECALHO

*&---------------------------------------------------------------------*
*&      Form  EXIT_PROGRAM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM exit_program .

  DATA: l_answer TYPE c.

  IF gd_changes EQ 'X'.
    CALL FUNCTION 'POPUP_TO_CONFIRM' ##FM_SUBRC_OK
      EXPORTING
        titlebar       = TEXT-p01 "Aviso
        text_question  = TEXT-p02 "Gravar possíveis alterações?
        text_button_1  = TEXT-p03 "Sim
        default_button = '1'
      IMPORTING
        answer         = l_answer
      EXCEPTIONS
        text_not_found = 1
        OTHERS         = 2.

    CASE l_answer.
      WHEN '1'.
        CLEAR gd_changes.

        PERFORM grava_alteracoes.

    ENDCASE.
  ENDIF.

  SET SCREEN 0.LEAVE SCREEN.

ENDFORM.                    " EXIT_PROGRAM
*&---------------------------------------------------------------------*
*&      Form  VALIDA_FUNCOES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_FUNCAO_PRE  text
*      -->P_FUNCAO  text
*      -->P_FUNCAO_POS  text
*----------------------------------------------------------------------*
FORM valida_funcoes  CHANGING   funcao_testar TYPE rs38l-name .

* Verificar se função existe
  CALL FUNCTION 'FUNCTION_EXISTS'
    EXPORTING
      funcname           = funcao_testar
    EXCEPTIONS
      function_not_exist = 1
      OTHERS             = 2.

  IF sy-subrc NE 0.

    CLEAR funcao_testar.

  ENDIF.

ENDFORM.                    " VALIDA_FUNCOES
*&---------------------------------------------------------------------*
*&      Form  GET_TAB_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM get_tab_status .

  DATA: var5(40).
  FIELD-SYMBOLS: <f5> TYPE any.
  CLEAR: gt_status, gt_status[], wa_status.

  ASSIGN r_wa_dyn_table_cab->* TO <cab>.

  LOOP AT gt_tab10 ASSIGNING FIELD-SYMBOL(<fs_tab10>).
    wa_status-status = <fs_tab10>-status.
    wa_status-icon_status = <fs_tab10>-icon_status.

    CLEAR: wa_status-box, wa_status-cont.
    READ TABLE gt_tab10t ASSIGNING FIELD-SYMBOL(<fs_tab10t>) WITH KEY spras = sy-langu
    status = wa_status-status.

    IF sy-subrc EQ 0.
      wa_status-text_status = <fs_tab10t>-text_status.
    ENDIF.

    LOOP AT <t_cab_table> ASSIGNING <cab>.
      var5 = '<CAB>-STATUS1'.
      ASSIGN (var5) TO <f5>.
      IF <f5> IS ASSIGNED.
        CHECK <f5> EQ wa_status-status.
        ADD 1 TO wa_status-cont.
      ENDIF.
    ENDLOOP.

    IF wa_status-cont IS NOT INITIAL.
      wa_status-box = 'X'.
    ENDIF.

    APPEND wa_status TO gt_status.
  ENDLOOP.

  SORT gt_status.
  DELETE ADJACENT DUPLICATES FROM gt_status.
  SORT gt_status BY status.

ENDFORM.                    " GET_TAB_STATUS
*&---------------------------------------------------------------------*
*&      Form  GET_FIELDCAT_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM get_fieldcat_status .

*Fieldcatalog
  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
      i_structure_name       = '/SBXC/ZCKP_TAB_STATUS'
      i_client_never_display = 'X'
    CHANGING
      ct_fieldcat            = t_fieldcat_status[]
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

  CLEAR wa_fieldcat.
  LOOP AT t_fieldcat_status INTO wa_fieldcat.
    IF wa_fieldcat-fieldname = 'BOX'.
      wa_fieldcat-checkbox = 'X'.
      wa_fieldcat-hotspot = 'X'.
      wa_fieldcat-edit = 'X'.
    ENDIF.
    MODIFY t_fieldcat_status FROM wa_fieldcat.
  ENDLOOP.

ENDFORM.                    " GET_FIELDCAT_STATUS
*&---------------------------------------------------------------------*
*&      Form  TOOLBAR_EXCLUDING_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_LT_UI_FUNCT_STATUS  text
*----------------------------------------------------------------------*
FORM toolbar_excluding_status  CHANGING pt_funct2 TYPE ui_functions.

  DATA: ls_ui_funct TYPE ui_func.

  ls_ui_funct = cl_gui_alv_grid=>mc_fc_auf.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_average.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_back_classic.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_abc.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_chain.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_crbatch.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_crweb.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_lineitems.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_master_data.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_more.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_report.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_xint.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_call_xxl.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_check.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_col_invisible.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_col_optimize.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_count.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_current_variant.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_data_save.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_delete_filter.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_deselect_all.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_detail.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_excl_all.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_expcrdata.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_expcrdesig.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_expcrtempl.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_expmdb.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_extend.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_f4.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_filter.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_find.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_fix_columns.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_graph.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_help.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_html.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_info.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_load_variant.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_append_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_copy_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_cut.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_delete_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_insert_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_move_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_paste_new_row.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_loc_undo.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_maintain_variant.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_maximum.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_minimum.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_pc_file.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_print.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_print_back.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_print_prev.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_refresh.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_reprep.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_save_variant.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_select_all.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_send.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_separator.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_sort.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_sort_asc.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_sort_dsc.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_subtot.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_sum.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_to_office.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_to_rep_tree.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_unfix_columns.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_url_copy_to_clipboard.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_variant_admin.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_view_crystal.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_view_excel.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_view_grid.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_view_lotus.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fc_word_processor.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_export.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_filter.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_paste.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_subtot.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_sum.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_variant.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_mb_view.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fg_sort.
  APPEND ls_ui_funct TO pt_funct2.
  ls_ui_funct = cl_gui_alv_grid=>mc_fg_edit.
  APPEND ls_ui_funct TO pt_funct2.

ENDFORM.                    " TOOLBAR_EXCLUDING_STATUS
*&---------------------------------------------------------------------*
*&      Form  DS_CLAUSES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_DS_CLAUSES  text
*----------------------------------------------------------------------*
FORM ds_clauses  CHANGING p_ds_clauses TYPE rsds_where.
  DATA: wa_dd03l TYPE dd03l." Campos de tabelas
  DATA: ls_bukrs LIKE LINE OF p_bukrs.

  SELECT SINGLE * FROM dd03l INTO wa_dd03l              "#EC CI_NOORDER
    WHERE tabname  = str_cab
    AND fieldname EQ 'COMP_CODE' ##WARN_OK.

  IF sy-subrc EQ 0.
    CLEAR: option, low.
    LOOP AT p_bukrs INTO ls_bukrs.

      AT FIRST.
        line_tmp = '( '.
        CONDENSE line_tmp.
      ENDAT.

      CLEAR: line, line_tmp.
      option = ls_bukrs-option.
      low = ls_bukrs-low.

      CONCATENATE '''' low '''' INTO line_tmp.
      CONDENSE line_tmp.

      IF p_ds_clauses-where_tab[] IS INITIAL.
        CONCATENATE '( COMP_CODE ' option line_tmp ')'
        INTO line SEPARATED BY space.
      ELSE.

        CONCATENATE 'OR' '( COMP_CODE ' option line_tmp ')'

        INTO line SEPARATED BY space.
      ENDIF.

      APPEND line TO p_ds_clauses-where_tab.

      AT LAST.
        line_tmp = ' )'.
        CONDENSE line_tmp.
      ENDAT.

    ENDLOOP.
  ENDIF.

ENDFORM.                    " DS_CLAUSES
*&---------------------------------------------------------------------*
*&      Form  GET_CONFIGURATION
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM get_configuration .

  DATA: wa_tab10 TYPE /sbxc/zckp_tab10.

  DATA: wa_status LIKE LINE OF p_status.
  CLEAR: gt_dd03l, gt_dd03l[],gt_tab04,gt_tab04[],
         gt_tab05, gt_tab05[].

  SELECT * FROM dd03l
    INTO TABLE gt_dd03l WHERE tabname = str_cab
    AND as4local = 'A'.

  SELECT * FROM dd03l
APPENDING CORRESPONDING FIELDS OF TABLE gt_dd03l WHERE tabname = str_lin
AND as4local = 'A'.

  SELECT * FROM dd03l
APPENDING CORRESPONDING FIELDS OF TABLE gt_dd03l WHERE tabname = '/SBXC/ZCKP_CTRL'
AND as4local = 'A'.

  CALL FUNCTION '/SBXC/ZCKP_LOAD_DD03L'
    TABLES
      lt_dd03l = gt_dd03l.
* Tab04 - Campos editáveis
  SELECT * FROM  /sbxc/zckp_tab04
    INTO TABLE gt_tab04
           WHERE  processo   IN p_proc.
* Utilizadores que podem ter os campos editaveis
  SELECT * FROM  /sbxc/zckp_tab05
    INTO TABLE gt_tab05
         WHERE processo   IN p_proc
         AND uname      = v_uname.
* Tab010 - Cores dos status
  SELECT * FROM /sbxc/zckp_tab10
    INTO TABLE gt_tab10
    WHERE processo IN p_proc.

  CALL FUNCTION '/SBXC/ZCKP_LOAD_TAB10'
    TABLES
      lt_tab10 = gt_tab10.
* Status - Textos
  SELECT * FROM /sbxc/zckptab10t
    INTO TABLE gt_tab10t.


  SELECT * FROM /sbxc/zckp_tab10 INTO wa_tab10
    WHERE processo IN p_proc
    AND status_out IN p_stout.
    wa_status-sign = 'I'.
    wa_status-option = 'EQ'.
    wa_status-low = wa_tab10-status.
    APPEND wa_status TO p_status.
  ENDSELECT.

ENDFORM.                    " GET_CONFIGURATION
*&---------------------------------------------------------------------*
*&      Form  SOW_DOCUMENT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_<CAB>  text
*----------------------------------------------------------------------*
FORM show_document  USING    p_cab TYPE any.

  DATA: i_attachx    TYPE  solix_tab.
  DATA: comp   TYPE i, lv_url TYPE char255.

  DATA: ls_xml    TYPE /sbxc/zsbx_st_img.

  CONSTANTS: c_app(12) TYPE c VALUE 'application',
             c_h(5)    TYPE c VALUE 'htlm'.


  IF disp_doc_active = 'X'.
    FIELD-SYMBOLS <url> TYPE any.
    ASSIGN COMPONENT 'URL' OF STRUCTURE p_cab TO <url>.

    FIELD-SYMBOLS <id_p> TYPE any.
*    ASSIGN COMPONENT 'UUID' OF STRUCTURE p_cab TO <id_p>.
    ASSIGN COMPONENT 'REF_DOC_NO_ORIG' OF STRUCTURE p_cab TO <id_p>.

    IF NOT <url> IS INITIAL.
      CALL METHOD go_attviewer->show_url
        EXPORTING
          url = <url>.
    ELSEIF NOT <id_p> IS INITIAL.

      FIELD-SYMBOLS <id_barc> TYPE any.
      ASSIGN COMPONENT 'BARCODE' OF STRUCTURE p_cab TO <id_barc>.

      PERFORM ler_imagem USING ls_xml comp <id_p> <id_barc> i_attachx.

* Load the HTML
      CALL METHOD go_attviewer->load_data(
        EXPORTING
          type                 = c_app "'application'
          subtype              = c_h "'htlm'
        IMPORTING
          assigned_url         = lv_url
        CHANGING
          data_table           = i_attachx
        EXCEPTIONS
          dp_invalid_parameter = 1
          dp_error_general     = 2
          cntl_error           = 3
          OTHERS               = 4 ) ##SUBRC_OK.

      CALL METHOD go_attviewer->show_url
        EXPORTING
          url = lv_url.
    ELSE.
      CALL METHOD go_attviewer->free.
      CREATE OBJECT go_attviewer
        EXPORTING
          parent = go_cell_att.
    ENDIF.

  ENDIF.
ENDFORM.                    " SOW_DOCUMENT
*&---------------------------------------------------------------------*
*&      Form  DOMAIN_XWARE_BNK
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM domain_xware_bnk .
  DATA: lt_dropdown TYPE lvc_t_dral,
        ls_dropdown TYPE lvc_s_dral,
        it_taba     TYPE STANDARD TABLE OF dd07v,
        it_tabb     TYPE STANDARD TABLE OF dd07v.
  CALL FUNCTION 'DD_DOMA_GET'
    EXPORTING
      domain_name   = 'XWARE_BNK'
      langu         = sy-langu
      withtext      = 'X'
    TABLES
      dd07v_tab_a   = it_taba
      dd07v_tab_n   = it_tabb
    EXCEPTIONS
      illegal_value = 1
      op_failure    = 2
      OTHERS        = 3.

  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
  LOOP AT  it_taba ASSIGNING FIELD-SYMBOL(<fs_taba>).
    ls_dropdown-handle = '1'.
    CONCATENATE <fs_taba>-domvalue_l '-'  <fs_taba>-ddtext INTO ls_dropdown-value SEPARATED BY space.
    ls_dropdown-int_value =  <fs_taba>-domvalue_l.
    APPEND ls_dropdown TO lt_dropdown.
  ENDLOOP.
  CALL METHOD go_grid_cab->set_drop_down_table
    EXPORTING
      it_drop_down_alias = lt_dropdown.
ENDFORM.                    " DOMAIN_XWARE_BNK
*&---------------------------------------------------------------------*
*&      Form  LER_IMAGEM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LS_XML  text
*      -->P_COMP  text
*      -->P_HEADER_UUID  text
*      -->P_I_ATTACHX  text
*----------------------------------------------------------------------*
FORM ler_imagem  USING xml TYPE /sbxc/st_img
                    comp TYPE i
*                    uuid TYPE /sbxc/zckp_invh-uuid
                    ref_doc TYPE /sbxc/zckp_invh-ref_doc_no_orig
                    barcode TYPE /sbxc/zckp_invh-barcode
                    i_attachx         TYPE  solix_tab.

  DATA: lr_saphety TYPE REF TO /sbxc/co_icnddocument_service.
*  DATA: lv_process TYPE /sbxc/zckp_processo.
  DATA: base64 TYPE /sbxc/z_data_response. "/sbxc/icnddocument_service_g26,
*        t_uuid TYPE /sbxc/icnddocument_service_g27.

  DATA: lt_system_fault TYPE REF TO cx_ai_system_fault,
        l_msg           TYPE string.

  DATA: lv_data TYPE REF TO data.

*  "Selecionar dados de acesso
*  SELECT SINGLE * INTO @DATA(ls_tab30)                  "#EC CI_NOORDER
*  FROM /sbxc/zckp_tab30.
*
*
*  CREATE DATA lv_data TYPE /sbxc/st_img.
*
*  CLEAR lr_saphety.
**  lv_process = 'WSIMG'.
*
*  TRY .
*      CREATE OBJECT lr_saphety
*        EXPORTING
*          logical_port_name = 'BASICHTTPBINDING_ICNDDOCUMENTSERVICE'.
*    CATCH cx_ai_system_fault INTO lt_system_fault.
*      l_msg = lt_system_fault->get_text( ).
*      MESSAGE l_msg TYPE lc_error.
*      RETURN.
*  ENDTRY.
*
**  IF sy-subrc <> 0.
**    RETURN.
**  ENDIF.
*  FIELD-SYMBOLS: <fs> TYPE any.
*  ASSIGN lv_data->* TO <fs>.
*  MOVE-CORRESPONDING xml TO <fs>.
** Criar XML em BASE64
*
*  t_uuid-user_id = ls_tab30-username.
* [REDACTED: sensitive line omitted from public export]
*  t_uuid-in_transport_document_id = uuid.
*  t_uuid-doc_type = '2'.
*
*  TRY .
*      lr_saphety->get_document_data_on_in_transp(
*          EXPORTING
*            input                  = t_uuid
*             IMPORTING
*                output = base64 ).
*    CATCH cx_ai_system_fault INTO lt_system_fault.
*      l_msg = lt_system_fault->get_text( ).
*      MESSAGE l_msg TYPE lc_error.
*      RETURN.
*  ENDTRY.

  """"""
  DATA: go_saphety  TYPE REF TO /sbxc/zcl_read_doc_saphety,
        gtp_doc_b64 TYPE string.

  TRY.
      CREATE OBJECT go_saphety.
      IF go_saphety IS NOT BOUND. "Objecto não se encontra instanciado.
      ELSE.
        base64 = go_saphety->read_doc( i_barcode = barcode i_ref_doc = ref_doc ).
      ENDIF.
    CATCH zcx_config_not_found.

  ENDTRY.


  DATA: img_base64 TYPE /sbxc/z_data_response-contentdatabytes.

  CLEAR: img_base64.
  img_base64 = base64-contentdatabytes.

  IF img_base64  NE space.

    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = img_base64 "imagem
      IMPORTING
        output_length = comp
      TABLES
        binary_tab    = i_attachx.
  ELSE.
    RETURN.
  ENDIF.
ENDFORM.                    " LER_IMAGEM
*&---------------------------------------------------------------------*
*& Form GET_DATA
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM get_data .

  "Selecionar dados da determinação dos processos
  REFRESH lt_tab12.
  SELECT * INTO TABLE lt_tab12
    FROM /sbxc/zckp_tab12
    WHERE processo IN p_proc.

  "Selecionar informação de emails
  REFRESH lt_hmail.
  SELECT * FROM /sbxc/zckp_hmail INTO TABLE lt_hmail
    WHERE processo IN p_proc AND
        ano IN p_ano AND
        seqno IN p_seqno.

  "Selecionar histórico de mensagens
  REFRESH lt_tab09.
  SELECT * INTO TABLE lt_tab09
    FROM /sbxc/zckp_tab09
     WHERE processo IN p_proc
       AND ano  IN p_ano
       AND seqno IN p_seqno.

ENDFORM.
