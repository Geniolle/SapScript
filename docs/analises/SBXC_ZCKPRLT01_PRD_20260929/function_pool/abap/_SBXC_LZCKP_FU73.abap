FUNCTION /sbxc/zckp_item_subdividir.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM)
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------
 data: w_item_sub like line of t_item_sub,
        w_item_sub_tot like line of t_item_sub_tot,
        header like /sbxc/zckp_invh.
  data: fieldcat_linha type slis_fieldcat_alv,
        fieldcat_tab type slis_t_fieldcat_alv,
        grupos type slis_t_sp_group_alv,
        wa_grupos type slis_sp_group_alv,
        wa_eventos type slis_alv_event,
        eventos type slis_t_event,
        layout type slis_layout_alv,
        is_variant type disvariant,
        reprepid type slis_reprep_id,
        grid_set type lvc_s_glay.
  data: wa_fieldcat like line of fieldcat_tab.
  data: programa like sy-repid.
  move idx_lin to g_idx_lin_ref.
  move-corresponding cab to header.
  check ctrl-status1 <> '3' and ctrl-status1 <> '4'.
** Preencher estrutura item
  refresh t_item_sub_tot.
  clear: w_item_sub_tot, w_item_sub_ref, w_header_ref.
  move-corresponding header to  w_header_ref. "w_item_sub_ref.
  loop at linha.
    move-corresponding linha to w_item_sub_tot.
    append w_item_sub_tot to t_item_sub_tot.
  endloop.
  read table t_item_sub_tot into w_item_sub_ref index g_idx_lin_ref.
  if sy-subrc eq 0.
    refresh t_item_sub.
    clear w_item_sub.
    move-corresponding w_item_sub_ref to w_item_sub.
    append w_item_sub to t_item_sub.
  endif.
* Preencher estrutura cabecalho
  clear layout.
  layout-colwidth_optimize = 'X'.
  layout-zebra = ' '.
  layout-no_vline = ' '.
  layout-no_hline = ' '.
  layout-def_status = 'A'.
  layout-edit = ' '.
  layout-edit_mode = ' '.
  layout-window_titlebar = sy-title.
  layout-box_fieldname = 'BOX'.
  is_variant-report = sy-repid.
  programa = sy-repid.
  grid_set-edt_cll_cb = 'X'.
* Eventos a capturar
  refresh eventos.
  refresh fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'GL_ACCOUNT'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 1.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'COSTCENTER'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 2.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'WBS_ELEM'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 2.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'ORDERID'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 2.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'ZUONR'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 2.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'ITEM_AMOUNT'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 3.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'QUANTITY'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 4.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  clear wa_fieldcat.
  wa_fieldcat-fieldname = 'PO_UNIT'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 5.
  wa_fieldcat-edit = ''.
  append wa_fieldcat to fieldcat_tab.
  wa_fieldcat-fieldname = 'TAX_CODE_SAP'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_INVI'.
  wa_fieldcat-col_pos = 5.
  wa_fieldcat-edit = ''.
  wa_fieldcat-edit = 'X'.
  append wa_fieldcat to fieldcat_tab.
  wa_eventos-name = 'DATA_CHANGED'.
  wa_eventos-form = 'F_DATA_CHANGED_DIV'.
  append wa_eventos to eventos.
  call function 'REUSE_ALV_GRID_DISPLAY'
    exporting
      it_fieldcat              = fieldcat_tab
      it_events                = eventos
      is_layout                = layout
      i_grid_settings          = grid_set
      i_callback_pf_status_set = 'STATUS_SUB'
      is_variant               = is_variant
      i_callback_program       = programa "'SAPLZCKP_COCKPIT'
      i_callback_user_command  = 'USER_COMMAND_SUB'
      i_save                   = 'X'
      it_special_groups        = grupos
      i_screen_start_column    = 10
      i_screen_start_line      = 5
      i_screen_end_column      = 100
      i_screen_end_line        = 20
    tables
      t_outtab                 = t_item_sub "#EC CI_FLDEXT_OK[2215424]
    exceptions
      program_error            = 1
      others                   = 2.
  refresh linha.
  sort t_item_sub_tot by processo ano seqno invoice_doc_item.
  loop at t_item_sub_tot into w_item_sub_tot.
    move-corresponding w_item_sub_tot to linha.
    append linha.
  endloop.
*  perform USER_COMMAND_SUB.
ENDFUNCTION.
