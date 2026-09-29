FUNCTION /sbxc/zckp_read_email.
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


* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.
  DATA: fieldcat_linha TYPE slis_fieldcat_alv,
             fieldcat_tab TYPE slis_t_fieldcat_alv,
             grupos TYPE slis_t_sp_group_alv,
             wa_grupos TYPE slis_sp_group_alv,
             wa_eventos TYPE slis_alv_event,
             eventos TYPE slis_t_event,
             layout TYPE slis_layout_alv,
             is_variant TYPE disvariant,
             reprepid TYPE slis_reprep_id,
             grid_set TYPE lvc_s_glay.

  DATA:    wa_fieldcat       LIKE LINE OF fieldcat_tab.

  DATA: programa LIKE sy-repid.

  DATA: BEGIN OF itab_doc_alv OCCURS 0,
         num_mess LIKE /sbxc/zckp_hmail-num_mess,
         sender LIKE /sbxc/zckp_hmail-sender,
         recip LIKE /sbxc/zckp_hmail-recip,
         data_envio LIKE /sbxc/zckp_hmail-data_envio,
         texto LIKE /sbxc/zckp_hmail-texto,
         smtp_addr LIKE /sbxc/zckp_hmail-smtp_addr,
 END OF itab_doc_alv.



  DATA: itab_doc LIKE itab_doc_alv OCCURS 0 WITH HEADER LINE.
  DATA: es_exit_caused_by_user TYPE  slis_exit_by_user,
          e_exit_caused_by_caller.
* Preencher estrutura cabecalho

*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
*  ENDLOOP.


  REFRESH msg_cockpit.
*  LOOP AT header.


  SELECT * FROM /sbxc/zckp_hmail INTO CORRESPONDING FIELDS OF TABLE itab_doc WHERE
       processo = header-processo AND
       ano = header-ano AND
       seqno = header-seqno.

* Eventos a capturar
  REFRESH eventos.
  REFRESH fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'NUM_MESS'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 1.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'SENDER'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 2.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'RECIP'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 3.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'DATA_ENVIO'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 4.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'TEXTO'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 5.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'SMTP_ADDR'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_HMAIL'.
  wa_fieldcat-col_pos = 6.
  APPEND wa_fieldcat TO fieldcat_tab.

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      it_fieldcat              = fieldcat_tab
      i_callback_user_command  = 'USER_COMMAND'
*     i_callback_top_of_page   = 'TOP_OF_PAGE'
*     i_background_id          = 'ALV_BACKGROUND'
      it_events                = eventos
      is_layout                = layout
*     i_grid_settings          = grid_set
*      i_callback_pf_status_set = 'SET_STATUS'
      is_variant               = is_variant
      i_callback_program       = programa
      i_save                   = 'X'
      it_special_groups        = grupos
      i_screen_start_column    = 10
      i_screen_start_line      = 5
      i_screen_end_column      = 150
      i_screen_end_line        = 30
    IMPORTING
      e_exit_caused_by_caller  = e_exit_caused_by_caller
      es_exit_caused_by_user   = es_exit_caused_by_user
    TABLES
      t_outtab                 = itab_doc
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.


ENDFUNCTION.
