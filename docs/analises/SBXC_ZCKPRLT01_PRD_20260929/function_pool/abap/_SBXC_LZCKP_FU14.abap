FUNCTION /sbxc/zckp_dbl_clk_lin.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(CAB)
*"     REFERENCE(CAMPO) TYPE  DD03L-FIELDNAME
*"     REFERENCE(LINHA)
*"  TABLES
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header LIKE /sbxc/zckp_invh.
  DATA: item LIKE  /sbxc/zckp_invi.

  DATA: fieldcat_linha TYPE slis_fieldcat_alv,
        fieldcat_tab   TYPE slis_t_fieldcat_alv,
        grupos         TYPE slis_t_sp_group_alv,
        wa_grupos      TYPE slis_sp_group_alv,
        wa_eventos     TYPE slis_alv_event,
        eventos        TYPE slis_t_event,
        layout         TYPE slis_layout_alv,
        is_variant     TYPE disvariant,
        reprepid       TYPE slis_reprep_id,
        grid_set       TYPE lvc_s_glay.

  DATA: wa_fieldcat LIKE LINE OF fieldcat_tab.
  DATA: programa LIKE sy-repid.



* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  MOVE-CORRESPONDING linha TO item.

  CASE campo.
    WHEN 'CLAS_MULT'.
* Verifica se ja foram feitas alterações aos montantes e valores para imp. cont. mult                                   4
      DATA: ent_merc     LIKE ekpo-wepos, nao_avaliada LIKE ekpo-weunb, poref LIKE /sbxc/zckp_tab12-poref.
      SELECT SINGLE poref FROM /sbxc/zckp_tab12 INTO poref WHERE poref = 'X' AND processo = header-processo.
      CHECK sy-subrc = 0.
      CLEAR: ent_merc, nao_avaliada.
      SELECT SINGLE wepos FROM ekpo INTO ent_merc     WHERE ebeln = item-po_number AND ebelp = item-po_item.
      SELECT SINGLE weunb FROM ekpo INTO nao_avaliada WHERE ebeln = item-po_number AND ebelp = item-po_item.
      IF ent_merc = '' OR nao_avaliada = 'X'.
        "Continuar
      ELSE.
        RETURN.
      ENDIF.
      CLEAR: w_item_sub_ref, w_header_ref.
      CLEAR itab_ekkn. REFRESH itab_ekkn.
      MOVE-CORRESPONDING item TO w_item_sub_ref.
      MOVE-CORRESPONDING header TO w_header_ref.

      SELECT * FROM /sbxc/zckp_ekkn WHERE
                  processo = item-processo AND
                  ano = item-ano  AND
                  seqno = item-seqno AND
                  ebeln = item-po_number
                  AND ebelp = item-po_item.

        MOVE-CORRESPONDING /sbxc/zckp_ekkn TO itab_ekkn.
        APPEND itab_ekkn.
      ENDSELECT.

      IF sy-subrc NE 0.

        SELECT * FROM ekpo INTO CORRESPONDING FIELDS OF ekpo
               WHERE ebeln = item-po_number AND
                    ebelp = item-po_item.

          IF ekpo-webre = ' '. " EF/EM não esta activo

            REFRESH: lt_xekbes,  lt_xekbe.
            CLEAR: lt_xekbes,  lt_xekbe.

            CALL FUNCTION 'ME_READ_HISTORY'
              EXPORTING
                ebeln  = ekpo-ebeln
                ebelp  = ekpo-ebelp
                webre  = 'X'
              TABLES
                xekbe  = lt_xekbe
                xekbes = lt_xekbes.

            SELECT * FROM ekkn WHERE ebeln = item-po_number
                                    AND ebelp = item-po_item.

              MOVE-CORRESPONDING ekkn TO itab_ekkn.
              READ TABLE lt_xekbes INTO ls_xekbes WITH  KEY ebelp = ekpo-ebelp
                                                     zekkn = ekkn-zekkn.

              IF header-doctypesaphety = 'NC' OR header-doctypesaphety = 'CREDITNOTE'
                or header-doctypesaphety = '381'. "ODC - 12_03_2021.
                itab_ekkn-wrbtr = ls_xekbes-rewwr.
                itab_ekkn-menge = ls_xekbes-remng.
              ELSE.
                itab_ekkn-wrbtr = ( itab_ekkn-vproz *  item-item_amount  ) / 100.
                itab_ekkn-menge = ( itab_ekkn-vproz *  item-quantity  ) / 100 .
              ENDIF.

              itab_ekkn-processo = item-processo.
              itab_ekkn-seqno    = item-seqno.
              itab_ekkn-ano      = item-ano.

              APPEND itab_ekkn.
            ENDSELECT.
          ENDIF.
        ENDSELECT.
      ENDIF.
      CLEAR layout.
      layout-colwidth_optimize = 'X'.
      layout-zebra = ' '.
      layout-no_vline = ' '.
      layout-no_hline = ' '.
      layout-def_status = 'A'.
      layout-edit = ' '.
      layout-edit_mode = ' '.
      layout-box_fieldname = 'BOX'.
      layout-window_titlebar = sy-title.
      is_variant-report = sy-repid.

      programa = sy-repid.
      grid_set-edt_cll_cb = 'X'.

* Eventos a capturar
      REFRESH eventos.
      REFRESH fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'EBELN'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-col_pos = 1.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'EBELP'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-col_pos = 2.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'ZEKKN'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-col_pos = 3.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'VPROZ'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 4.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'SAKTO'.
      wa_fieldcat-tabname = 'ITAB_EKKN'.
      wa_fieldcat-ref_tabname = 'SKA1'.
      wa_fieldcat-ref_fieldname = 'SAKNR'.
      wa_fieldcat-no_zero = 'X'.
      wa_fieldcat-datatype = 'CHAR'.
      wa_fieldcat-inttype   = 'C'.
      wa_fieldcat-outputlen = '10'.
      wa_fieldcat-rollname = 'SAKNR'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 5.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'KOSTL'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 6.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'PS_PSP_PNR'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 7.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'AUFNR'.
      wa_fieldcat-ref_tabname = 'EKKN'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 8.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'MENGE'.
      wa_fieldcat-ref_tabname = 'COBL_MRM_D'.
      wa_fieldcat-edit = 'X'.
      wa_fieldcat-col_pos = 9.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'WRBTR'.
      wa_fieldcat-ref_tabname = 'COBL_MRM_D'.
      wa_fieldcat-col_pos = 10.
      wa_fieldcat-edit = 'X'.


      APPEND wa_fieldcat TO fieldcat_tab.
      SORT itab_ekkn_alv ASCENDING BY ebeln ebelp zekkn.

      wa_eventos-name = 'DATA_CHANGED'.
      wa_eventos-form = 'F_DATA_CHANGED'.
      APPEND wa_eventos TO eventos.

      CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
        EXPORTING
          it_fieldcat              = fieldcat_tab
          it_events                = eventos
          is_layout                = layout
          i_grid_settings          = grid_set
          i_callback_pf_status_set = 'STATUS_CLASSM'
          is_variant               = is_variant
          i_callback_program       = programa "'SAPLZCKP_COCKPIT'                                   1
          i_callback_user_command  = 'USER_COMMAND_CLASSM'
          i_save                   = 'X'
          it_special_groups        = grupos
          i_screen_start_column    = 10
          i_screen_start_line      = 5
          i_screen_end_column      = 100
          i_screen_end_line        = 20
        TABLES
          t_outtab                 = itab_ekkn "#EC CI_FLDEXT_OK[2610650]
        EXCEPTIONS
          program_error            = 1
          OTHERS                   = 2.
    WHEN 'PO_NUMBER'.
      IF item-po_number IS NOT INITIAL.
        SET PARAMETER ID 'BES' FIELD item-po_number.
        CALL TRANSACTION 'ME23N' AND SKIP FIRST SCREEN.
      ENDIF.
    WHEN 'REF_DOC'.
      IF item-ref_doc IS NOT INITIAL.
*        SET PARAMETER ID 'MBN' FIELD item-ref_doc.
*        SET PARAMETER ID 'MJA' FIELD item-ref_doc_year.
*        CALL TRANSACTION 'MB03' AND SKIP FIRST SCREEN.
        "MB03 obsoleta
        CONSTANTS: lco_action_display  TYPE goaction VALUE 'A04',
                   lco_refdoc_material TYPE refdoc   VALUE 'R02',
                   lco_okcode_go       TYPE okcode   VALUE 'OK_GO'.
        CALL FUNCTION 'MIGO_DIALOG'
          EXPORTING
            i_action            = lco_action_display
            i_refdoc            = lco_refdoc_material
            i_notree            = abap_true
            i_skip_first_screen = abap_true
            i_deadend           = abap_true
            i_okcode            = lco_okcode_go
            i_new_rollarea      = abap_true
            i_mblnr             = item-ref_doc
            i_mjahr             = item-ref_doc_year
          EXCEPTIONS
            illegal_combination = 1
            OTHERS              = 2.
        IF sy-subrc <> 0.
* Implement suitable error handling here
        ENDIF.

      ENDIF.
    WHEN 'GL_ACCOUNT'.
      IF item-gl_account IS NOT INITIAL.
        SET PARAMETER ID 'SAK' FIELD item-gl_account .
        SET PARAMETER ID 'BUK' FIELD header-comp_code .
        CALL TRANSACTION 'FS00' AND SKIP FIRST SCREEN.
      ENDIF.
    WHEN 'COSTCENTER'.
      IF item-costcenter IS NOT INITIAL.
        SET PARAMETER ID 'KOS' FIELD item-costcenter .
        CALL TRANSACTION 'KS03' AND SKIP FIRST SCREEN.
      ENDIF.
    WHEN OTHERS.
  ENDCASE.

ENDFUNCTION.

**&---------------------------------------------------------------------*
**&      Form  f_data_changed
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->RR_DATA_CHANGED  text
**----------------------------------------------------------------------*
*FORM f_data_changed USING rr_data_changed TYPE REF TO
*                                          cl_alv_changed_data_protocol.
*  DATA: ls_mod_cell TYPE lvc_s_modi OCCURS 0 WITH HEADER LINE,
*    ls_good_cell TYPE lvc_s_modi OCCURS 0 WITH HEADER LINE,
*    text1(30).
*
*  READ TABLE rr_data_changed->mt_good_cells INTO ls_good_cell INDEX 1.
*  CHECK sy-subrc EQ 0.
*  READ TABLE itab_ekkn INDEX ls_good_cell-row_id.
*  CHECK sy-subrc EQ 0.
*  CONCATENATE 'itab_ekkn-' ls_good_cell-fieldname  INTO text1.
*  ASSIGN (text1) TO <f1>.
*  IF <f1> IS ASSIGNED.
*    <f1> = ls_good_cell-value.
*  ENDIF.
*ENDFORM.                    "f_data_changed                                  1
**&---------------------------------------------------------------------*
**&      Form  f_data_changed_div
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->RR_DATA_CHANGED  text
**----------------------------------------------------------------------*
*FORM f_data_changed_div USING rr_data_changed TYPE REF TO
*                                          cl_alv_changed_data_protocol.
**  PERFORM USER_COMMAND_classm.
*ENDFORM.                    "f_data_changed*
