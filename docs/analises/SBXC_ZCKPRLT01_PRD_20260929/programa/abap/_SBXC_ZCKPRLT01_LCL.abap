*----------------------------------------------------------------------*
*       CLASS lcl_event_receiver DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_event_receiver DEFINITION.

  PUBLIC SECTION.
    CLASS-METHODS
      handle_double_click
        FOR EVENT double_click OF cl_gui_alv_grid
        IMPORTING e_row
                  e_column
                  es_row_no
                  sender. " grid instance that has triggered the event

    CLASS-METHODS  handle_data_changed
      FOR EVENT data_changed OF cl_gui_alv_grid
      IMPORTING er_data_changed
                sender.

    CLASS-METHODS  handle_hotspot_click
      FOR EVENT hotspot_click OF cl_gui_alv_grid
      IMPORTING e_row_id
                e_column_id
                es_row_no.

    CLASS-METHODS  handle_context_menu
      FOR EVENT context_menu_request OF cl_gui_alv_grid
      IMPORTING e_object
                sender.

    CLASS-METHODS on_f4 FOR EVENT onf4 OF cl_gui_alv_grid
      IMPORTING sender
                e_fieldname
                e_fieldvalue
                es_row_no
                er_event_data
                et_bad_cells
                e_display.

    CLASS-DATA:
      ms_row        TYPE lvc_s_row  READ-ONLY, " stores selected row
      ms_col        TYPE lvc_s_col READ-ONLY,

      lt_std_fcodes TYPE ui_functions,
      lt_own_fcodes TYPE ui_functions.
ENDCLASS.                    "lcl_event_receiver DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_event_receiver IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_event_receiver IMPLEMENTATION.

  METHOD handle_double_click.

    DATA: l_selected TYPE lvc_t_row.
    DATA: linha_sel  TYPE sy-index.

    DATA: var1(40),
          var2(40),
          var3(40),
          nao_le(1).

    FIELD-SYMBOLS: <f1> TYPE any,
                   <f2> TYPE any,
                   <f3> TYPE any.

    lin_or_cab = 'LIN'.
    CASE sender.
      WHEN go_grid_proc.

        DATA: var10(40).
        FIELD-SYMBOLS: <var10> TYPE any.

        IF e_row-index IS NOT INITIAL.
          lcl_event_receiver=>ms_row = e_row. " save selected row for REFRESH
          linha_sel = e_row.
          lcl_event_receiver=>ms_col = e_column.

          READ TABLE gt_proc INDEX e_row ASSIGNING FIELD-SYMBOL(<fs_proc>).

          var10 = '<WA_CAB_TABLE>-PROCESSO'.

          REFRESH <t_cab_table>.
          ASSIGN r_wa_dyn_table_cab_tmp->* TO <wa_cab_table>.

          LOOP AT <t_cab_table_tmp> ASSIGNING <wa_cab_table>.
            ASSIGN (var10) TO <var10>.
            CHECK <var10> EQ <fs_proc>-processo.

            APPEND <wa_cab_table> TO <t_cab_table>.
          ENDLOOP.
          REFRESH <t_lin_table>.

          PERFORM get_tab_status.

          CALL METHOD go_grid_status->refresh_table_display( is_stable = gc_stable ).
          CALL METHOD go_grid_cab->refresh_table_display( is_stable = gc_stable ).
          CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).
        ELSE.
          RETURN.
        ENDIF.
      WHEN go_grid_lin.

        IF e_row-index IS NOT INITIAL.
          lcl_event_receiver=>ms_row = e_row. " save selected row for REFRESH
          lcl_event_receiver=>ms_col = e_column.
          PERFORM double_click_lin USING e_column e_row.
        ELSE.
          RETURN.
        ENDIF.
      WHEN go_grid_cab.

        IF e_row-index IS NOT INITIAL.
          lcl_event_receiver=>ms_row = e_row. " save selected row for REFRESH
          lcl_event_receiver=>ms_col = e_column.

          ASSIGN r_wa_dyn_table_cab->* TO <cab>.

          READ TABLE <t_cab_table> INDEX e_row ASSIGNING <cab>.
          PERFORM show_document USING <cab>.
          IF <cab> IS ASSIGNED.
            MOVE-CORRESPONDING <cab> TO key.
          ENDIF.
          var1 = '<cab>-VENDOR'.
          ASSIGN (var1) TO <f1>.
          SET PARAMETER ID 'LIF' FIELD  <f1>.
          ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.

          LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
            var1 = '<WA_LIN_TABLE>-PROCESSO'.
            var2 = '<WA_LIN_TABLE>-ANO'.
            var3 = '<WA_LIN_TABLE>-SEQNO'.

            ASSIGN (var1) TO <f1>.
            ASSIGN (var2) TO <f2>.
            ASSIGN (var3) TO <f3>.
            IF <f1> = key-processo AND
               <f2> = key-ano AND
               <f3> = key-seqno.
              nao_le = 'X'.
              EXIT.
            ENDIF.
          ENDLOOP.

          IF nao_le IS INITIAL.

            REFRESH <t_lin_table>.

            IF NOT <wa_lin_table> IS ASSIGNED.
              ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
            ENDIF.

            SELECT * FROM (str_lin)
            INTO CORRESPONDING FIELDS OF <wa_lin_table>
            WHERE processo = key-processo
                AND ano = key-ano
                AND seqno = key-seqno.
              APPEND <wa_lin_table> TO <t_lin_table>.
            ENDSELECT.

            FIELD-SYMBOLS: <fs_ponum>     TYPE any, <fs_poitm> TYPE any,
                           <fs_refdc>     TYPE any, <fs_refitm> TYPE any,
                           <fs_refyer>    TYPE any, <fs_amount> TYPE any,
                           <fs_invdcitem> TYPE any,
                           <fs_quant>     TYPE any.
            FIELD-SYMBOLS: <fs_ponum2>        TYPE any, <fs_poitm2> TYPE any,
                           <fs_refdc2>        TYPE any, <fs_refitm2> TYPE any,
                           <fs_refyer2>       TYPE any, <fs_amount2> TYPE any,
                           <fs_invdcitem2>    TYPE any,
                           <fs_quant2>        TYPE any,
                           <fs_item_text>     TYPE any, <fs_po_unit> TYPE any,
                           <fs_po_unit_2>     TYPE any, <fs_po_item_text2> TYPE any.
            DATA: lv_campo(50).
            IF NOT <wa_lin_table2> IS ASSIGNED.
              ASSIGN r_wa_dyn_table_lin2->* TO <wa_lin_table2>.
            ELSE.
              CLEAR <wa_lin_table2>.
            ENDIF.
            IF NOT <t_lin_table2> IS ASSIGNED.
              ASSIGN r_dyn_table_lin->* TO <t_lin_table2>.
            ELSE.
              REFRESH <t_lin_table2>.
            ENDIF.
            IF NOT <wa_lin_table> IS ASSIGNED.
              ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
            ENDIF.

            LOOP AT <t_lin_table> ASSIGNING <wa_lin_table2>.
              APPEND <wa_lin_table2> TO <t_lin_table2>.
            ENDLOOP.
*
            SELECT * FROM (str_lin)
            INTO CORRESPONDING FIELDS OF <wa_lin_table>
            WHERE processo = key-processo
                AND ano = key-ano
                AND seqno = key-seqno .

              lv_campo = '<WA_LIN_TABLE>-PO_NUMBER'.
              ASSIGN (lv_campo) TO <fs_ponum>.

              lv_campo = '<WA_LIN_TABLE>-PO_ITEM'.
              ASSIGN (lv_campo) TO <fs_poitm>.

              lv_campo = '<WA_LIN_TABLE>-REF_DOC'.
              ASSIGN (lv_campo) TO <fs_refdc>.

              lv_campo = '<WA_LIN_TABLE>-REF_DOC_ITEM'.
              ASSIGN (lv_campo) TO <fs_refitm>.

              lv_campo = '<WA_LIN_TABLE>-REF_DOC_YEAR'.
              ASSIGN (lv_campo) TO <fs_refyer>.

              lv_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
              ASSIGN (lv_campo) TO <fs_amount>.

              lv_campo = '<WA_LIN_TABLE>-QUANTITY'.
              ASSIGN (lv_campo) TO <fs_quant>.

              lv_campo = '<WA_LIN_TABLE>-ITEM_TEXT'.
              ASSIGN (lv_campo) TO <fs_item_text>.

              lv_campo = '<WA_LIN_TABLE>-PO_UNIT'.
              ASSIGN (lv_campo) TO <fs_po_unit>.

              lv_campo = '<WA_LIN_TABLE>-INVOICE_DOC_ITEM'.
              ASSIGN (lv_campo) TO <fs_invdcitem>.

              LOOP AT <t_lin_table2> ASSIGNING <wa_lin_table2>.
                lv_campo = '<WA_LIN_TABLE2>-PO_NUMBER'.
                ASSIGN (lv_campo) TO <fs_ponum2>.

                lv_campo = '<WA_LIN_TABLE2>-PO_ITEM'.
                ASSIGN (lv_campo) TO <fs_poitm2>.

                lv_campo = '<WA_LIN_TABLE2>-REF_DOC'.
                ASSIGN (lv_campo) TO <fs_refdc2>.

                lv_campo = '<WA_LIN_TABLE2>-REF_DOC_ITEM'.
                ASSIGN (lv_campo) TO <fs_refitm2>.

                lv_campo = '<WA_LIN_TABLE2>-REF_DOC_YEAR'.
                ASSIGN (lv_campo) TO <fs_refyer2>.

                lv_campo = '<WA_LIN_TABLE2>-INVOICE_DOC_ITEM'.
                ASSIGN (lv_campo) TO <fs_invdcitem2>.

                IF <fs_ponum2> = <fs_ponum> AND
                         <fs_poitm2> = <fs_poitm> AND
                         <fs_refdc2>  = <fs_refdc> AND
                         <fs_refitm2> = <fs_refitm> AND
                         <fs_refyer2> = <fs_refyer> AND
                         <fs_invdcitem2> = <fs_invdcitem> AND
                          <fs_ponum> NE space.

*Valor e qt
                  lv_campo = '<WA_LIN_TABLE2>-ITEM_AMOUNT'.
                  ASSIGN (lv_campo) TO <fs_amount2>.

                  lv_campo = '<WA_LIN_TABLE2>-QUANTITY'.
                  ASSIGN (lv_campo) TO <fs_quant2>.

                  lv_campo = '<WA_LIN_TABLE2>-PO_UNIT'.
                  ASSIGN (lv_campo) TO <fs_po_unit_2>.

                  lv_campo = '<WA_LIN_TABLE2>-ITEM_TEXT'.
                  ASSIGN (lv_campo) TO <fs_po_item_text2>.

                  <fs_amount> = <fs_amount2>.
                  <fs_quant> = <fs_quant2>.
                  <fs_po_unit> = <fs_po_unit_2>.
                  <fs_item_text> = <fs_po_item_text2>.

                  MOVE-CORRESPONDING <wa_lin_table> TO <wa_lin_table2>.
                  MODIFY <t_lin_table2> FROM <wa_lin_table2> INDEX sy-tabix.

                ENDIF.
              ENDLOOP.
            ENDSELECT.

            REFRESH <t_lin_table>.

            LOOP AT <t_lin_table2> ASSIGNING <wa_lin_table>.
              lv_campo = '<WA_LIN_TABLE>-COSTCENTER'.
              ASSIGN (lv_campo) TO <f1>.
              IF sy-subrc EQ 0 AND <f1> CO ' 0'.
                CLEAR <f1>.
                UNASSIGN <f1>.
              ENDIF.

              lv_campo = '<WA_LIN_TABLE>-GL_ACCOUNT'.
              ASSIGN (lv_campo) TO <f1>.
              IF sy-subrc EQ 0 AND <f1> CO ' 0'.
                CLEAR <f1>.
                UNASSIGN <f1>.
              ENDIF.

              lv_campo = '<WA_LIN_TABLE>-ORDERID'.
              ASSIGN (lv_campo) TO <f1>.
              IF sy-subrc EQ 0 AND <f1> CO ' 0'.
                CLEAR <f1>.
                UNASSIGN <f1>.
              ENDIF.
              APPEND <wa_lin_table> TO <t_lin_table>.
            ENDLOOP.

          ENDIF.

* Verificar se existe função de alteração da tabela de linhas

          READ TABLE it_fm_display INTO wa_fm_display WITH KEY processo = key-processo.

          IF wa_fm_display-fm_lin_disp NE space AND
            sy-subrc EQ 0.

            CALL FUNCTION 'FUNCTION_EXISTS'
              EXPORTING
                funcname           = wa_fm_display-fm_lin_disp
              EXCEPTIONS
                function_not_exist = 1
                OTHERS             = 2.

            IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
            ELSE.

              IF NOT <wa_lin_table> IS ASSIGNED     .
                ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
              ENDIF.

              PERFORM double_click USING e_column e_row.

              CALL FUNCTION wa_fm_display-fm_lin_disp
                IMPORTING
                  refresh = refresh_table
                TABLES
                  cor     = f_cor       " Cor para células
                  linha   = <t_lin_table>
                CHANGING
                  cab     = <cab>.

              gs_layout-stylefname = 'CELL_COLOR'.

            ENDIF.
          ENDIF.

          gs_layout-stylefname = 'CELLSTYLE'.
          CALL METHOD go_grid_lin->set_selected_rows
            EXPORTING
              it_index_rows = l_selected.

          CALL METHOD go_grid_lin->set_frontend_layout
            EXPORTING
              is_layout = gs_layout.
        ELSE.
          RETURN.
        ENDIF.
    ENDCASE.

    FIELD-SYMBOLS: <status_cab> TYPE any.
    DATA l_status_cab(40).

    l_status_cab = '<CAB>-STATUS1'.
    ASSIGN (l_status_cab) TO <status_cab>.

    LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.

      ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
                  <fs_style>.
      PERFORM estilo_campos USING 'LIN' <status_cab>.
      MODIFY <t_lin_table> FROM <wa_lin_table> INDEX sy-tabix.

    ENDLOOP.

    REFRESH lt_filter.
    ls_filter-fieldname = 'PROCESSO'.
    ls_filter-low       = key-processo.
    ls_filter-high      = space.
    ls_filter-option    = 'EQ'.
    ls_filter-sign      = 'I'.
    APPEND ls_filter TO lt_filter.

    ls_filter-fieldname = 'ANO'.
    ls_filter-low       = key-ano.
    ls_filter-high      = space.
    ls_filter-option    = 'EQ'.
    ls_filter-sign      = 'I'.
    APPEND ls_filter TO lt_filter.

    ls_filter-fieldname = 'SEQNO'.
    ls_filter-low       = key-seqno.
    ls_filter-high      = space.
    ls_filter-option    = 'EQ'.
    ls_filter-sign      = 'I'.
    APPEND ls_filter TO lt_filter.

    CALL METHOD go_grid_lin->set_filter_criteria
      EXPORTING
        it_filter = lt_filter.

    CALL METHOD go_grid_lin->refresh_table_display(
        is_stable =
                    gc_stable ).

    CALL METHOD go_grid_cab->refresh_table_display(
        is_stable =
                    gc_stable ).
  ENDMETHOD.                    "handle_double_click

  METHOD handle_data_changed.

    DATA: ls_mod_cell TYPE lvc_s_modi,
          l_dcpfm     TYPE usr01-dcpfm,
          pspnr       TYPE prps-pspnr,
          l_prctr     TYPE proj-prctr,
          l_aufnr     TYPE aufk-aufnr,
          l_ebeln     TYPE ekpo-ebeln.

    FIELD-SYMBOLS: <wa_temp_cab> TYPE any,
                   <ftmp>        TYPE any,
                   <ftmp2>       TYPE any,
                   <item_amount> TYPE any,
                   <taxcode>     TYPE any,
                   <quant_fact>  TYPE any,
                   <ebeln>       TYPE any,
                   <ebelp>       TYPE any,
                   <bpmng>       TYPE any,
                   <anln2>       TYPE any.
    FIELD-SYMBOLS: <cc> TYPE any.

    DATA: lt_tax_values TYPE /sbxc/zckp_t_values,
          l_tax_values  TYPE /sbxc/zckp_s_values.
    DATA: n_lines  TYPE i,
          lv_knumh TYPE a003-knumh,
          lv_kbetr TYPE konp-kbetr.

    gd_changes = 'X'. "<- houve modificações

    DATA: var(35).
    FIELD-SYMBOLS: <f3> TYPE any.

    IF er_data_changed->mt_good_cells[] IS NOT INITIAL.

      CASE sender.
        WHEN go_grid_cab.

          lin_or_cab = 'CAB'.

          LOOP AT er_data_changed->mt_mod_cells INTO ls_mod_cell WHERE error IS INITIAL.

            IF ls_mod_cell-fieldname EQ 'GROSS_AMOUNT' OR ls_mod_cell-fieldname EQ 'QUANTITY' OR
            ls_mod_cell-fieldname EQ  'DEL_COSTS'.
              SELECT SINGLE dcpfm FROM usr01
                INTO l_dcpfm
                WHERE bname = v_uname.
              CASE l_dcpfm.
                WHEN ' '. "1.234.567,89
                  TRANSLATE ls_mod_cell-value USING '. '.
                  TRANSLATE ls_mod_cell-value USING ',.'.
                  CONDENSE ls_mod_cell-value NO-GAPS.
                WHEN 'X'. "1,234,567.89
                  TRANSLATE ls_mod_cell-value USING ', '.
                  CONDENSE ls_mod_cell-value NO-GAPS.
                WHEN 'Y'. "1 234 567,89
                  TRANSLATE ls_mod_cell-value USING ',.'.
                  CONDENSE ls_mod_cell-value NO-GAPS.
              ENDCASE.
            ELSEIF ls_mod_cell-fieldname EQ 'MWSKZ'.
              DATA: l_campo1(40),
                    lv_tabix TYPE sy-tabix.
              ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
              ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
              LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
                lv_tabix = sy-tabix.
                MOVE '<WA_LIN_TABLE>-TAX_CODE_SAP' TO var.
                ASSIGN (var) TO <f3>.
                IF <f3> IS INITIAL.
                  TRANSLATE ls_mod_cell-value TO UPPER CASE.
                  MOVE ls_mod_cell-value TO <f3>.

                  CLEAR l_campo1.
                  l_campo1 = '<WA_TEMP_CAB>-COMP_CODE'.
                  ASSIGN (l_campo1) TO <ftmp>.
*               Det país da empresa
                  SELECT SINGLE land1 INTO t001-land1
                    FROM t001 WHERE bukrs = <ftmp>.
*               Det. taxa iva
                  CLEAR: lv_kbetr, lv_knumh.
                  "ODC - 02_12_2020
                  SELECT knumh INTO lv_knumh
                  FROM a003
                  WHERE kappl = 'TX'
                  AND aland = t001-land1
                  AND mwskz = ls_mod_cell-value
                    AND kschl NE 'MWVI'.
                    IF sy-subrc EQ 0.
                      SELECT SINGLE kbetr INTO @DATA(lv_kbetr_a)
                        FROM konp
                        WHERE knumh = @lv_knumh
                        AND kopos = 1.
                      lv_kbetr = lv_kbetr + lv_kbetr_a.
                    ELSE.
                      CLEAR lv_kbetr.
                    ENDIF.
                  ENDSELECT.
*                  SELECT SINGLE knumh INTO lv_knumh     "#EC CI_NOORDER
*                    FROM a003
*                    WHERE kappl = 'TX'
*                    AND aland = t001-land1
*                    AND mwskz = ls_mod_cell-value.
*
*                  SELECT SINGLE kbetr INTO @DATA(lv_kbetr)"konp-kbetr
*                    FROM konp
*                    WHERE knumh = @lv_knumh
*                    AND kopos = 1.
                  "Fim ODC - 02_12_2020
                  l_campo1 = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
                  ASSIGN (l_campo1) TO <ftmp>.
                  <ftmp> = lv_kbetr / 10.

                  l_campo1 = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
                  ASSIGN (l_campo1) TO <item_amount>.

                  l_campo1 = '<WA_LIN_TABLE>-TAX_AMOUNT'.
                  ASSIGN (l_campo1) TO <ftmp>.
                  <ftmp> = <item_amount> * ( lv_kbetr / 1000 ).
                  MODIFY <t_lin_table> FROM <wa_lin_table> INDEX lv_tabix.
                ENDIF.
              ENDLOOP.
            ELSEIF ls_mod_cell-fieldname EQ 'VENDOR'.
              DATA: lv_lifnr TYPE lfa1-lifnr,
                    lv_name1 TYPE lfa1-name1.
              DATA: l_campo_n(25).
              CONDENSE ls_mod_cell-value NO-GAPS.
              CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
                EXPORTING
                  input  = ls_mod_cell-value
                IMPORTING
                  output = lv_lifnr.

              SELECT SINGLE name1 INTO lv_name1
                FROM lfa1
                WHERE lifnr EQ lv_lifnr.

              READ TABLE <t_cab_table> ASSIGNING <wa_cab_table> INDEX ls_mod_cell-row_id.
              CLEAR l_campo_n.
              l_campo_n = '<WA_CAB_TABLE>-NAME1'.
              ASSIGN (l_campo_n) TO <ftmp>.
              <ftmp> = lv_name1.
              CLEAR: l_campo_n.
              l_campo_n = '<WA_CAB_TABLE>-VENDOR'.
              ASSIGN (l_campo_n) TO <ftmp2>.
              <ftmp2> = lv_lifnr.

              MODIFY <t_cab_table> FROM <wa_cab_table> INDEX ls_mod_cell-row_id.
            ENDIF.

          ENDLOOP.

        WHEN go_grid_lin .

          lin_or_cab = 'LIN'.
          LOOP AT er_data_changed->mt_mod_cells INTO ls_mod_cell.

            CHECK ls_mod_cell-error IS INITIAL.

            ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
            READ TABLE <t_lin_table> ASSIGNING <wa_lin_table> INDEX ls_mod_cell-row_id.
            CHECK sy-subrc = 0.
            CONCATENATE '<WA_LIN_TABLE>-' ls_mod_cell-fieldname INTO var.
            ASSIGN (var) TO <f3>.

* Formata campos de quantidade e valor
            IF ls_mod_cell-fieldname EQ 'ITEM_AMOUNT' OR ls_mod_cell-fieldname EQ 'QUANTITY'
            OR ls_mod_cell-fieldname EQ 'TAX_AMOUNT' OR ls_mod_cell-fieldname EQ 'BPMNG'.
              SELECT SINGLE dcpfm FROM usr01
                INTO l_dcpfm
                WHERE bname = v_uname.
              CASE l_dcpfm.
                WHEN ' '. "1.234.567,89
                  TRANSLATE ls_mod_cell-value USING '. '.
                  TRANSLATE ls_mod_cell-value USING ',.'.
                  CONDENSE ls_mod_cell-value NO-GAPS.
                WHEN 'X'. "1,234,567.89
                  TRANSLATE ls_mod_cell-value USING ', '.
                  CONDENSE ls_mod_cell-value NO-GAPS.
                WHEN 'Y'. "1 234 567,89
                  TRANSLATE ls_mod_cell-value USING ',.'.
                  CONDENSE ls_mod_cell-value NO-GAPS.
              ENDCASE.
            ENDIF.
*Se cod iva alterado
            IF ls_mod_cell-fieldname EQ 'TAX_CODE_SAP'.
              IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
                MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
              ELSE.
                ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
              ENDIF.
              TRANSLATE ls_mod_cell-value TO UPPER CASE.

              DATA: l_campo(40).
              l_campo = '<WA_TEMP_CAB>-COMP_CODE'.
              ASSIGN (l_campo) TO <ftmp>.
*           Det país da empresa
              IF <ftmp> IS ASSIGNED.
                SELECT SINGLE land1 INTO t001-land1
                  FROM t001 WHERE bukrs = <ftmp>.
*           Det. taxa iva

                REFRESH: lt_tax_values.
                CLEAR lv_knumh.
                SELECT knumh kschl INTO (lv_knumh, l_tax_values-kschl)
                  FROM a003
                  WHERE kappl = 'TX'
                  AND aland = t001-land1
                  AND mwskz = ls_mod_cell-value.
                  IF sy-subrc EQ 0.
                    SELECT SINGLE kbetr INTO @DATA(lv_kbetr1)
                      FROM konp
                      WHERE knumh = @lv_knumh
                      AND kopos = 1.
                  ELSE.
                    CLEAR lv_kbetr1.
                  ENDIF.
                  l_tax_values-mwskz = ls_mod_cell-value.
                  l_tax_values-knumh = lv_knumh.
                  l_tax_values-kbetr = lv_kbetr1.

                  "CCT 01.10.2020
                  IF l_tax_values-kschl <> 'MWVI'.
                    APPEND l_tax_values TO lt_tax_values.
                  ENDIF.
                  CLEAR l_tax_values.
                ENDSELECT.
                CLEAR n_lines.
                DESCRIBE TABLE lt_tax_values LINES n_lines.
                IF n_lines > 1.
                  LOOP AT lt_tax_values INTO l_tax_values.
                    l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
                    ASSIGN (l_campo) TO <ftmp>.
                    IF sy-tabix EQ 1.
                      CLEAR <ftmp>.
                    ENDIF.
                    <ftmp> = <ftmp> + l_tax_values-kbetr / 10.

                    l_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
                    ASSIGN (l_campo) TO <item_amount>.



                    l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
                    ASSIGN (l_campo) TO <ftmp2>.
                    IF sy-tabix EQ 1.
                      CLEAR <ftmp2>.
                    ENDIF.
                    <ftmp2> = <ftmp2> + ( <item_amount> * ( l_tax_values-kbetr / 1000 ) ).

                  ENDLOOP.
                ELSE.
                  l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
                  ASSIGN (l_campo) TO <ftmp>.
                  <ftmp> = lv_kbetr1 / 10.

                  l_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
                  ASSIGN (l_campo) TO <item_amount>.

                  l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
                  ASSIGN (l_campo) TO <ftmp>.
                  <ftmp> = <item_amount> * ( lv_kbetr1 / 1000 ).
                ENDIF.

              ENDIF.
            ENDIF.

*Se montante alterado, recalcula imposto
            IF ls_mod_cell-fieldname EQ 'ITEM_AMOUNT'.
              IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
                MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
              ELSE.
                ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
              ENDIF.

              l_campo = '<WA_LIN_TABLE>-TAX_CODE_SAP'.
              ASSIGN (l_campo) TO <taxcode>.

              l_campo = '<WA_TEMP_CAB>-COMP_CODE'.
              ASSIGN (l_campo) TO <ftmp>.
*           Det país da empresa
              SELECT SINGLE land1 INTO t001-land1
                FROM t001 WHERE bukrs = <ftmp>.
*           Det. taxa iva

              REFRESH: lt_tax_values.
              CLEAR lv_knumh.
              SELECT knumh kschl INTO (lv_knumh, l_tax_values-kschl)
                FROM a003
                WHERE kappl = 'TX'
                AND aland = t001-land1
                AND mwskz = <taxcode>.
                IF sy-subrc EQ 0.
                  SELECT SINGLE kbetr INTO @DATA(lv_kbetr2)
                    FROM konp
                    WHERE knumh = @lv_knumh
                    AND kopos = 1.
                ELSE.
                  CLEAR lv_kbetr2.
                ENDIF.
                l_tax_values-mwskz = <taxcode>.
                l_tax_values-knumh = lv_knumh.
                l_tax_values-kbetr = lv_kbetr2.
                "CCT 01.10.2020
                IF l_tax_values-kschl <> 'MWVI'.
                  APPEND l_tax_values TO lt_tax_values.
                ENDIF.
                CLEAR l_tax_values.
              ENDSELECT.
              CLEAR n_lines.
              DESCRIBE TABLE lt_tax_values LINES n_lines.
              IF n_lines > 1.
                LOOP AT lt_tax_values INTO l_tax_values.
                  l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
                  ASSIGN (l_campo) TO <ftmp>.
                  IF sy-tabix EQ 1.
                    CLEAR <ftmp>.
                  ENDIF.
                  <ftmp> = <ftmp> + l_tax_values-kbetr / 10.


                  l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
                  ASSIGN (l_campo) TO <ftmp2>.
                  IF sy-tabix EQ 1.
                    CLEAR <ftmp2>.
                  ENDIF.
                  <ftmp2> = <ftmp2> + ( ls_mod_cell-value * ( l_tax_values-kbetr / 1000 ) ).

                ENDLOOP.
              ELSE.

                l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
                ASSIGN (l_campo) TO <ftmp>.
                <ftmp> = lv_kbetr2 / 10.

                l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
                ASSIGN (l_campo) TO <item_amount>.
                <item_amount> = ls_mod_cell-value * ( <ftmp> / 100 ).
              ENDIF.

            ENDIF.
            IF <f3> IS ASSIGNED.
              <f3> = ls_mod_cell-value.
            ENDIF.

* Se UMP alterada então altera campo quantidade
            IF ls_mod_cell-fieldname EQ 'BPMNG'.
              IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
                MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
              ELSE.
                ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
              ENDIF.

              l_campo = '<WA_LIN_TABLE>-PO_NUMBER'.
              ASSIGN (l_campo) TO <ebeln>.
              l_campo = '<WA_LIN_TABLE>-PO_ITEM'.
              ASSIGN (l_campo) TO <ebelp>.

              l_campo = '<WA_LIN_TABLE>-BPMNG'.
              ASSIGN (l_campo) TO <bpmng>.

              SELECT SINGLE bpumn, bpumz FROM ekpo INTO (@DATA(lv_bpumn), @DATA(lv_bpumz))
                  WHERE
                      ebeln = @<ebeln> AND
                      ebelp = @<ebelp>.

              l_campo = '<WA_LIN_TABLE>-QUANTITY'.
              ASSIGN (l_campo) TO <quant_fact>.
              <quant_fact> = ( <bpmng> *  lv_bpumn ) / lv_bpumz.

            ENDIF.

            IF ls_mod_cell-fieldname EQ 'QUANTITY'.
              IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
                MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
              ELSE.
                ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
              ENDIF.

              l_campo = '<WA_LIN_TABLE>-PO_NUMBER'.
              ASSIGN (l_campo) TO <ebeln>.
              l_campo = '<WA_LIN_TABLE>-PO_ITEM'.
              ASSIGN (l_campo) TO <ebelp>.

              l_campo = '<WA_LIN_TABLE>-QUANTITY'.
              ASSIGN (l_campo) TO <quant_fact>.

              SELECT SINGLE bpumn, bpumz FROM ekpo INTO (@DATA(lv_bpumn1), @DATA(lv_bpumz1))
                  WHERE
                      ebeln = @<ebeln> AND
                      ebelp = @<ebelp>.

              l_campo = '<WA_LIN_TABLE>-BPMNG'.
              ASSIGN (l_campo) TO <bpmng>.
              <bpmng> =  ( <quant_fact> *  lv_bpumz1 ) / lv_bpumn1.

            ENDIF.
* Deterimina o CC pelo pep
            IF ls_mod_cell-fieldname EQ 'WBS_ELEM'.

              CALL FUNCTION 'CONVERSION_EXIT_ABPSP_INPUT' ##FM_SUBRC_OK
                EXPORTING
                  input     = ls_mod_cell-value
                IMPORTING
                  output    = pspnr
                EXCEPTIONS
                  not_found = 1
                  OTHERS    = 2.

              SELECT SINGLE prctr FROM prps
                 INTO l_prctr
                 WHERE pspnr = pspnr.
              l_campo = '<WA_LIN_TABLE>-COSTCENTER'.
              ASSIGN (l_campo) TO <cc>.
              <cc> = l_prctr .
            ENDIF.
* Determina CC pela ordem
            IF ls_mod_cell-fieldname EQ 'ORDERID'.
              IF ls_mod_cell-value CO ' 0123456789'.
                UNPACK ls_mod_cell-value TO l_aufnr.
              ELSE.
                MOVE ls_mod_cell-value TO l_aufnr.
                CONDENSE l_aufnr NO-GAPS.
              ENDIF.
              SELECT SINGLE prctr FROM aufk
                 INTO l_prctr
                 WHERE aufnr =  l_aufnr.
              l_campo = '<WA_LIN_TABLE>-COSTCENTER'.
              ASSIGN (l_campo) TO <cc>.
              <cc> = l_prctr .
            ENDIF.
* Determina pedido por campo material do fornecedor
            ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
            l_campo = '<WA_TEMP_CAB>-COMP_CODE'.
            ASSIGN (l_campo) TO <ftmp>.

            IF ls_mod_cell-fieldname EQ 'IDNLF'.
              l_campo = '<WA_LIN_TABLE>-PO_NUMBER'.
              ASSIGN (l_campo) TO <ebeln>.
              IF <ebeln> IS INITIAL. "pedido nao preenchido
                SELECT SINGLE ebeln FROM ekpo           "#EC CI_NOORDER
                   INTO  l_ebeln
                   WHERE bukrs = <ftmp> AND
                         idnlf =  ls_mod_cell-value.
                <ebeln> = l_ebeln .
              ENDIF.
            ENDIF.
***
            " Se alterar o imobilizado considerar o subnº do imobilizado igual a 0
            IF ls_mod_cell-fieldname EQ 'ANLN1'.
              l_campo = '<WA_LIN_TABLE>-ANLN2'.
              ASSIGN (l_campo) TO <anln2>.
            ENDIF.
            MODIFY <t_lin_table> FROM <wa_lin_table> INDEX ls_mod_cell-row_id.
            IF <ftmp> IS ASSIGNED AND <ftmp2> IS ASSIGNED.
              CLEAR: <ftmp>, <ftmp2>.
            ENDIF.
          ENDLOOP.


*Valor controlo
          FIELD-SYMBOLS: <gross_amount> TYPE any, <tax_amount> TYPE any,
                         <del_costs>    TYPE any,
                         <processo>     TYPE any, <seqno> TYPE any,
                         <ano>          TYPE any,
                         <valor_cont>   TYPE any.

          IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
            MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
          ELSE.
            ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
          ENDIF.

          l_campo = '<WA_TEMP_CAB>-VALOR_CONTROLO'.
          ASSIGN (l_campo) TO <valor_cont>.

          l_campo = '<WA_TEMP_CAB>-GROSS_AMOUNT'.
          ASSIGN (l_campo) TO <gross_amount>.

          l_campo = '<WA_TEMP_CAB>-DEL_COSTS'.
          ASSIGN (l_campo) TO <del_costs>.

          CLEAR  <valor_cont>.
          LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
            l_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
            ASSIGN (l_campo) TO <item_amount>.
            l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
            ASSIGN (l_campo) TO <tax_amount>.
            <valor_cont> = <valor_cont> + <del_costs> + <item_amount> + <tax_amount>.
          ENDLOOP.

          DATA: yprocesso   TYPE /sbxc/zckp_invh-processo,
                yseqno      TYPE /sbxc/zckp_invh-seqno,
                yano        TYPE /sbxc/zckp_invh-ano,
                yvalor_cont TYPE /sbxc/zckp_invh-valor_controlo,
                yindex      TYPE sy-tabix.

          yvalor_cont = <valor_cont>.
          l_campo = '<WA_LIN_TABLE>-PROCESSO'.
          ASSIGN (l_campo) TO <processo>.
          l_campo = '<WA_LIN_TABLE>-SEQNO'.
          ASSIGN (l_campo) TO <seqno>.
          l_campo = '<WA_LIN_TABLE>-ANO'.
          ASSIGN (l_campo) TO <ano>.
          yprocesso = <processo>.
          yseqno = <seqno>.
          yano = <ano>.
          UNASSIGN: <processo>, <seqno>, <ano>.


          LOOP AT <t_cab_table> INTO <wa_temp_cab>.
            yindex = sy-tabix.
            l_campo = '<WA_TEMP_CAB>-PROCESSO'.
            ASSIGN (l_campo) TO <processo>.
            l_campo = '<WA_TEMP_CAB>-SEQNO'.
            ASSIGN (l_campo) TO <seqno>.
            l_campo = '<WA_TEMP_CAB>-ANO'.
            ASSIGN (l_campo) TO <ano>.

            IF <processo> EQ yprocesso AND <seqno> EQ yseqno AND <ano> EQ yano.
              l_campo = '<WA_TEMP_CAB>-VALOR_CONTROLO'.
              ASSIGN (l_campo) TO <valor_cont>.
              <valor_cont> = <gross_amount> * -1.
              <valor_cont> = <valor_cont> +    yvalor_cont.
              MODIFY <t_cab_table> FROM <wa_temp_cab> INDEX yindex.
              EXIT.
            ENDIF.

          ENDLOOP.

          IF er_data_changed->mt_deleted_rows[] IS INITIAL
            AND er_data_changed->mt_inserted_rows[] IS INITIAL.
            CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).
          ENDIF.
          IF <cab> IS ASSIGNED AND <wa_temp_cab> IS ASSIGNED.
            MOVE-CORRESPONDING <cab> TO <wa_temp_cab>.
          ELSE.
            ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
          ENDIF.

        WHEN OTHERS.
      ENDCASE.
* valor controlo
*      CALL METHOD go_grid_cab->refresh_table_display(
*          is_stable      = gc_stable
*          i_soft_refresh = 'X' ).
*      CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).

    ELSE.
      RETURN.
    ENDIF.
  ENDMETHOD.                    "handle_data_changed

  METHOD handle_context_menu.

    DATA: li_sel_col TYPE i.           "selected column
    DATA: li_sel_row TYPE i.
    FIELD-SYMBOLS: <cab>   TYPE any.
    DATA: campo(35).

    DATA: lt_fcodes TYPE ui_funcattr,
          ls_func   TYPE ui_func,
          lt_func   TYPE ui_functions. " Inactivate

    CLEAR: field, origem.

    r_click = 'X'.

    CASE sender.
      WHEN go_grid_cab.

        lin_or_cab = 'CAB'.

        IF lt_own_fcodes IS INITIAL. " nosso código
          PERFORM define_fcode_tables TABLES lt_std_fcodes lt_own_fcodes.
        ENDIF.

*  all standard functions
        CALL METHOD e_object->get_functions
          IMPORTING
            fcodes = lt_fcodes.

        LOOP AT lt_fcodes ASSIGNING FIELD-SYMBOL(<fs_fcode>).
          ls_func = <fs_fcode>-fcode.
          APPEND ls_func TO lt_func.
        ENDLOOP.

        e_object->hide_functions( lt_func ).
        e_object->add_separator( ).

        PERFORM adiciona_opcoes_h CHANGING e_object.

        CALL METHOD go_grid_cab->get_current_cell
          IMPORTING
            e_col = li_sel_col
            e_row = li_sel_row.

        CALL METHOD cl_gui_cfw=>flush
          EXCEPTIONS
            cntl_system_error = 1
            cntl_error        = 2
            OTHERS            = 3.
        IF sy-subrc NE 0.
* add your handling, for example
          CALL FUNCTION 'POPUP_TO_INFORM'
            EXPORTING
              titel = TEXT-012 "'erro'
              txt2  = sy-subrc
              txt1  = TEXT-011. "'Error in Flush'(500).
        ENDIF.

        READ TABLE t_fieldcat_cab TRANSPORTING ALL FIELDS INTO field INDEX li_sel_col.
        ASSIGN r_wa_dyn_table_cab->* TO <cab>.

        READ TABLE  <t_cab_table> ASSIGNING <cab>
                     INDEX     li_sel_row.

        CONCATENATE '<cab>' '-' field-fieldname INTO campo.
        ASSIGN (campo) TO <valor>.

        IF <valor> IS ASSIGNED.
          conteudo = <valor>.
        ENDIF.
        origem = 'CAB'.


      WHEN go_grid_lin.

        lin_or_cab = 'LIN'.

        IF lt_own_fcodes IS INITIAL. " nosso código
          PERFORM define_fcode_tables TABLES lt_std_fcodes lt_own_fcodes.
        ENDIF.

*  all standard functions
        CALL METHOD e_object->get_functions
          IMPORTING
            fcodes = lt_fcodes.

        LOOP AT lt_fcodes ASSIGNING FIELD-SYMBOL(<fs_fcode1>).
          ls_func = <fs_fcode1>-fcode.
          APPEND ls_func TO lt_func.
        ENDLOOP.

        e_object->hide_functions( lt_func ).
        e_object->add_separator( ).

        PERFORM adiciona_opcoes_l CHANGING e_object.

        CALL METHOD go_grid_lin->get_current_cell
          IMPORTING
            e_col = li_sel_col
            e_row = li_sel_row.

        CALL METHOD cl_gui_cfw=>flush
          EXCEPTIONS
            cntl_system_error = 1
            cntl_error        = 2
            OTHERS            = 3.

        IF sy-subrc NE 0.
          CALL FUNCTION 'POPUP_TO_INFORM'
            EXPORTING
              titel = TEXT-012 "'erro'
              txt2  = sy-subrc
              txt1  = TEXT-011. "'Error in Flush'(500).
        ENDIF.

        READ TABLE t_fieldcat_lin TRANSPORTING ALL FIELDS INTO field INDEX li_sel_col.
        ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
        READ TABLE  <t_lin_table> ASSIGNING <wa_lin_table>
                     INDEX     li_sel_row.

        CONCATENATE '<wa_lin_table>' '-' field-fieldname INTO campo.
        ASSIGN (campo) TO <valor>.
        IF <valor> IS ASSIGNED.
          conteudo = <valor>.
        ENDIF.
        origem = 'LIN'.

      WHEN OTHERS.
    ENDCASE.


    CALL METHOD e_object->hide_functions
      EXPORTING
        fcodes = lt_std_fcodes.

    CALL METHOD e_object->show_functions
      EXPORTING
        fcodes = lt_own_fcodes.


    CALL METHOD cl_gui_cfw=>flush
      EXCEPTIONS
        cntl_system_error = 1
        cntl_error        = 2
        OTHERS            = 3.
    IF sy-subrc NE 0.
      CALL FUNCTION 'POPUP_TO_INFORM'
        EXPORTING
          titel = TEXT-012
          txt2  = sy-subrc
          txt1  = TEXT-011.
    ENDIF.


  ENDMETHOD.                    "handle_context_menu
  METHOD handle_hotspot_click.

    DATA: var4(40), n_linhas(1) TYPE n.
    FIELD-SYMBOLS: <f4> TYPE any.

    TYPES: ty_r_st TYPE RANGE OF /sbxc/zckp_tab10-status.

    DATA: r_st  TYPE RANGE OF /sbxc/zckp_tab10-status,
          wa_st TYPE LINE OF ty_r_st.

    READ TABLE gt_status INTO wa_status INDEX e_row_id-index.

    IF sy-subrc EQ 0.
      lin_or_cab = 'STA'.
    ENDIF.

    IF e_column_id = 'BOX'.
      IF wa_status-box IS INITIAL.
        wa_status-box = 'X'.
      ELSE.
        CLEAR wa_status-box.

      ENDIF.
      MODIFY gt_status FROM wa_status INDEX e_row_id-index.
    ENDIF.
    REFRESH r_st.
    REFRESH <t_cab_table>.
    LOOP AT gt_status ASSIGNING FIELD-SYMBOL(<fs_status2>)
      WHERE box = 'X'.

      wa_st-sign = 'I'.
      wa_st-option = 'EQ'.
      wa_st-low = <fs_status2>-status.
      APPEND wa_st TO r_st.

      ADD 1 TO n_linhas.
    ENDLOOP.

    CLEAR var4.
    var4 = '<WA_CAB_TABLE>-STATUS1'.

    IF n_linhas EQ 1.
      REFRESH <t_cab_table>.
    ENDIF.

    ASSIGN r_wa_dyn_table_cab_tmp->* TO <wa_cab_table>.
    LOOP AT <t_cab_table_tmp> ASSIGNING <wa_cab_table>.
      ASSIGN (var4) TO <f4>.
      CHECK <f4> IN r_st.

      APPEND <wa_cab_table> TO <t_cab_table>.
    ENDLOOP.

    REFRESH <t_lin_table>.

    CALL METHOD go_grid_status->refresh_table_display( is_stable = gc_stable ).
    CALL METHOD go_grid_cab->refresh_table_display( is_stable = gc_stable ).
    CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).

  ENDMETHOD.                    "handle_hotspot_click


*Class implementation to handle the ONF4 event

  METHOD on_f4.
    DATA: mwskz TYPE bseg-mwskz.
    FIELD-SYMBOLS: <valor_cont> TYPE any,  <del_costs> TYPE any,
                   <processo>   TYPE any, <seqno> TYPE any,
                   <ano>        TYPE any.

    DATA: lt_tax_values TYPE /sbxc/zckp_t_values,
          l_tax_values  TYPE /sbxc/zckp_s_values.
    DATA: n_lines  TYPE i,
          lv_knumh TYPE a003-knumh.

    ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.

    READ TABLE  <t_lin_table> ASSIGNING <wa_lin_table>
                 INDEX     es_row_no-row_id.


    FIELD-SYMBOLS: <fsproc>  TYPE any, <fsano> TYPE any, <fsseqno> TYPE any,
                   <fsmwskz> TYPE any, <fsbuk> TYPE any.
    DATA: l_campo(30).

    l_campo = '<WA_LIN_TABLE>-PROCESSO'.
    ASSIGN (l_campo) TO <fsproc>.

    l_campo = '<WA_LIN_TABLE>-ANO'.
    ASSIGN (l_campo) TO <fsano>.

    l_campo = '<WA_LIN_TABLE>-SEQNO'.
    ASSIGN (l_campo) TO <fsseqno>.

    ASSIGN r_wa_dyn_table_cab->* TO <linha>.

    SELECT SINGLE * FROM (str_cab)
    INTO CORRESPONDING FIELDS OF <linha>
    WHERE processo = <fsproc>
      AND ano = <fsano>
      AND seqno = <fsseqno>.

    l_campo = '<LINHA>-COMP_CODE'.
    ASSIGN (l_campo) TO <fsbuk>.

    SELECT SINGLE land1 INTO t001-land1
      FROM t001 WHERE bukrs = <fsbuk>.

    SELECT SINGLE kalsm INTO @DATA(lv_kalsm)
      FROM t005 WHERE land1 = @t001-land1.

    CALL FUNCTION 'FI_F4_MWSKZ'
      EXPORTING
        i_kalsm = lv_kalsm
        i_stbuk = <fsbuk>
      IMPORTING
        e_mwskz = mwskz.

    l_campo =  '<WA_LIN_TABLE>-TAX_CODE_SAP'.
    ASSIGN (l_campo) TO <fsmwskz>.

    er_event_data->m_event_handled = 'X'.

    DATA: wa_tab05 TYPE /sbxc/zckp_tab05.
    SELECT SINGLE * FROM /sbxc/zckp_tab05
      INTO wa_tab05
      WHERE processo = <fsproc>
      AND uname = v_uname
      AND estrutura = 'LIN'
      AND campo = 'TAX_CODE_SAP'.

    IF sy-subrc EQ 0.
      <fsmwskz> = mwskz.

      FIELD-SYMBOLS: <wa_temp_cab> TYPE any,
                     <ftmp>        TYPE any,
                     <ftmp2>       TYPE any,
                     <item_amount> TYPE any.

      ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.

      l_campo = '<WA_TEMP_CAB>-COMP_CODE'.
      ASSIGN (l_campo) TO <ftmp>.

*   Det país da empresa
      SELECT SINGLE land1 INTO t001-land1
        FROM t001 WHERE bukrs = <ftmp>.

*   Det. taxa iva

      REFRESH: lt_tax_values.
      CLEAR lv_knumh.
      SELECT knumh kschl INTO (lv_knumh, l_tax_values-kschl)
        FROM a003
        WHERE kappl = 'TX'
        AND aland = t001-land1
        AND mwskz = <fsmwskz>.
        IF sy-subrc EQ 0.
          SELECT SINGLE kbetr INTO @DATA(lv_kbetr3)
            FROM konp
            WHERE knumh = @lv_knumh
            AND kopos = 1.
        ELSE.
          CLEAR lv_kbetr3.
        ENDIF.
        l_tax_values-mwskz = <fsmwskz>.
        l_tax_values-knumh = lv_knumh.
        l_tax_values-kbetr = lv_kbetr3.
        "CCT 01.10.2020
        IF l_tax_values-kschl <> 'MWVI'.
          APPEND l_tax_values TO lt_tax_values.
        ENDIF.
        CLEAR l_tax_values.
      ENDSELECT.
      CLEAR n_lines.
      DESCRIBE TABLE lt_tax_values LINES n_lines.
      IF n_lines > 1.
        LOOP AT lt_tax_values INTO l_tax_values.
          l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
          ASSIGN (l_campo) TO <ftmp>.
          <ftmp> = <ftmp> + l_tax_values-kbetr / 10.

          l_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
          ASSIGN (l_campo) TO <item_amount>.



          l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
          ASSIGN (l_campo) TO <ftmp2>.
          <ftmp2> = <ftmp2> + ( <item_amount> * ( l_tax_values-kbetr / 1000 ) ).
        ENDLOOP.
      ELSE.
        l_campo = '<WA_LIN_TABLE>-TAX_IMPOSTO_SAP'.
        ASSIGN (l_campo) TO <ftmp>.
        <ftmp> = lv_kbetr3 / 10.

        l_campo = '<WA_LIN_TABLE>-ITEM_AMOUNT'.
        ASSIGN (l_campo) TO <item_amount>.

        l_campo = '<WA_LIN_TABLE>-TAX_AMOUNT'.
        ASSIGN (l_campo) TO <ftmp>.
        <ftmp> = <item_amount> * ( lv_kbetr3 / 1000 ).
      ENDIF.

      l_campo = '<wa_temp_cab>-VALOR_CONTROLO'.
      ASSIGN (l_campo) TO <valor_cont>.

      l_campo = '<WA_TEMP_CAB>-DEL_COSTS'.
      ASSIGN (l_campo) TO <del_costs>.

      <valor_cont> = <valor_cont> + <del_costs> + <item_amount> + <ftmp>.

      MODIFY <t_lin_table> FROM <wa_lin_table> INDEX es_row_no-row_id.
      IF <ftmp> IS ASSIGNED AND <ftmp2> IS ASSIGNED.
        CLEAR: <ftmp>, <ftmp2>.
      ENDIF.

      CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).

      ASSIGN r_wa_dyn_table_cab->* TO <wa_temp_cab>.
      DATA: yprocesso   TYPE /sbxc/zckp_invh-processo,
            yseqno      TYPE /sbxc/zckp_invh-seqno,
            yano        TYPE /sbxc/zckp_invh-ano,
            yvalor_cont TYPE /sbxc/zckp_invh-valor_controlo,
            yindex      TYPE sy-tabix.

      yvalor_cont = <valor_cont>.
      l_campo = '<WA_LIN_TABLE>-PROCESSO'.
      ASSIGN (l_campo) TO <processo>.
      l_campo = '<WA_LIN_TABLE>-SEQNO'.
      ASSIGN (l_campo) TO <seqno>.
      l_campo = '<WA_LIN_TABLE>-ANO'.
      ASSIGN (l_campo) TO <ano>.
      yprocesso = <processo>.
      yseqno = <seqno>.
      yano = <ano>.
      UNASSIGN: <processo>, <seqno>, <ano>.
      LOOP AT <t_cab_table> INTO <wa_temp_cab>.
        yindex = sy-tabix.
        l_campo = '<WA_TEMP_CAB>-PROCESSO'.
        ASSIGN (l_campo) TO <processo>.
        l_campo = '<WA_TEMP_CAB>-SEQNO'.
        ASSIGN (l_campo) TO <seqno>.
        l_campo = '<WA_TEMP_CAB>-ANO'.
        ASSIGN (l_campo) TO <ano>.

        IF <processo> EQ yprocesso AND <seqno> EQ yseqno AND <ano> EQ yano.
          l_campo = '<WA_TEMP_CAB>-VALOR_CONTROLO'.
          ASSIGN (l_campo) TO <valor_cont>.
          <valor_cont> =    yvalor_cont.
          MODIFY <t_cab_table> FROM <wa_temp_cab> INDEX yindex.
          EXIT.
        ENDIF.

      ENDLOOP.
* valor controlo
      CALL METHOD go_grid_cab->refresh_table_display( is_stable = gc_stable ).
    ELSE.
      RETURN.
    ENDIF.
  ENDMETHOD.                                                "on_f4


ENDCLASS.                    "lcl_event_receiver IMPLEMENTATION


*----------------------------------------------------------------------*
*       CLASS lcl_event_toolbar DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_event_toolbar DEFINITION.

  PUBLIC SECTION.

    CLASS-METHODS:
      handle_toolbar
        FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive,

      handle_menu_button
        FOR EVENT menu_button OF cl_gui_alv_grid
        IMPORTING e_object e_ucomm,

      handle_user_command
        FOR EVENT user_command OF cl_gui_alv_grid
        IMPORTING e_ucomm,

      handle_toolbar_lin
        FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive,

      handle_toolbar_status
        FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive.

  PRIVATE SECTION.

ENDCLASS.                    "lcl_event_toolbar DEFINITION


*----------------------------------------------------------------------*
*       CLASS lcl_event_toolbar IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_event_toolbar IMPLEMENTATION.

  METHOD handle_toolbar_status.

    lin_or_cab = 'STA'.

    SORT it_botoes_s BY function.
    DELETE ADJACENT DUPLICATES FROM it_botoes_s COMPARING function.
    SORT it_botoes_s BY num.

    LOOP AT it_botoes_s ASSIGNING FIELD-SYMBOL(<fs_botoes>).

      CHECK <fs_botoes>-num_pai EQ space.

      CLEAR gs_grid4_toolbar.
      MOVE <fs_botoes>-function TO gs_grid4_toolbar-function.
      SELECT  id INTO gs_grid4_toolbar-icon             "#EC CI_NOORDER
        UP TO 1 ROWS
        FROM icon WHERE name = <fs_botoes>-icon.
      ENDSELECT.

      READ TABLE lt_zckp_tab3t ASSIGNING FIELD-SYMBOL(<fs_3t>) WITH KEY est_botao  = <fs_botoes>-est_botao
                                                                            num    = <fs_botoes>-num
                                                                            alvhl  = <fs_botoes>-alvhl.
      IF sy-subrc EQ 0.
        gs_grid4_toolbar-quickinfo = <fs_3t>-quickinfo.
        gs_grid4_toolbar-text      = <fs_3t>-text.
      ENDIF.

      MOVE <fs_botoes>-butn_type TO gs_grid4_toolbar-butn_type.
      MOVE <fs_botoes>-disabled TO gs_grid4_toolbar-disabled.

      APPEND gs_grid4_toolbar TO e_object->mt_toolbar.
    ENDLOOP.

  ENDMETHOD.                           "handle_toolbar_status


  METHOD handle_toolbar.

    SORT it_botoes BY function.
    DELETE ADJACENT DUPLICATES FROM it_botoes COMPARING function.
    SORT it_botoes BY num.

    LOOP AT it_botoes ASSIGNING FIELD-SYMBOL(<fs_botoes>) WHERE num_pai EQ space.

      CLEAR gs_grid2_toolbar.
      MOVE <fs_botoes>-function TO gs_grid2_toolbar-function.
      SELECT  id INTO gs_grid2_toolbar-icon             "#EC CI_NOORDER
        UP TO 1 ROWS
        FROM icon WHERE name = <fs_botoes>-icon.
      ENDSELECT.

      READ TABLE lt_zckp_tab3t ASSIGNING FIELD-SYMBOL(<fs_3t>) WITH KEY est_botao  = <fs_botoes>-est_botao
                                                                              alvhl      = <fs_botoes>-alvhl
                                                                              num        = <fs_botoes>-num.
      IF sy-subrc EQ 0.
        gs_grid2_toolbar-quickinfo = <fs_3t>-quickinfo.
        gs_grid2_toolbar-text      = <fs_3t>-text.
      ENDIF.

      MOVE <fs_botoes>-butn_type TO gs_grid2_toolbar-butn_type.
      MOVE <fs_botoes>-disabled TO gs_grid2_toolbar-disabled.
      APPEND gs_grid2_toolbar TO e_object->mt_toolbar.
    ENDLOOP.
  ENDMETHOD.                    "handle_toolbar

  METHOD handle_toolbar_lin.

    LOOP AT it_botoes ASSIGNING FIELD-SYMBOL(<fs_botoes>).
      DELETE e_object->mt_toolbar WHERE function =  <fs_botoes>-function.
    ENDLOOP.

    SORT it_botoes_l BY function.
    DELETE ADJACENT DUPLICATES FROM it_botoes_l COMPARING function.
    SORT it_botoes_l BY num.

*   Carrega botões da tabela de linhas
    LOOP AT it_botoes_l ASSIGNING FIELD-SYMBOL(<fs_botoes1>).

      CLEAR gs_grid3_toolbar.
      MOVE <fs_botoes1>-function TO gs_grid3_toolbar-function.
      SELECT SINGLE id INTO gs_grid3_toolbar-icon       "#EC CI_NOORDER
         FROM icon WHERE name = <fs_botoes1>-icon.


      READ TABLE lt_zckp_tab3t ASSIGNING FIELD-SYMBOL(<fs_3t>) WITH KEY est_botao  = <fs_botoes1>-est_botao
                                                                        alvhl      = <fs_botoes1>-alvhl
                                                                        num        = <fs_botoes1>-num.
      IF sy-subrc EQ 0.
        gs_grid3_toolbar-quickinfo = <fs_3t>-quickinfo.
        gs_grid3_toolbar-text      = <fs_3t>-text.
      ENDIF.

      MOVE <fs_botoes1>-butn_type TO gs_grid3_toolbar-butn_type.
      MOVE <fs_botoes1>-disabled TO gs_grid3_toolbar-disabled.

      APPEND gs_grid3_toolbar TO e_object->mt_toolbar.
    ENDLOOP.

  ENDMETHOD.                    "handle_toolbar_lin


  METHOD handle_menu_button.

    DATA: text  TYPE gui_text.
    DATA: linha    TYPE lvc_t_row,
          id_linha TYPE lvc_t_roid.

* Ler linha seleccionada
    CALL METHOD go_grid_cab->get_selected_rows
      IMPORTING
        et_index_rows = linha
        et_row_no     = id_linha.

* Verificar se opção seleccionada tem filhos
    DATA pai TYPE /sbxc/zckp_tab03-num_pai.

    DATA: lv_function TYPE ui_func.

* Ler código de função (num)
    LOOP AT it_botoes ASSIGNING FIELD-SYMBOL(<fs_botoes>)
      WHERE function = e_ucomm.
      pai = <fs_botoes>-num.
    ENDLOOP.


* Ler filhos
    LOOP AT it_botoes  ASSIGNING FIELD-SYMBOL(<fs_botoes1>)
      WHERE num_pai = pai.

      READ TABLE lt_zckp_tab3t ASSIGNING FIELD-SYMBOL(<fs_3t>) WITH KEY est_botao  = <fs_botoes1>-est_botao
                                                                           num    = <fs_botoes1>-num
                                                                           alvhl  = <fs_botoes1>-alvhl.
      IF sy-subrc EQ 0.
        text = <fs_3t>-quickinfo.
      ENDIF.

      lv_function = <fs_botoes1>-function.
      CALL METHOD e_object->add_function
        EXPORTING
          fcode = lv_function
          text  = text.
    ENDLOOP.

  ENDMETHOD.                    "handle_menu_button

  METHOD handle_user_command.

    DATA: linha_cab      TYPE lvc_t_row,
          linha_lin      TYPE lvc_t_row,
          linha_geral    TYPE lvc_t_row,
          wa_linha       TYPE lvc_s_row,
          id_linha_lin   TYPE lvc_t_roid,
          id_linha_cab   TYPE lvc_t_roid,
          wa_linha_geral TYPE lvc_s_row,
          lin_proc(35),
          lin_ano(35),
          lin_seqno(35),
          ls_row         TYPE lvc_s_row,
          ls_row_l       TYPE lvc_s_row,
          ls_col         TYPE lvc_s_col,
          ls_col_l       TYPE lvc_s_col.
*          flag_lin_empty TYPE c. " Tabela de linhas vazia ao entrar no user command = X

    DATA: n_lines TYPE i.

    go_grid_lin->check_changed_data( ).
    go_grid_cab->check_changed_data( ).

    FIELD-SYMBOLS: <lin_proc>  TYPE any,
                   <lin_ano>   TYPE any,
                   <lin_seqno> TYPE any.
    FIELD-SYMBOLS: <lin_proc2>  TYPE any,
                   <lin_ano2>   TYPE any,
                   <lin_seqno2> TYPE any.

* Menu de contexto
    CLEAR: linha_geral, linha_geral[], refresh_table,
            fname, l_status_final, funcao.

* Ver se seleccionou cabeçalho ou linha
    READ TABLE it_botoes TRANSPORTING NO FIELDS
    WITH KEY function = e_ucomm.

    IF sy-subrc EQ 0.
      lin_or_cab = 'CAB'.
    ELSE.
      READ TABLE it_botoes_l TRANSPORTING NO FIELDS
      WITH KEY function = e_ucomm.
      IF sy-subrc EQ 0.
        lin_or_cab = 'LIN'.
      ELSE.
        READ TABLE it_botoes_s TRANSPORTING NO FIELDS
        WITH KEY function = e_ucomm.
        lin_or_cab = 'STA'.
      ENDIF.

    ENDIF.

* Ler Linha(s) Seleccionadas
    IF r_click IS INITIAL.
      IF lin_or_cab = 'CAB'.
        CALL METHOD go_grid_cab->get_selected_rows
          IMPORTING
            et_index_rows = linha_cab
            et_row_no     = id_linha_cab.
        LOOP AT linha_cab INTO wa_linha.
          MOVE-CORRESPONDING wa_linha TO wa_linha_geral.
          APPEND wa_linha_geral TO linha_geral.
        ENDLOOP.

        IF linha_geral[] IS INITIAL.
          CALL METHOD go_grid_cab->get_current_cell
            IMPORTING
              es_row_id = ls_row
              es_col_id = ls_col.

          APPEND ls_row TO linha_geral.
        ENDIF.


      ELSEIF lin_or_cab = 'LIN'.
        CALL METHOD go_grid_lin->get_selected_rows
          IMPORTING
            et_index_rows = linha_lin
            et_row_no     = id_linha_lin.

        LOOP AT linha_lin INTO wa_linha.
          MOVE-CORRESPONDING wa_linha TO wa_linha_geral.
          APPEND wa_linha_geral TO linha_geral.
        ENDLOOP.

        IF linha_geral[] IS INITIAL.
          CALL METHOD go_grid_lin->get_current_cell
            IMPORTING
              es_row_id = ls_row_l
              es_col_id = ls_col_l.

          APPEND ls_row_l TO linha_geral.
        ENDIF.

      ELSEIF lin_or_cab = 'STA'.

        CALL METHOD go_grid_cab->get_current_cell
          IMPORTING
            es_row_id = ls_row
            es_col_id = ls_col.

        APPEND ls_row TO linha_geral.
      ENDIF.

    ELSE.
      IF lin_or_cab = 'CAB'.
        CALL METHOD go_grid_cab->get_current_cell
          IMPORTING
            es_row_id = ls_row
            es_col_id = ls_col.

        APPEND ls_row TO linha_geral.
      ELSE.
        CALL METHOD go_grid_lin->get_current_cell
          IMPORTING
            es_row_id = ls_row_l
            es_col_id = ls_col_l.

        APPEND ls_row_l TO linha_geral.
      ENDIF.

      CALL METHOD go_grid_cab->get_selected_rows
        IMPORTING
          et_index_rows = linha_cab
          et_row_no     = id_linha_cab.
    ENDIF.

    IF lin_or_cab EQ 'LIN' AND ls_row_l-index EQ '0000000000' AND linha_lin[] IS INITIAL
    AND e_ucomm NE 'APP'.
      MESSAGE TEXT-010 TYPE 'S' DISPLAY LIKE 'E'.
*      CHECK 1 = 2.
      RETURN.
    ENDIF.

* Verificar estrutura do botão
    LOOP AT linha_geral INTO wa_linha_geral.
      PERFORM estr_botao USING e_ucomm wa_linha_geral-index lin_or_cab
                         CHANGING est_botao fname.
    ENDLOOP.

* Ordenar items de forma decrescente para eliminar do final para o inicio
    SORT linha_geral BY index DESCENDING.

    funcao = fname.

    CONCATENATE fname '_PRE' INTO funcao_pre.
    CONCATENATE fname '_POS' INTO funcao_pos.


* verifica se existem funções
    PERFORM valida_funcoes CHANGING: funcao_pre,
                                     funcao,
                                     funcao_pos.

* verificar se pelo menos uma função está activa
    IF  funcao_pre NE space OR
           funcao NE space OR
           funcao_pos NE space.

*Guardar data de lançamento actual
*    DATA wa_zckp_ctrl LIKE /sbxc/zckp_ctrl.
      FIELD-SYMBOLS: <status1>    TYPE any, <pstng_date> TYPE any, <empresa> TYPE any.
      LOOP AT  <t_cab_table> ASSIGNING <cab>.
        ASSIGN COMPONENT 'STATUS1'    OF STRUCTURE <cab> TO <status1>.
        ASSIGN COMPONENT 'PSTNG_DATE' OF STRUCTURE <cab> TO <pstng_date>.
        ASSIGN COMPONENT 'COMP_CODE' OF STRUCTURE <cab> TO <empresa>.
        CLEAR wa_zckp_ctrl.
        MOVE-CORRESPONDING <cab> TO wa_zckp_ctrl.
        IF <status1> IS ASSIGNED AND <pstng_date> IS ASSIGNED.
          IF <status1> = '0' OR <status1> = '5' OR /sbxc/zckp_ctrl-status1 = ' '.
            IF <pstng_date> IS INITIAL. "ODC - 18_03_2021
*  CCF Ini 23.02.2023 10:03:45
* 130058, Data de Lançamento | Alteração em Massa
* Pretendem que nas empresas 2* a data de lançamento assuma a que colocaram no CKP
              IF <empresa> IS ASSIGNED AND <empresa>(1) = '2'.
              ELSE.
                <pstng_date> = sy-datum.
              ENDIF.
*   CCF Fim 23.02.2023 10:03:45
            ENDIF.



            UPDATE (str_cab) SET pstng_date = <pstng_date>
                      WHERE processo = wa_zckp_ctrl-processo
                        AND ano      = wa_zckp_ctrl-ano
                        AND seqno    = wa_zckp_ctrl-seqno.

*            ENDIF.

          ENDIF.
        ENDIF.
      ENDLOOP.
      COMMIT WORK AND WAIT.
    ELSE.
      RETURN.
    ENDIF.
    CASE e_ucomm .
      WHEN 'MULT'.

        CASE origem.
          WHEN 'CAB'.
            UNASSIGN <cab>.
            LOOP AT  <t_cab_table> ASSIGNING <cab>.
              <valor> = conteudo.
              MODIFY <t_cab_table> FROM <cab>.
*              UNASSIGN <cab>.
            ENDLOOP.
            CALL METHOD go_grid_cab->refresh_table_display(
                is_stable =
                            gc_stable ).

          WHEN 'LIN'.

            UNASSIGN <wa_lin_table>.
            LOOP AT  <t_lin_table> ASSIGNING <wa_lin_table>.
              <valor> = conteudo.
              MODIFY <t_lin_table> FROM <wa_lin_table>.
            ENDLOOP.
            CALL METHOD go_grid_lin->refresh_table_display(
                is_stable =
                            gc_stable ).

          WHEN OTHERS.

        ENDCASE.

      WHEN OTHERS.

        FIELD-SYMBOLS: <wa_cab_table1> TYPE any,
                       <wa_cab_table2> TYPE any.

        IF lin_or_cab EQ 'CAB'.

          ASSIGN r_wa_dyn_table_cab->* TO <cab>.

          SORT linha_geral BY index DESCENDING.
          DELETE ADJACENT DUPLICATES FROM linha_geral.
          LOOP AT linha_geral INTO wa_linha. " Linhas seleccionadas

            CLEAR erro_status.
            UNASSIGN <cab>.

            READ TABLE  <t_cab_table> ASSIGNING <cab>
                                      INDEX wa_linha-index.
            IF <cab> IS ASSIGNED.
              MOVE-CORRESPONDING <cab> TO key.
            ENDIF.
            CLEAR wa_zckp_ctrl.
            SELECT SINGLE * FROM /sbxc/zckp_ctrl
              INTO wa_zckp_ctrl
              WHERE processo =  key-processo
                AND ano = key-ano
                AND seqno = key-seqno.

            MOVE wa_zckp_ctrl-status1 TO status1_tmp.


* Se linha não preenchida com dados correspondentes ao cabeçalho lê da base de dados
*            CLEAR: flag_lin_empty.
            IF <t_lin_table> IS INITIAL.

*              flag_lin_empty = 'X'.

              SELECT * FROM (str_lin)
                  APPENDING CORRESPONDING FIELDS OF TABLE <t_lin_table>
                   WHERE processo = key-processo
                      AND ano = key-ano
                      AND seqno = key-seqno.

            ELSE.

* Verificar se linhas pertencem ao cabeçalho
              IF NOT <wa_lin_table> IS ASSIGNED.
                ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
              ENDIF.

              DATA erro.
              erro = 'X'.
              UNASSIGN <wa_lin_table>.
              LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.

                lin_proc = '<WA_LIN_TABLE>-PROCESSO'.
                lin_ano = '<WA_LIN_TABLE>-ANO'.
                lin_seqno = '<WA_LIN_TABLE>-SEQNO'.

                ASSIGN (lin_proc) TO <lin_proc>.
                ASSIGN (lin_ano) TO <lin_ano>.
                ASSIGN (lin_seqno) TO <lin_seqno>.

                IF <lin_proc> = key-processo AND
                   <lin_ano> = key-ano AND
                   <lin_seqno> = key-seqno.
                  erro = ' '.
                ENDIF.
              ENDLOOP.

              IF erro = 'X'.
*        As linhas não pertenciam ao cabeçalho. Ler da BD
                REFRESH <t_lin_table>.
                SELECT * FROM (str_lin)
                 APPENDING CORRESPONDING FIELDS OF TABLE <t_lin_table>
                   WHERE processo = key-processo
                     AND ano = key-ano
                     AND seqno = key-seqno.
              ENDIF.
            ENDIF.

            IF lin_or_cab = 'CAB'.
              CLEAR idx_lin.
            ELSE.
              idx_lin = wa_linha-index.
            ENDIF.

            PERFORM funcao_cabecalho.

            AT FIRST.
              IF funcao_pre NE space. " função existe

                CLEAR erro_status.
                IF <cab> IS  ASSIGNED.
                  PERFORM check_status_final USING l_status_final
                                                 <cab>
                                        CHANGING erro_status.

                  CHECK erro_status IS INITIAL.

                  CALL FUNCTION funcao_pre
                    EXPORTING
                      e_ucomm = e_ucomm
                      idx_lin = idx_lin
                    IMPORTING
                      refresh = refresh_table
                    TABLES
                      cor     = f_cor       " Cor para células
                      linha   = <t_lin_table>
                      status  = gt_status
                    CHANGING
                      cab     = <cab>
                      ctrl    = wa_zckp_ctrl.
                ENDIF.
              ENDIF.

            ENDAT.

            IF funcao NE space.
              CLEAR erro_status.
              IF <cab> IS  ASSIGNED.
                PERFORM check_status_final USING l_status_final
                                                 <cab>
                                        CHANGING erro_status.

                CHECK erro_status IS INITIAL.

                DESCRIBE TABLE linha_cab LINES n_lines.
                EXPORT n_lines FROM n_lines TO MEMORY ID 'N_LINES'.

*       Módulo de função associado ao botão
                CALL FUNCTION funcao
                  EXPORTING
                    e_ucomm = e_ucomm
                    idx_lin = idx_lin
                  IMPORTING
                    refresh = refresh_table
                  TABLES
                    cor     = f_cor       " Cor para células
                    linha   = <t_lin_table>
                    status  = gt_status
                  CHANGING
                    cab     = <cab>
                    ctrl    = wa_zckp_ctrl.

                IF wa_zckp_ctrl-status1 NE status1_tmp.
                  "Alterar status no cabeçalho
                  FIELD-SYMBOLS: <status_cab> TYPE any.
                  DATA l_status_cab(40).

                  l_status_cab = '<CAB>-STATUS1'.
                  ASSIGN (l_status_cab) TO <status_cab>.
                  <status_cab> = wa_zckp_ctrl-status1.

                  MODIFY <t_cab_table> FROM <cab> INDEX wa_linha-index.
                ENDIF.

                REFRESH f_cor.
                CALL FUNCTION wa_fm_display-fm_cab_disp
                  EXPORTING
                    status1_tmp = status1_tmp
                    est_cab     = /sbxc/zckp_tab00-est_cab
                    t_tab12     = lt_tab12
                    t_hmail     = lt_hmail
                  TABLES
                    cor         = f_cor       " Cor para células
                  CHANGING
                    cab         = <cab>
                    ctrl        = wa_zckp_ctrl.
              ENDIF.
              LOOP AT f_cor INTO w_cor.
                w_cor-tabix = wa_linha-index. " linha de cabeçalho
                MODIFY f_cor FROM w_cor.
              ENDLOOP.

* Estilos
              IF <t_lin_table>[] IS NOT INITIAL AND wa_linha-index IS NOT INITIAL.
                l_status_cab = '<cab>-status1'.
                ASSIGN (l_status_cab) TO <status_cab>.
                IF <wa_cab_table> IS ASSIGNED.
                  ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_cab_table> TO <fs_style>.
                ENDIF.
                PERFORM estilo_campos USING 'CAB' wa_zckp_ctrl-status1.

                IF <cab> IS ASSIGNED.
                  ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <cab> TO <fs_color>.
                ENDIF.
                PERFORM cor_campos USING 'CAB' wa_linha-index.
                MODIFY <t_cab_table> FROM <cab> INDEX wa_linha-index.
              ENDIF.

              l_status_cab = '<CAB>-STATUS1'.
              ASSIGN (l_status_cab) TO <status_cab>.

              IF <t_lin_table>[] IS NOT INITIAL.
                UNASSIGN <wa_lin_table>.
                LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
                  ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
                 <fs_style>.
                  IF <status_cab> IS ASSIGNED.
                    PERFORM estilo_campos USING 'LIN' <status_cab>.
                  ENDIF.
                  IF  <wa_lin_table> IS ASSIGNED.
                    ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <wa_lin_table> TO
                    <fs_color>.
                  ENDIF.
                  PERFORM cor_campos USING 'LIN' sy-tabix.

                  MODIFY <t_lin_table> FROM <wa_lin_table> INDEX sy-tabix.
                ENDLOOP.
              ENDIF.

            ENDIF.
            AT LAST.

              IF funcao_pos NE space. " função existe
                IF <cab> IS  ASSIGNED.
                  CLEAR erro_status.
                  PERFORM check_status_final USING l_status_final
                                                 <cab>
                                        CHANGING erro_status.

                  CHECK erro_status IS INITIAL.

                  CALL FUNCTION funcao_pos
                    EXPORTING
                      e_ucomm = e_ucomm
                      idx_lin = idx_lin
                    IMPORTING
                      refresh = refresh_table
                    TABLES
                      cor     = f_cor       " Cor para células
                      linha   = <t_lin_table>
                      status  = gt_status
                    CHANGING
                      cab     = <cab>
                      ctrl    = wa_zckp_ctrl.

                  go_grid_lin->check_changed_data( ).
                  go_grid_cab->check_changed_data( ).
                  gd_changes = 'X'.
                ENDIF.
              ENDIF.
            ENDAT.
          ENDLOOP.
        ELSE.

          IF lin_or_cab EQ 'STA'.
            <t_cab_table>[] = <t_cab_table_tmp>[].
          ENDIF.

          DATA: lv_s_cont TYPE i.
          DESCRIBE TABLE <t_cab_table> LINES lv_s_cont.

*          CHECK lv_s_cont GT 0.
          IF lv_s_cont GT 0.

            DATA: l_cab_proc(30), l_cab_seqno(30), l_cab_ano(30).
            FIELD-SYMBOLS: <cab_proc>  TYPE any, <cab_ano> TYPE any, <cab_seqno> TYPE any.
            DATA: l_cab_index TYPE sy-tabix.
            SORT linha_geral BY index DESCENDING.
            DELETE ADJACENT DUPLICATES FROM linha_geral.
            LOOP AT linha_geral INTO wa_linha. " Linhas seleccionadas

              UNASSIGN <cab>.
              CLEAR l_cab_index.
              LOOP AT <t_cab_table> ASSIGNING <cab> .
                l_cab_proc = '<CAB>-PROCESSO'.
                ASSIGN (l_cab_proc) TO <cab_proc>.

                l_cab_ano = '<CAB>-ANO'.
                ASSIGN (l_cab_ano) TO <cab_ano>.

                l_cab_seqno = '<CAB>-SEQNO'.
                ASSIGN (l_cab_seqno) TO <cab_seqno>.

                IF <cab_proc> EQ key-processo AND
                   <cab_ano> EQ key-ano AND
                   <cab_seqno> EQ key-seqno.

                  l_cab_index = sy-tabix.
                  EXIT.
                ENDIF.
              ENDLOOP.

              CLEAR wa_zckp_ctrl.
              SELECT SINGLE * FROM /sbxc/zckp_ctrl
                INTO wa_zckp_ctrl
                WHERE processo =  key-processo
                  AND ano = key-ano
                  AND seqno = key-seqno.

              MOVE wa_zckp_ctrl-status1 TO status1_tmp.


              IF lin_or_cab = 'CAB'.
                CLEAR idx_lin.
              ELSE.
                idx_lin = wa_linha-index.
              ENDIF.

              PERFORM funcao_cabecalho.

              AT FIRST.

                IF funcao_pre NE space. " função existe
                  IF <cab> IS  ASSIGNED.
                    CLEAR erro_status.
                    PERFORM check_status_final USING l_status_final
                                                   <cab>
                                          CHANGING erro_status.

                    CHECK erro_status IS INITIAL.

                    CALL FUNCTION funcao_pre
                      EXPORTING
                        e_ucomm = e_ucomm
                        idx_lin = idx_lin
                      IMPORTING
                        refresh = refresh_table
                      TABLES
                        cor     = f_cor       " Cor para células
                        linha   = <t_lin_table>
                        status  = gt_status
                      CHANGING
                        cab     = <cab>
                        ctrl    = wa_zckp_ctrl.
                  ENDIF.
                ENDIF.
              ENDAT.

              IF funcao NE space.
                IF <cab> IS  ASSIGNED.
                  CLEAR erro_status.

                  PERFORM check_status_final USING l_status_final
                                                   <cab>
                                          CHANGING erro_status.

                  CHECK erro_status IS INITIAL.

*       Módulo de função associado ao botão
                  CALL FUNCTION funcao
                    EXPORTING
                      e_ucomm = e_ucomm
                      idx_lin = idx_lin
                    IMPORTING
                      refresh = refresh_table
                    TABLES
                      cor     = f_cor       " Cor para células
                      linha   = <t_lin_table>
                      status  = gt_status
                    CHANGING
                      cab     = <cab>
                      ctrl    = wa_zckp_ctrl.

                  IF wa_zckp_ctrl-status1 NE status1_tmp.
                    "Alterar status no cabeçalho
                    l_status_cab = '<CAB>-STATUS1'.
                    ASSIGN (l_status_cab) TO <status_cab>.
                    <status_cab> = wa_zckp_ctrl-status1.
                  ENDIF.


                  REFRESH f_cor.
                  CALL FUNCTION wa_fm_display-fm_cab_disp
                    EXPORTING
                      status1_tmp = status1_tmp
                      est_cab     = /sbxc/zckp_tab00-est_cab
                      t_tab12     = lt_tab12
                      t_hmail     = lt_hmail
                    TABLES
                      cor         = f_cor       " Cor para células
                    CHANGING
                      cab         = <cab>
                      ctrl        = wa_zckp_ctrl.
                ENDIF.
                LOOP AT f_cor INTO w_cor.
                  w_cor-tabix = wa_linha-index. " linha de cabeçalho
                  MODIFY f_cor FROM w_cor.
                ENDLOOP.

* Estilos
                IF <t_lin_table>[] IS NOT INITIAL AND l_cab_index IS NOT INITIAL.

                  l_status_cab = '<CAB>-STATUS1'.
                  ASSIGN (l_status_cab) TO <status_cab>.
                  IF <wa_cab_table> IS ASSIGNED AND <fs_style> IS ASSIGNED.
                    ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_cab_table> TO
                    <fs_style>.
                  ENDIF.
                  PERFORM estilo_campos USING 'CAB' <status_cab>.

                  IF <cab> IS ASSIGNED.
                    ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <cab> TO <fs_color>.
                  ENDIF.
                  PERFORM cor_campos USING 'CAB' l_cab_index.
                  MODIFY <t_cab_table> FROM <cab> INDEX l_cab_index.
                ENDIF.

                l_status_cab = '<CAB>-STATUS1'.
                ASSIGN (l_status_cab) TO <status_cab>.

                IF <t_lin_table>[] IS NOT INITIAL.

                  UNASSIGN <wa_lin_table>.
                  LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
                    ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
                   <fs_style>.
                    PERFORM estilo_campos USING 'LIN' <status_cab>.
                    ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <wa_lin_table> TO
                    <fs_color>.
                    PERFORM cor_campos USING 'LIN' sy-tabix.

                    MODIFY <t_lin_table> FROM <wa_lin_table> INDEX sy-tabix.

                  ENDLOOP.
                ENDIF.

              ENDIF.

              AT LAST.
                IF funcao_pos NE space. " função existe
                  IF <cab> IS  ASSIGNED.
                    CLEAR erro_status.
                    PERFORM check_status_final USING l_status_final
                                                   <cab>
                                          CHANGING erro_status.

                    CHECK erro_status IS INITIAL.

                    CALL FUNCTION funcao_pos
                      EXPORTING
                        e_ucomm = e_ucomm
                        idx_lin = idx_lin
                      IMPORTING
                        refresh = refresh_table
                      TABLES
                        cor     = f_cor       " Cor para células
                        linha   = <t_lin_table>
                        status  = gt_status
                      CHANGING
                        cab     = <cab>
                        ctrl    = wa_zckp_ctrl.

                    go_grid_lin->check_changed_data( ).
                    go_grid_cab->check_changed_data( ).
                    gd_changes = 'X'.
                  ENDIF.
                ENDIF.
              ENDAT.

            ENDLOOP.
          ELSE.
            RETURN.
          ENDIF.
        ENDIF.

        IF refresh_table IS INITIAL.

        ELSE.

          IF refresh_table EQ 'D'. "Vem de um função da tababela STATUS
            "Função Desmarca tudo
            REFRESH <t_cab_table>.
            CLEAR refresh_table.
            gs_layout-stylefname = 'CELL_COLOR'.
            gs_layout-stylefname = 'CELLSTYLE'.
            CLEAR r_click.

            CALL METHOD go_grid_lin->set_frontend_layout
              EXPORTING
                is_layout = gs_layout.

            CALL METHOD go_grid_lin->refresh_table_display(
                is_stable =
                            gc_stable ).
            CALL METHOD go_grid_cab->refresh_table_display(
                is_stable =
                            gc_stable ).
            CALL METHOD go_grid_status->refresh_table_display(
                is_stable =
                            gc_stable ).

            RETURN.
          ENDIF.

          REFRESH: <t_cab_table>.
          SELECT * FROM /sbxc/zckp_ctrl
            INTO TABLE it_ctrl
            WHERE processo IN p_proc
                  AND ano IN p_ano
                  AND seqno IN p_seqno
                  AND data_in IN p_data
                  AND hora_in IN p_hora
                  AND user_in IN p_user
                  AND status1 IN p_stout.

          SORT it_ctrl BY processo ano seqno.
          LOOP AT it_ctrl ASSIGNING FIELD-SYMBOL(<fs_ctrl>)."INTO wa_zckp_ctrl.
            MOVE-CORRESPONDING <fs_ctrl> TO /sbxc/zckp_ctrl.

            IF sy-subrc EQ 0.
              READ TABLE gt_tab10 INTO gs_tab10
              WITH KEY processo = /sbxc/zckp_ctrl-processo
              status = /sbxc/zckp_ctrl-status1.
              /sbxc/zckp_ctrl-icon_status = gs_tab10-icon_status.
            ENDIF.

            IF refresh_table EQ 'I' OR refresh_table = 'X'.
              "Função Inverter STATUS

              IF sy-subrc NE 0.
                READ TABLE gt_status TRANSPORTING NO FIELDS
                              WITH KEY box = 'X'
                                       status = /sbxc/zckp_ctrl-status1.
                IF NOT <cab> IS ASSIGNED.
                  ASSIGN r_wa_dyn_table_cab->* TO <cab>.
                ENDIF.
                MOVE-CORRESPONDING /sbxc/zckp_ctrl  TO <cab>.
                SELECT SINGLE * FROM (str_cab)
              INTO CORRESPONDING FIELDS OF <cab>
              WHERE
                processo = /sbxc/zckp_ctrl-processo
                AND ano = /sbxc/zckp_ctrl-ano
                AND seqno = /sbxc/zckp_ctrl-seqno
                AND comp_code IN p_bukrs
                AND ref_doc_no IN p_refdoc
                AND vendor  IN p_vendor
                AND doc_fi IN p_docfi.

              ENDIF.
              CHECK sy-subrc EQ 0.
            ENDIF.

            IF NOT <cab> IS ASSIGNED.
              ASSIGN r_wa_dyn_table_cab->* TO <cab>.
            ENDIF.

            MOVE-CORRESPONDING /sbxc/zckp_ctrl  TO <cab>.
            SELECT SINGLE * FROM (str_cab)
            INTO CORRESPONDING FIELDS OF <cab>
            WHERE processo = /sbxc/zckp_ctrl-processo
              AND ano = /sbxc/zckp_ctrl-ano
              AND seqno = /sbxc/zckp_ctrl-seqno
              AND comp_code IN p_bukrs
              AND ref_doc_no IN p_refdoc
              AND vendor  IN p_vendor
              AND doc_fi IN p_docfi.

            CHECK sy-subrc = 0.
            MOVE-CORRESPONDING /sbxc/zckp_ctrl TO <cab>.
*Preencher status OUT
            FIELD-SYMBOLS: <status2>    TYPE any, <status_out> TYPE any, <mensagem> TYPE any.

            UNASSIGN: <status2>, <status_out>.
            ASSIGN COMPONENT 'STATUS_OUT' OF STRUCTURE <cab> TO <status_out>.
            ASSIGN COMPONENT 'STATUS1' OF STRUCTURE <cab> TO <status2>.
            IF <status_out> IS ASSIGNED AND <status2> IS  ASSIGNED.
              <status_out> = gs_tab10-status_out.
            ENDIF.

            UNASSIGN <pstng_date>.
            ASSIGN COMPONENT 'PSTNG_DATE' OF STRUCTURE <cab> TO <pstng_date>.
            ASSIGN COMPONENT 'COMP_CODE' OF STRUCTURE <cab> TO <empresa>.
            IF <pstng_date> IS ASSIGNED.
              IF <pstng_date> IS INITIAL. "ODC - 18_03_2021
*  CCF Ini 23.02.2023 10:03:45
* 130058, Data de Lançamento | Alteração em Massa
* Pretendem que nas empresas 2* a data de lançamento assuma a que colocaram no CKP
                IF <empresa> IS ASSIGNED AND <empresa>(1) = '2'.
                ELSE.
                  <pstng_date> = sy-datum.
                ENDIF.
*   CCF Fim 23.02.2023 10:03:45
              ENDIF.

            ENDIF.

            UNASSIGN <mensagem>.
            ASSIGN COMPONENT 'MENSAGEM' OF STRUCTURE <cab> TO <mensagem>.
            IF <mensagem> IS ASSIGNED.
              CLEAR <mensagem>.

              READ TABLE lt_tab09 ASSIGNING FIELD-SYMBOL(<fs_09>)
                    WITH KEY processo = /sbxc/zckp_ctrl-processo
                             ano      = /sbxc/zckp_ctrl-ano
                             seqno    = /sbxc/zckp_ctrl-seqno.
              IF sy-subrc EQ 0.
                <mensagem> = <fs_09>-mensagem.
              ENDIF.
*              SELECT SINGLE mensagem FROM /sbxc/zckp_tab09  INTO <mensagem>
*                                  WHERE processo = /sbxc/zckp_ctrl-processo
*                                    AND ano      = /sbxc/zckp_ctrl-ano
*                                    AND seqno    = /sbxc/zckp_ctrl-seqno.
            ENDIF.

            ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <cab> TO <fs_style>.
            PERFORM estilo_campos USING 'CAB' /sbxc/zckp_ctrl-status1.

            ASSIGN COMPONENT 'CELL_COLOR' OF STRUCTURE <cab> TO <fs_color>.
            PERFORM funcao_cabecalho_ini.


            PERFORM cor_campos USING 'CAB' '1'.
*

*Doc_fi************************************************************ini
*stat 3 ou 4
            DATA: lv_campo(80), wa_bkpf TYPE bkpf.

            FIELD-SYMBOLS: <doc_fi>     TYPE any,
                           <comp_code>  TYPE any,
                           <ano>        TYPE any,
                           <ref_doc_no> TYPE any.

            IF /sbxc/zckp_ctrl-status1 EQ '3' OR /sbxc/zckp_ctrl-status1 EQ '4'.
              MOVE '<CAB>-DOC_FI' TO lv_campo.
              ASSIGN (lv_campo) TO <doc_fi>.

              IF <doc_fi> IS ASSIGNED AND <doc_fi> IS INITIAL.

                MOVE '<CAB>-COMP_CODE' TO lv_campo.
                ASSIGN (lv_campo) TO <comp_code>.
                MOVE '<CAB>-ANO' TO lv_campo.
                ASSIGN (lv_campo) TO <ano>.
                MOVE '<CAB>-REF_DOC_NO' TO lv_campo.
                ASSIGN (lv_campo) TO <ref_doc_no>.

                CLEAR wa_bkpf.
                SELECT SINGLE belnr awkey gjahr cpudt FROM bkpf "#EC CI_NOORDER
                 INTO (wa_bkpf-belnr, wa_bkpf-awkey, wa_bkpf-gjahr, wa_bkpf-cpudt) WHERE
                   bukrs = <comp_code> AND
                   gjahr = <ano> AND
                   xblnr = <ref_doc_no>.

                UPDATE (str_cab) SET
                           doc_fi = wa_bkpf-belnr
                           ano_lanc = wa_bkpf-gjahr
                           data_criacao = wa_bkpf-cpudt
                           doc_lo = wa_bkpf-awkey(10)
                     WHERE processo = /sbxc/zckp_ctrl-processo AND
                           ano = /sbxc/zckp_ctrl-ano AND
                           seqno = /sbxc/zckp_ctrl-seqno.

                <doc_fi> = wa_bkpf-belnr.

              ENDIF.

            ENDIF.
*Doc_fi************************************************************fim
* CCF 02.10.2020 especifico SF calculo data base de pagamento
            DATA: processo      LIKE /sbxc/zckp_invh-processo, seqno LIKE /sbxc/zckp_invh-seqno,
                  baseline_date LIKE /sbxc/zckp_invh-baseline_date.
            FIELD-SYMBOLS: <baseline_date> TYPE any.
            processo = /sbxc/zckp_ctrl-processo.
            seqno = /sbxc/zckp_ctrl-seqno.

            CALL FUNCTION '/SBXC/ZCKP_BASELINE_DATE'
              EXPORTING
                processo      = processo
                seqno         = seqno
              CHANGING
                baseline_date = baseline_date.

            ASSIGN COMPONENT 'BASELINE_DATE' OF STRUCTURE <cab> TO <baseline_date>.
            <baseline_date> = baseline_date.

* fim 02.10.2021

            APPEND <cab> TO <t_cab_table>.

          ENDLOOP.

          IF refresh_table = 'X'.
**********************************************************************
* Atualiza tabela STATUS
            PERFORM get_tab_status.
**********************************************************************
          ENDIF.
          CLEAR refresh_table.
        ENDIF.

        gs_layout-stylefname = 'CELL_COLOR'.

        gs_layout-stylefname = 'CELLSTYLE'.

**********************************************************************
        IF lin_or_cab NE 'STA' AND lin_or_cab NE 'CAB'.
          CLEAR: <t_cab_table_tmp>, <t_cab_table_tmp>[].
          <t_cab_table_tmp>[] = <t_cab_table>[].
        ELSEIF lin_or_cab EQ 'CAB'.
*
          LOOP AT <t_cab_table> ASSIGNING <wa_cab_table2>.

            lin_proc = '<WA_CAB_TABLE2>-PROCESSO'.
            lin_ano = '<WA_CAB_TABLE2>-ANO'.
            lin_seqno = '<WA_CAB_TABLE2>-SEQNO'.
*
            ASSIGN (lin_proc) TO <lin_proc>.
            ASSIGN (lin_ano) TO <lin_ano>.
            ASSIGN (lin_seqno) TO <lin_seqno>.
            LOOP AT <t_cab_table_tmp> ASSIGNING <wa_cab_table1>.
              lin_proc = '<WA_CAB_TABLE1>-PROCESSO'.
              lin_ano = '<WA_CAB_TABLE1>-ANO'.
              lin_seqno = '<WA_CAB_TABLE1>-SEQNO'.
*
              ASSIGN (lin_proc) TO <lin_proc2>.
              ASSIGN (lin_ano) TO <lin_ano2>.
              ASSIGN (lin_seqno) TO <lin_seqno2>.
              IF <lin_proc> EQ <lin_proc2> AND <lin_ano> EQ <lin_ano2> AND <lin_seqno> EQ <lin_seqno2>.
                MOVE-CORRESPONDING <wa_cab_table2> TO <wa_cab_table1>.
                EXIT.
              ENDIF.
            ENDLOOP.
          ENDLOOP.
        ENDIF.
**********************************************************************

    ENDCASE.

*Obter items para a linha selecionada

    IF e_ucomm EQ 'INT' OR
      e_ucomm EQ 'INT2' OR
      e_ucomm EQ 'FBR2' OR
      e_ucomm EQ 'FB01' OR
      e_ucomm EQ 'F-43' OR
      e_ucomm EQ 'F-41' OR
      e_ucomm EQ 'BAPI'.
      REFRESH <t_lin_table>.
      IF NOT <wa_lin_table> IS ASSIGNED.
        ASSIGN r_wa_dyn_table_lin->* TO <wa_lin_table>.
      ENDIF.

      SELECT * FROM (str_lin)
      INTO CORRESPONDING FIELDS OF <wa_lin_table>
      WHERE processo = key-processo
          AND ano = key-ano
          AND seqno = key-seqno.
        APPEND <wa_lin_table> TO <t_lin_table>.
      ENDSELECT.
    ENDIF.

    FIELD-SYMBOLS <xware_bnk> TYPE any.
    LOOP AT <t_cab_table> ASSIGNING <cab>.
      ASSIGN COMPONENT 'XWARE_BNK' OF STRUCTURE <cab> TO <xware_bnk>.
      IF <xware_bnk> IS INITIAL.
        <xware_bnk> = '1'.
        MODIFY <t_cab_table> FROM <cab>.
      ENDIF.
    ENDLOOP.

    CALL METHOD go_grid_lin->refresh_table_display(
        is_stable = gc_stable ).

    UNASSIGN <wa_lin_table>.
    LOOP AT <t_lin_table> ASSIGNING <wa_lin_table>.
      lin_proc = '<WA_LIN_TABLE>-PROCESSO'.
      lin_ano = '<WA_LIN_TABLE>-ANO'.
      lin_seqno = '<WA_LIN_TABLE>-SEQNO'.

      ASSIGN (lin_proc) TO <lin_proc>.
      ASSIGN (lin_ano) TO <lin_ano>.
      ASSIGN (lin_seqno) TO <lin_seqno>.

      SELECT SINGLE * FROM /sbxc/zckp_ctrl INTO @DATA(ls_ctrl)
        WHERE processo = @<lin_proc>
              AND ano = @<lin_ano>
              AND seqno = @<lin_seqno>.
      MOVE-CORRESPONDING ls_ctrl TO <cabecalho>.

      ASSIGN COMPONENT 'CELLSTYLE' OF STRUCTURE <wa_lin_table> TO
                  <fs_style>.
      PERFORM estilo_campos USING 'LIN' ls_ctrl-status1.
      MODIFY <t_lin_table> FROM <wa_lin_table> INDEX sy-tabix.
    ENDLOOP.

    CLEAR r_click.

    CALL METHOD go_grid_lin->set_frontend_layout
      EXPORTING
        is_layout = gs_layout.

    CALL METHOD go_grid_lin->refresh_table_display(
        is_stable =
                    gc_stable ).
    CALL METHOD go_grid_cab->refresh_table_display(
        is_stable =
                    gc_stable ).
    CALL METHOD go_grid_status->refresh_table_display(
        is_stable =
                    gc_stable ).

    CALL METHOD go_grid_cab->set_selected_rows
      EXPORTING
        it_index_rows            = linha_cab
        it_row_no                = id_linha_cab
        is_keep_other_selections = 'X'.

  ENDMETHOD.                           "handle_user_command


ENDCLASS.                    "lcl_event_toolbar IMPLEMENTATION
*&---------------------------------------------------------------------*
*&      Form  DOUBLE_CLICK_LIN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_E_COLUMN  text
*      -->P_E_ROW  text
*----------------------------------------------------------------------*
FORM double_click_lin  USING    e_column TYPE any
                                e_row TYPE any.
  DATA field TYPE dd03l-fieldname.

  field = e_column.
  ASSIGN r_wa_dyn_table_cab->* TO <cab>.
  READ TABLE <t_lin_table> ASSIGNING <wa_lin_table> INDEX e_row.

  CALL FUNCTION '/SBXC/ZCKP_DBL_CLK_LIN'
    EXPORTING
      cab   = <cab>
      campo = field
      linha = <wa_lin_table>
    TABLES
      cor   = f_cor.
ENDFORM.                    " DOUBLE_CLICK_LIN
*&---------------------------------------------------------------------*
*&      Form  CHECK_STATUS_FINAL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM check_status_final USING p_status_final TYPE any
                              p_cab TYPE any
                     CHANGING p_erro TYPE any.

  IF p_status_final IS NOT INITIAL.

    DATA: l_status1(35),
          t_status_i    TYPE /sbxc/text_status,
          t_status_f    TYPE /sbxc/text_status.

    FIELD-SYMBOLS: <l_status_i> TYPE any.

    l_status1 = 'P_CAB-STATUS1'.
    ASSIGN (l_status1) TO <l_status_i>.

    SELECT SINGLE status_dep FROM /sbxc/zckp_tab11 " Status seguintes não permitidos
      INTO @DATA(lv_status_dep)
      WHERE processo = @key-processo
        AND status = @<l_status_i>
        AND status_dep = @p_status_final
        AND disabled = @space.

    IF sy-subrc EQ 0.

      SELECT status, text_status FROM /sbxc/zckptab10t
      INTO (@DATA(lv_status), @DATA(lv_text_status))
      WHERE spras = @sy-langu
        AND ( status = @<l_status_i>
        OR status = @p_status_final ).

        IF lv_status EQ <l_status_i>.
          t_status_i = lv_text_status.
        ELSE.
          t_status_f = lv_text_status.
        ENDIF.
      ENDSELECT.

      MESSAGE s005 WITH <l_status_i> t_status_i p_status_final
      t_status_f DISPLAY LIKE 'E'.
      p_erro = 'X'.
    ENDIF.
  ELSE.
    RETURN.
  ENDIF.
ENDFORM.                    " CHECK_STATUS_FINAL
**&---------------------------------------------------------------------*
**&      Form  CONVERT_AMOUNT
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**  -->  p1        text
**  <--  p2        text
**----------------------------------------------------------------------*
*FORM convert_amount CHANGING p_value.
*
*  DATA: l_dcpfm TYPE usr01-dcpfm.
*
*  SELECT SINGLE dcpfm FROM usr01
*    INTO l_dcpfm
*    WHERE bname = v_uname.
*
*  CASE l_dcpfm.
*    WHEN ' '. "1.234.567,89
*      TRANSLATE p_value USING '. '.
*      TRANSLATE p_value USING ',.'.
*      CONDENSE p_value NO-GAPS.
*    WHEN 'X'. "1,234,567.89
*      TRANSLATE p_value USING ', '.
*      CONDENSE p_value NO-GAPS.
*    WHEN 'Y'. "1 234 567,89
*      TRANSLATE p_value USING ',.'.
*      CONDENSE p_value NO-GAPS.
*  ENDCASE.
*
*ENDFORM.                    " CONVERT_AMOUNT
*&---------------------------------------------------------------------*
*&      Form  ESTR_BOTAO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_E_UCOMM  text
*      -->P_L_INDEX  text
*      -->P_LIN_OR_CAB  text
*      <--P_EST_BOTAO  text
*      <--P_FNAME  text
*----------------------------------------------------------------------*
FORM estr_botao  USING    p_e_ucomm TYPE any
                          p_l_index TYPE lvc_s_row-index
                          p_lin_or_cab TYPE char03
                 CHANGING p_est_botao TYPE /sbxc/zckp_tab03-est_botao
                          p_fname TYPE rs38l-name.

  FIELD-SYMBOLS: <fs_botoes> TYPE gt_botoes.
  READ TABLE it_botoes ASSIGNING <fs_botoes>
  WITH KEY function   = p_e_ucomm.

  IF sy-subrc NE 0.
    READ TABLE it_botoes_l ASSIGNING <fs_botoes>
    WITH KEY function   = p_e_ucomm.
  ENDIF.

  IF sy-subrc NE 0.
    READ TABLE it_botoes_s ASSIGNING <fs_botoes>
    WITH KEY function   = p_e_ucomm.
  ENDIF.

  p_fname = <fs_botoes>-fm.
  IF p_fname IS INITIAL AND p_e_ucomm NE 'REFR'.
    MESSAGE s006 WITH p_est_botao DISPLAY LIKE 'E'.
  ENDIF.

ENDFORM.                    " ESTR_BOTAO
