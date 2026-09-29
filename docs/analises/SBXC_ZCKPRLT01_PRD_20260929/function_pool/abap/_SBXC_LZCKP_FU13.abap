FUNCTION /sbxc/zckp_dbl_clk.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(CAMPO) TYPE  DD03L-FIELDNAME
*"     REFERENCE(ANEXO_ON) TYPE  FLAG OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR
*"  CHANGING
*"     REFERENCE(CAB)
*"----------------------------------------------------------------------

  TABLES: t001.
* Estruturas de cabeçalho e linha
  DATA: header LIKE /sbxc/zckp_invh.
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: lv_knumh LIKE a003-knumh, lv_kbetr LIKE konp-kbetr.
  DATA: f_cor TYPE TABLE OF /sbxc/zckp_color WITH HEADER LINE.

  DATA: lt_tax_values TYPE /sbxc/zckp_t_values,
        l_tax_values  TYPE /sbxc/zckp_s_values.
  DATA: n_lines TYPE i.

* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.

  header-valor_controlo = header-gross_amount * -1.

*  IF anexo_on IS INITIAL
*    and campo ne 'VENDOR'and campo ne 'USER_TRATAMENTO' and campo ne 'DOC_FI' and campo ne 'DOC_PAGAMENTO'
*    and campo ne 'DOC_ESTORNO'and campo ne 'DOC_LO'.
*  CALL FUNCTION '/SBXC/ZCKP_IMG'
*    EXPORTING
*      e_ucomm       = sy-ucomm
*    TABLES
*      linha         = linha
*      cor           = COR
*    CHANGING
*      cab           = cab
*            .
*  ENDIF.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
**********************************************************************

    IF header-mwskz IS NOT INITIAL AND item-tax_code_sap IS INITIAL.
      item-tax_code_sap = header-mwskz.
    ENDIF.

    SELECT SINGLE land1 INTO t001-land1
       FROM t001 WHERE bukrs = header-comp_code.

    REFRESH: lt_tax_values.

    SELECT knumh kschl INTO (lv_knumh, l_tax_values-kschl)
              FROM a003
              WHERE kappl = 'TX'
              AND aland = t001-land1
              AND mwskz = item-tax_code_sap.
      IF sy-subrc EQ 0.
        SELECT SINGLE kbetr INTO lv_kbetr
          FROM konp
          WHERE knumh = lv_knumh
          AND kopos = 1.
      ELSE.
        CLEAR lv_kbetr.
      ENDIF.
      l_tax_values-mwskz = item-tax_code_sap.
      l_tax_values-knumh = lv_knumh.
      l_tax_values-kbetr = lv_kbetr.
      IF l_tax_values-kschl <> 'MWVI'.
        APPEND l_tax_values TO lt_tax_values.
      ENDIF.
      CLEAR l_tax_values.
    ENDSELECT.
    CLEAR n_lines.
    DESCRIBE TABLE lt_tax_values LINES n_lines.
    IF n_lines > 1.
      CLEAR item-tax_imposto_sap.
      LOOP AT lt_tax_values INTO l_tax_values.
        item-tax_imposto_sap = item-tax_imposto_sap + l_tax_values-kbetr / 10.
      ENDLOOP.
      IF item-po_number IS  INITIAL.
        item-tax_amount = item-item_amount * ( item-tax_imposto_sap / 100 ).
        header-valor_controlo = header-valor_controlo + item-item_amount + item-tax_amount.
      ENDIF.
    ELSE.
      item-tax_imposto_sap = lv_kbetr / 10.
      IF item-po_number IS  INITIAL.
        item-tax_amount = item-item_amount * ( item-tax_imposto_sap / 100 ).
        header-valor_controlo = header-valor_controlo + item-item_amount + item-tax_amount.
      ENDIF.
    ENDIF.

*    SELECT SINGLE knumh INTO lv_knumh
*       FROM a003
*       WHERE kappl = 'TX'
*       AND aland = t001-land1
*       AND mwskz = item-tax_code_sap.
*
*    SELECT SINGLE kbetr INTO lv_kbetr
*       FROM konp
*       WHERE knumh = lv_knumh
*       AND kopos = 1.
*    item-tax_imposto_sap = lv_kbetr / 10.
*    IF item-po_number IS  INITIAL.
*      item-tax_amount = item-item_amount * ( item-tax_imposto_sap / 100 ).
*      header-valor_controlo = header-valor_controlo + item-item_amount + item-tax_amount.
*    ENDIF.
**********************************************************************
    APPEND item.
  ENDLOOP.
  header-valor_controlo = header-valor_controlo + header-del_costs .

  CASE campo.

    WHEN 'VENDOR'.
      IF header-vendor IS NOT INITIAL.
        PERFORM exibe_forn USING header-vendor  header-comp_code.
      ENDIF.
    WHEN 'USER_TRATAMENTO'.
      IF header-user_tratamento  IS NOT INITIAL.
        PERFORM exibe_user USING header-user_tratamento.
      ENDIF.
    WHEN 'DOC_FI'.
      IF   header-doc_fi <> '@B1@' AND header-doc_fi IS NOT INITIAL.
        PERFORM exibe_doc_fi USING header-doc_fi  header-comp_code header-ano_lanc.
      ELSEIF  header-doc_fi = '@B1@'.
        PERFORM exibe_varios USING header.
      ENDIF.
    WHEN 'DOC_PAGAMENTO'.
      IF header-doc_pagamento IS NOT INITIAL.
        PERFORM exibe_doc_fi USING header-doc_pagamento  header-comp_code header-ano_lanc.
      ENDIF.
    WHEN 'DOC_ESTORNO'.
      IF header-doc_estorno IS NOT INITIAL.
* verifica se doc de fi ou log
        SELECT SINGLE belnr FROM rbkp INTO rbkp-belnr
           WHERE belnr = header-doc_estorno AND
                 bukrs = header-comp_code.

        IF sy-subrc = 0.
          PERFORM exibe_doc_lo USING header-doc_estorno  header-comp_code header-ano_lanc.
        ELSE.
          PERFORM exibe_doc_fi USING header-doc_estorno  header-comp_code header-ano_lanc.
        ENDIF.
      ENDIF.
    WHEN 'DOC_LO'.

      IF   header-doc_lo <> '@B1@' AND header-doc_lo IS NOT INITIAL.
        PERFORM exibe_doc_lo USING header-doc_lo  header-comp_code header-ano_lanc.
      ELSEIF  header-doc_lo = '@B1@'.
        PERFORM exibe_varios USING header.
      ENDIF.

*    WHEN 'ITEM_TEXT'.
*      PERFORM exibe_texto.
*      header-item_text = sgtxt.
    WHEN 'ICON_EMAIL'.
      IF header-icon_email IS NOT INITIAL.
        CALL FUNCTION '/SBXC/ZCKP_READ_EMAIL'
          EXPORTING
            e_ucomm = 'HEMAIL'
*           IDX_LIN =
*   IMPORTING
*           REFRESH =
          TABLES
            linha   = linha
*           COR     =
          CHANGING
            cab     = cab
*           CTRL    =
          .
      ENDIF.
    WHEN OTHERS.
  ENDCASE.

  REFRESH linha.

  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.

  MOVE-CORRESPONDING  header TO cab.

  refresh = 'X'.

ENDFUNCTION.
