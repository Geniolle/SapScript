
FUNCTION /sbxc/zckp_item_fat.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header TYPE /sbxc/zckp_invh.
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  FIELD-SYMBOLS <status> TYPE any.

  DATA: item_aux2  TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.

  REFRESH item_aux.

  DATA: status1(20).

  status1 = 'CAB-STATUS1'.
  ASSIGN (status1) TO <status>.
* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    IF  item-po_number IS NOT INITIAL.
      SELECT SINGLE txz01 idnlf elikz  erekz bprme matnr FROM ekpo INTO (item-item_text, item-idnlf,
        item-elikz, item-erekz, item-bprme, item-matnr)
                      WHERE ebeln = item-po_number AND
                            ebelp = item-po_item.

    ENDIF.
    APPEND item.
  ENDLOOP.
  READ TABLE item INDEX 1.

*  PERFORM valida_iva TABLES item USING  header.

*  PERFORM valida_qtd_fatura2 TABLES item USING header .
  IF  <status> IS INITIAL OR <status> = '0' OR <status> = '5' OR <status> = '2'.
    APPEND LINES OF item TO item_aux.
    IF item-mrmok IS INITIAL AND  item-po_number IS NOT INITIAL.
      PERFORM valida_qtd_fatura_sc TABLES item USING header .
    ENDIF.
    REFRESH linha.


* valida imputação contabilistica multipla
    IF  item-po_number IS NOT INITIAL.
      PERFORM valida_imp_mult TABLES item_aux USING  header.
    ENDIF.
    SORT item_aux BY po_number po_item.


    READ TABLE item_aux WITH KEY po_number = space.
    IF sy-subrc NE 0.

      DELETE ADJACENT DUPLICATES FROM item_aux COMPARING po_number po_item ref_doc ref_doc_year ref_doc_item invoice_doc_item.
*      "Validar se tem registos repetidos com os campos de referencia vazios.
*      REFRESH item_aux2.
*      item_aux2[] = item_aux[].
*      LOOP AT item_aux WHERE ref_doc IS INITIAL AND ref_doc_year IS INITIAL AND ref_doc_item IS INITIAL.
*        READ TABLE item_aux2 WITH KEY po_number = item_aux-po_number
*                                      po_item   = item_aux-po_item
*                                      invoice_doc_item = item_aux-invoice_doc_item.
*        IF sy-subrc EQ 0.
*          DELETE item_aux.
*        ENDIF.
*      ENDLOOP.
    ENDIF.
    DATA: lv_ind TYPE /sbxc/zckp_invi-invoice_doc_item.

    lv_ind = 0.
    header-valor_controlo = header-gross_amount * -1.
    LOOP AT item_aux.

      SELECT SINGLE bpumn bpumz FROM ekpo INTO (ekpo-bpumn, ekpo-bpumz)
                WHERE
                    ebeln = item_aux-po_number AND
                    ebelp = item_aux-po_item.

      item_aux-bpmng = item_aux-quantity * ekpo-bpumz / ekpo-bpumn.


      MOVE-CORRESPONDING item_aux TO linha.
      item_aux-invoice_doc_item = lv_ind + 1.
      APPEND linha.
      header-valor_controlo = header-valor_controlo + item_aux-item_amount + item_aux-tax_amount.
    ENDLOOP.
    header-valor_controlo = header-valor_controlo + header-del_costs .
  ELSE.
    REFRESH linha.
    header-valor_controlo = header-gross_amount * -1.
    LOOP AT item.

      SELECT SINGLE bpumn bpumz FROM ekpo INTO (ekpo-bpumn, ekpo-bpumz)
                WHERE
                    ebeln = item-po_number AND
                    ebelp = item-po_item.

      item-bpmng = item-quantity * ekpo-bpumz / ekpo-bpumn.

      MOVE-CORRESPONDING item TO linha.
*      item_aux-INVOICE_DOC_ITEM = lv_ind + 1.
      APPEND linha.
      header-valor_controlo = header-valor_controlo +  item-item_amount + item-tax_amount.
    ENDLOOP.
    header-valor_controlo = header-valor_controlo + header-del_costs.
  ENDIF.

 "ODC - 11_06_2020
* Det país da empresa
     SELECT SINGLE land1 INTO @DATA(lv_land)
        FROM t001 WHERE bukrs = @header-comp_code.
"Fim ODC - 11_06_2020

  MOVE-CORRESPONDING  header TO cab.

  "ODC - 11_06_2020
  IF  <status> IS INITIAL OR <status> = '0' OR <status> = '5' OR <status> = '2'.
  LOOP AT linha.
     MOVE-CORRESPONDING linha TO item.
    IF item-tax_code_sap IS NOT INITIAL.

*   Det. taxa iva
     SELECT SINGLE knumh INTO @DATA(lv_knumh)
       FROM a003
      WHERE kappl = 'TX'
        AND aland = @lv_land
        AND mwskz = @item-tax_code_sap.
      IF lv_knumh IS NOT INITIAL.
          SELECT SINGLE kbetr INTO @DATA(lv_kbetr)
             FROM konp
           WHERE knumh = @lv_knumh
             AND kopos = 1.
             item-tax_amount = item-item_amount * ( lv_kbetr / 1000 ).
             item-tax_imposto_sap = lv_kbetr / 10. "04_11_2020
      ENDIF.

      "Preencher EAN
      IF item-matnr IS NOT INITIAL.
        select single n~ean11 into item-ean11
          from mean as n inner join mara as a
          on n~matnr eq a~matnr
          and n~meinh eq a~meins
          where a~matnr eq item-matnr.
      ENDIF.

    MOVE-CORRESPONDING item TO linha.
    MODIFY linha.

    ENDIF.
  ENDLOOP.
  endif.
  "Fim ODC - 11_06_2020

  refresh = 'X'.

ENDFUNCTION.
