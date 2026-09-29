FUNCTION /SBXC/ZCKP_ITEM_COPY.
*"--------------------------------------------------------------------
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
*"--------------------------------------------------------------------


* Estruturas de cabeçalho e linha
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: wa_item TYPE /sbxc/zckp_invi,
        p_idx_lin LIKE sy-index.
data: l_doc_item type /sbxc/zckp_invi-invoice_doc_item.
*refresh = 'X'.
  MOVE idx_lin TO p_idx_lin.
  check ctrl-status1 <> '3' and ctrl-status1 <> '4'.
* Preencher estrutura item
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
     if l_doc_item le item-invoice_doc_item.
      l_doc_item = item-invoice_doc_item.
    endif.
    APPEND item.
  ENDLOOP.

  READ TABLE item INTO wa_item INDEX p_idx_lin.
  IF sy-subrc EQ 0.
    CLEAR: wa_item-item_amount,
           wa_item-invoice_doc_item,
           wa_item-po_number,
           wa_item-po_item.
    wa_item-invoice_doc_item = l_doc_item + 1.
    APPEND wa_item TO item.
  ENDIF.


  REFRESH linha.
  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.


ENDFUNCTION.
