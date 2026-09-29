FUNCTION /sbxc/zckp_item_new.
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
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: head TYPE /sbxc/zckp_invh.
  DATA: wa_item TYPE /sbxc/zckp_invi,
        p_idx_lin LIKE sy-index.
  data: l_linhas type i.


*refresh = 'X'.
  MOVE idx_lin TO p_idx_lin.

  MOVE-CORRESPONDING cab TO head.
  check ctrl-status1 <> '3' and ctrl-status1 <> '4'.
CLEAR: wa_item.
* Preencher estrutura item
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    if wa_item-INVOICE_DOC_ITEM le item-INVOICE_DOC_ITEM.
      wa_item-INVOICE_DOC_ITEM = item-INVOICE_DOC_ITEM.
    endif.
    APPEND item.
  ENDLOOP.




  wa_item-processo = head-processo.
  wa_item-seqno = head-seqno.
  wa_item-ano = head-ano.
  add 1 to wa_item-INVOICE_DOC_ITEM.
  APPEND wa_item TO item.



  REFRESH linha.
  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.


ENDFUNCTION.
