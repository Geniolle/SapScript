FUNCTION /sbxc/zckp_item_del.
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
  DATA: head TYPE /sbxc/zckp_invh.
  DATA: wa_item   TYPE /sbxc/zckp_invi,
        p_idx_lin LIKE sy-index.
  DATA: l_linhas TYPE i.


  refresh = 'X'.
  MOVE idx_lin TO p_idx_lin.

  MOVE-CORRESPONDING cab TO head.
  CHECK ctrl-status1 <> '3' AND ctrl-status1 <> '4'.
  CLEAR: wa_item.
* Preencher estrutura item
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.


  IF  p_idx_lin IS NOT INITIAL.
    DELETE item INDEX p_idx_lin.
  ENDIF.

  REFRESH linha.
  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.


ENDFUNCTION.
