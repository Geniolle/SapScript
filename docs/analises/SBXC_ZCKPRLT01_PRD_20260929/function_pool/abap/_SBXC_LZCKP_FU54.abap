FUNCTION /sbxc/zckp_altera_massa_pos.
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
  DATA: wa_item TYPE /sbxc/zckp_invi,
        p_idx_lin LIKE sy-index,
        l_fieldname TYPE c LENGTH 60.


** Preencher estrutura item
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  LOOP AT lt_sval WHERE value IS NOT INITIAL.
    CONDENSE lt_sval-value NO-GAPS.
    CHECK lt_sval-value ne '0.00'.
    LOOP AT t_linhas_sel.
      READ TABLE item INTO wa_item INDEX  t_linhas_sel-indice.
      CONCATENATE 'WA_ITEM-' lt_sval-fieldname INTO l_fieldname.
      ASSIGN (l_fieldname) TO <field>.
      <field> = lt_sval-value.
      IF t_linhas_sel-indice <> 0.
        MODIFY item INDEX t_linhas_sel-indice FROM wa_item.
      ENDIF.
    ENDLOOP.
  ENDLOOP.



  REFRESH linha.
  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.
*
  REFRESH t_linhas_sel.

ENDFUNCTION.
