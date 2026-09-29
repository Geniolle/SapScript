FUNCTION /SBXC/ZCKP_DETERMINA_PROCESSO1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(HEADERDATA) TYPE  /SBXC/ZCKP_INVH
*"     REFERENCE(TAX)
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"     REFERENCE(PROCESSO) TYPE  /SBXC/ZCKP_INVH-PROCESSO
*"     REFERENCE(REGRA) TYPE  /SBXC/ZCKP_REGRA
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI
*"      RETURN STRUCTURE  BAPIRET2
*"----------------------------------------------------------------------


  MOVE-CORRESPONDING headerdata TO /sbxc/zckp_invh.
  DATA: item LIKE  /sbxc/zckp_invi-invoice_doc_item,
        org_cmp LIKE ekko-ekorg.

  DATA: ls_tab13 TYPE /sbxc/zckp_tab13.
  DATA: lt_tab12 TYPE TABLE OF /sbxc/zckp_tab12 WITH HEADER LINE.
  SELECT SINGLE * FROM /sbxc/zckp_tab13
    INTO ls_tab13
    WHERE bukrs = headerdata-comp_code.
  IF sy-subrc NE 0.
    SELECT SINGLE * FROM /sbxc/zckp_tab13
       INTO ls_tab13.
  ENDIF.
  SELECT * FROM /sbxc/zckp_tab12
      INTO TABLE lt_tab12.

  DATA: lv_bukrs TYPE bukrs, lv_bsart TYPE bsart, lv_busab TYPE busab, lv_knttp TYPE knttp,
        lv_poref TYPE /sbxc/zckp_poref.

  IF ls_tab13-xbukrs EQ 'X'.
    CONCATENATE regra 'BUKRS/' INTO regra.
    lv_bukrs = headerdata-comp_code.
  ENDIF.

  IF ls_tab13-xbusab EQ 'X'.
    CONCATENATE regra 'BUSAB/' INTO regra.
    SELECT SINGLE busab INTO lv_busab
      FROM lfb1
      WHERE lifnr = headerdata-vendor
      AND bukrs = headerdata-comp_code.
  ENDIF.

  IF ls_tab13-xbsart EQ 'X'.
    CONCATENATE regra 'BSART/' INTO regra.
  ENDIF.

  IF ls_tab13-xknttp EQ 'X'.
    CONCATENATE regra 'KNTTP/' INTO regra.
  ENDIF.
  IF ls_tab13-xporef EQ 'X'.
    CONCATENATE regra 'POREF/' INTO regra.
  ENDIF.

  IF ls_tab13-xporef EQ 'X'.
     READ TABLE itemdata INDEX 1.
    IF pedido eq space and ( itemdata-po_number is not initial or itemdata-ref_1 is not initial ) .
       processo = 'OUTROS'.
    elseif pedido is not initial.
      lv_poref = 'X'.
    ENDIF.
  endif.

*  check processo ne 'OUTROS'.
if processo ne 'OUTROS'.

  LOOP AT itemdata.

    IF ls_tab13-xbsart EQ 'X'.
      SELECT SINGLE bsart INTO lv_bsart FROM ekko
        WHERE ebeln = itemdata-po_number.
    ENDIF.

    IF ls_tab13-xknttp EQ 'X'.
      SELECT SINGLE knttp INTO lv_knttp FROM ekpo
        WHERE ebeln = itemdata-po_number
        AND ebelp = itemdata-po_item.
    ENDIF.

  ENDLOOP.


  LOOP AT lt_tab12 WHERE bukrs = lv_bukrs AND
      bsart = lv_bsart AND
      busab = lv_busab AND
      knttp = lv_knttp AND
      poref = lv_poref.
  ENDLOOP.
  IF sy-subrc EQ 0.
    processo = lt_tab12-processo.
  ELSE.
    processo = 'OUTROS'.
  ENDIF.
endif.
ENDFUNCTION.
