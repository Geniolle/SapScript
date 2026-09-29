FUNCTION /sbxc/zckp_ws_invoice2cockpit.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(MESSAGE) TYPE  /SBXC/CKP_ST_023
*"----------------------------------------------------------------------


  DATA: headerdata TYPE /sbxc/zckp_invh,
        itemdata   TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE,
        return     TYPE TABLE OF bapiret2 WITH HEADER LINE.

  DATA: msg       TYPE /sbxc/ckp_st_023,
        sender    TYPE /sbxc/ckp_st_026,
        receiver  TYPE /sbxc/ckp_st_026,
        invoice   TYPE /sbxc/ckp_st_tab_010 WITH HEADER LINE,
        docstat   TYPE /sbxc/ckp_st_tab_008 WITH HEADER LINE,
        seller    TYPE /sbxc/ckp_st_022,
        buyer     TYPE /sbxc/ckp_st_022,
        sellerid  TYPE /sbxc/ckp_st_tab_009 WITH HEADER LINE,
        buyerid   TYPE /sbxc/ckp_st_tab_009 WITH HEADER LINE,
        lineitem  TYPE /sbxc/ckp_st_tab_011 WITH HEADER LINE,
        lineref   TYPE /sbxc/ckp_st_tab_013 WITH HEADER LINE,
        quantity  TYPE /sbxc/ckp_st_025.


  msg = message.

  invoice[] = msg-invoice[].

  LOOP AT invoice.
    REFRESH: lineitem, buyerid, sellerid.
    CLEAR: headerdata.

    buyer = invoice-buyer.
    buyerid[] = buyer-id.
    IF buyerid-typecode EQ 'VAT'.
      SELECT SINGLE bukrs  "#EC CI_NOORDER
        INTO headerdata-comp_code
        FROM t001 WHERE stceg = sellerid-id.
    ENDIF.

    headerdata-doc_type = 'FT'. "Teste
    headerdata-ref_doc_no = invoice-documentnumber.

    seller = invoice-seller.
    sellerid[] = seller-id.
    IF sellerid-typecode EQ 'VAT'.
      headerdata-stcd1 = sellerid-id.
      SELECT SINGLE lifnr name1  "#EC CI_NOORDER
        INTO (headerdata-vendor,
             headerdata-name1 )
        FROM lfa1 WHERE stceg = sellerid-id.
    ENDIF.

    headerdata-currency = invoice-currencycode.
    headerdata-doc_date = invoice-documentdate.

    if invoice-totgrossamount co '1234567890.'.
    headerdata-gross_amount = invoice-totgrossamount.
    else.

    endif.

    lineitem[] = invoice-lineitem.

    LOOP AT lineitem.
      REFRESH: lineref.
      lineref[] = lineitem-reference.

      itemdata-invoice_doc_item = lineitem-linenumber.

      LOOP AT lineref.
        IF lineref-type EQ 'ORDER'.
          itemdata-po_number = lineref-refdocid.
          itemdata-po_item = lineref-lineid.
        ENDIF.
      ENDLOOP.

      itemdata-item_text = lineitem-description.


      quantity = lineitem-quantity.


*      if itemdata-item_amount co '1234567890.'.
*      itemdata-item_amount = lineitem-netamount.
*      endif.
*
*      if itemdata-quantity co '1234567890.'.
*      itemdata-quantity = quantity-value.
*      endif.
*
*      if itemdata-po_unit co '1234567890.'.
*      itemdata-po_unit = quantity-unit.
*      endif.
*
*      if itemdata-tax_amount co '1234567890.'.
*      itemdata-tax_amount = lineitem-vatamount.
*      endif.

    ENDLOOP.



    CALL FUNCTION '/SBXC/ZCKP_MM_CRIA_FACTURAS'
      EXPORTING
        headerdata = headerdata
      TABLES
        itemdata   = itemdata
        return     = return.

  ENDLOOP. "invoice

ENDFUNCTION.
