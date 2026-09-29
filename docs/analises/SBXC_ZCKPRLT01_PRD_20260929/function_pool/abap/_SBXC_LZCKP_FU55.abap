FUNCTION /sbxc/zckp_bapi_invoice1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) TYPE  SYST-UCOMM
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      COR OPTIONAL
*"      LINHA
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------

  DATA: it_cab TYPE TABLE OF  /sbxc/zckp_invh WITH HEADER LINE.
  DATA: it_lin TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: it_tax TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE,
        l_ind  TYPE sy-tabix.
**********************************************************************
  TABLES: a003, konp, lfa1.
**BAPI FI
  DATA: i_bapiache03 LIKE bapiache03,
        i_bapiacap03 LIKE bapiacap03 OCCURS 0 WITH HEADER LINE,
        i_bapiacgl03 LIKE bapiacgl03 OCCURS 0 WITH HEADER LINE,
        i_bapiactx01 LIKE bapiactx01 OCCURS 0 WITH HEADER LINE,
        i_bapiaccr01 LIKE bapiaccr01 OCCURS 0 WITH HEADER LINE,
        return       LIKE bapiret2   OCCURS 0 WITH HEADER LINE,
        t_extension  LIKE bapiextc   OCCURS 0 WITH HEADER LINE,
        belnr        LIKE bapiache03-obj_key.
  DATA: cont TYPE i.
**********************************************************************

  DATA: fat_estornada_mm(1).

  REFRESH: t_fimsg.
  CLEAR: t_fimsg.
  refresh = 'X'.

  MOVE-CORRESPONDING cab TO it_cab.
  APPEND it_cab.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO it_lin.
    APPEND it_lin.
  ENDLOOP.


  READ TABLE it_cab INDEX 1.

  SELECT SINGLE stceg FROM lfa1 INTO lfa1-stceg WHERE lifnr = it_cab-vendor.
  CONDENSE   lfa1-stceg.

**********************************************************************
  IF ctrl-status1 = '4' OR ctrl-status1 = '3'.
    MESSAGE s012(/sbxc/zckp_cockpit).
  ELSE.
    REFRESH i_bapiacap03.
    REFRESH i_bapiaccr01.
    REFRESH i_bapiacgl03.
    REFRESH i_bapiactx01.
    REFRESH t_extension.
    CLEAR: i_bapiache03, i_bapiacap03, i_bapiaccr01, i_bapiacgl03,
           i_bapiactx01, t_extension.

    i_bapiache03-username     = sy-uname.
    i_bapiache03-comp_code    = it_cab-comp_code.
    i_bapiache03-fisc_year    = it_cab-pstng_date(4).
    i_bapiache03-doc_date     = it_cab-doc_date.
    i_bapiache03-pstng_date   = it_cab-pstng_date.
    i_bapiache03-doc_type     = it_cab-doc_type.
    i_bapiache03-ref_doc_no   = it_cab-ref_doc_no.
    i_bapiache03-header_txt   = it_cab-header_txt.

    cont = 1.
    i_bapiacap03-itemno_acc   = cont.
    i_bapiacap03-vendor_no    = it_cab-vendor.

    IF it_cab-conta_forn IS INITIAL.
      SELECT SINGLE akont INTO it_cab-conta_forn
        FROM lfb1 WHERE lifnr = it_cab-vendor
        AND bukrs = it_cab-comp_code.
    ENDIF.

    i_bapiacap03-gl_account   = it_cab-conta_forn.
    i_bapiacap03-alloc_nmbr   = it_cab-alloc_nmbr.
*    i_bapiacap03-pmnttrms     = it_cab-pmnttrms.
    i_bapiacap03-pmnt_block   = it_cab-pmnt_block.
*    i_bapiacap03-w_tax_code   = it_cab-w_tax_code.
*    i_bapiacap03-w_tax_code   = it_cab-cod_irf.
    i_bapiacap03-pmtmthsupl   = it_cab-pmtmthsupl.
*    i_bapiacap03-ref_key_1    = it_cab-ref_key_1.
*    i_bapiacap03-ref_key_2    = it_cab-ref_key_2.
*    i_bapiacap03-ref_key_3    = it_cab-ref_key_3.
*    i_bapiacap03-pymt_meth    = it_cab-pymt_meth.
    i_bapiacap03-item_text    = it_cab-item_text.
    APPEND i_bapiacap03.

    i_bapiaccr01-itemno_acc   = cont.
    i_bapiaccr01-currency     = it_cab-currency.
**Valor com sinais contrários nas notas de crédito
    IF it_cab-doctypesaphety NE 'NC' AND it_cab-doctypesaphety NE 'CREDITNOTE'
      and it_cab-doctypesaphety ne '381'. "ODC - 12_03_2021.
      i_bapiaccr01-amt_doccur = it_cab-gross_amount * -1.
    ELSE.
      i_bapiaccr01-amt_doccur = it_cab-gross_amount.
    ENDIF.
    i_bapiaccr01-exch_rate    = it_cab-exch_rate.
    i_bapiaccr01-amt_base     = it_cab-gross_amount.
    APPEND i_bapiaccr01.


    LOOP AT it_lin.

      cont = cont + 1.
      i_bapiacgl03-itemno_acc = cont.
      IF it_lin-gl_account CO '0123456789 ' AND NOT it_lin-gl_account IS INITIAL.
        UNPACK it_lin-gl_account TO it_lin-gl_account.
      ENDIF.
      i_bapiacgl03-gl_account = it_lin-gl_account.
      IF it_lin-anln1 CO '0123456789 ' AND NOT it_lin-anln1 IS INITIAL.
        UNPACK it_lin-anln1 TO it_lin-anln1.
      ENDIF.
      i_bapiacgl03-asset_no = it_lin-anln1.

      IF it_lin-anln2 CO '0123456789 ' AND NOT it_lin-anln2 IS INITIAL.
        UNPACK it_lin-anln2 TO it_lin-anln2.
      ENDIF.
      i_bapiacgl03-sub_number = it_lin-anln2.

      i_bapiacgl03-tax_code   = it_lin-tax_code_sap.
      i_bapiacgl03-item_text  = it_lin-item_text.
      i_bapiacgl03-bus_area   = it_lin-gsber.
* Apenas para testes - Centro Lucro
*      i_bapiacgl03-profit_ctr = '0000001900'.

      IF it_lin-costcenter CO '0123456789 ' AND NOT it_lin-costcenter IS INITIAL.
        UNPACK it_lin-costcenter TO it_lin-costcenter.
      ENDIF.
      i_bapiacgl03-costcenter = it_lin-costcenter.
      IF it_lin-orderid CO '0123456789 ' AND NOT it_lin-orderid IS INITIAL.
        UNPACK it_lin-orderid TO it_lin-orderid.
      ENDIF.
      i_bapiacgl03-orderid = it_lin-orderid.

      i_bapiacgl03-alloc_nmbr = it_lin-zuonr.

      IF it_lin-wbs_elem NE space.
        CALL FUNCTION 'CONVERSION_EXIT_ABPSP_OUTPUT'
          EXPORTING
            input  = it_lin-wbs_elem
          IMPORTING
            output = i_bapiacgl03-wbs_element.
      ENDIF.

      i_bapiacgl03-po_pr_qnt = it_lin-quantity.
      i_bapiacgl03-po_pr_uom = it_lin-po_unit.
**        i_bapiacgl03-ref_key_1 = it_lin-ref_key_1.
**        i_bapiacgl03-ref_key_2 = it_lin-ref_key_2.
**        i_bapiacgl03-ref_key_3 = it_lin-ref_key_3.
      APPEND i_bapiacgl03.

      i_bapiaccr01-itemno_acc   = cont.
      i_bapiaccr01-currency     = it_cab-currency.
*  *Valor com sinais contrários nas notas de crédito
      IF it_cab-doctypesaphety NE 'NC' AND it_cab-doctypesaphety NE 'CREDITNOTE'
        and it_cab-doctypesaphety ne '381'. "ODC - 12_03_2021.
        i_bapiaccr01-amt_doccur   = it_lin-item_amount.
      ELSE.
        i_bapiaccr01-amt_doccur   = it_lin-item_amount * -1.
      ENDIF.
      i_bapiaccr01-exch_rate    = it_cab-exch_rate.
      i_bapiaccr01-amt_base     = it_lin-item_amount.
      APPEND i_bapiaccr01.

    ENDLOOP.

*Agrupar por cód iva
    LOOP AT it_lin.
      READ TABLE it_tax WITH KEY tax_code_sap = it_lin-tax_code_sap.
      l_ind = sy-tabix.
      IF sy-subrc NE 0.
        APPEND it_lin TO it_tax.
      ELSE.
        it_tax-tax_amount = it_tax-tax_amount + it_lin-tax_amount.
        it_tax-item_amount =  it_tax-item_amount + it_lin-item_amount.
        MODIFY it_tax INDEX l_ind.
      ENDIF.
    ENDLOOP.

    DATA: it_mwdat TYPE TABLE OF rtax1u15 WITH HEADER LINE.
    DATA: xbukrs TYPE bkpf-bukrs,
          xmwskz TYPE bseg-mwskz,
          xwaers TYPE bkpf-waers,
          xwrbtr TYPE bseg-wrbtr.

    LOOP AT it_tax.
      "ODC - 06_06_2018
      xbukrs = it_cab-comp_code.
      xmwskz = it_tax-tax_code_sap.
      xwaers = it_cab-currency.
      xwrbtr = it_tax-item_amount.
      IF xmwskz IS INITIAL.
        MESSAGE s063(/SBXC/ZCKP_COCKPIT).
      ELSE.
        CALL FUNCTION 'CALCULATE_TAX_FROM_NET_AMOUNT'
          EXPORTING
            i_bukrs = xbukrs
            i_mwskz = xmwskz
            i_waers = xwaers
            i_wrbtr = xwrbtr
          TABLES
            t_mwdat = it_mwdat.

        LOOP AT it_mwdat.
          cont = cont + 1.
          it_tax-tax_amount = it_mwdat-wmwst.
          i_bapiactx01-itemno_acc = cont.
          i_bapiactx01-tax_code = it_tax-tax_code_sap.
          i_bapiactx01-cond_key = it_mwdat-kschl.
          i_bapiactx01-tax_rate = it_mwdat-msatz.
          i_bapiactx01-acct_key = it_mwdat-ktosl.
          APPEND i_bapiactx01.

          i_bapiaccr01-itemno_acc   = cont.
          i_bapiaccr01-currency     = it_cab-currency.
          IF it_cab-doctypesaphety NE 'NC' AND it_cab-doctypesaphety NE 'CREDITNOTE'
            and it_cab-doctypesaphety ne '381'. "ODC - 12_03_2021..
            i_bapiaccr01-amt_doccur   = it_tax-tax_amount.
          ELSE.
            i_bapiaccr01-amt_doccur   = it_tax-tax_amount * -1.
          ENDIF.
          i_bapiaccr01-exch_rate    = it_cab-exch_rate.
          i_bapiaccr01-amt_base     = it_mwdat-kawrt.
          APPEND i_bapiaccr01.
        ENDLOOP.
        "Fim ODC 06_06_2018
*      SELECT * FROM a003 WHERE kappl = 'TX'
*                           AND aland = 'PT'
*                           AND mwskz = it_tax-tax_code_sap.
*
*        SELECT SINGLE * FROM konp WHERE knumh = a003-knumh
*                                    AND mwsk1 = a003-mwskz
*                                    AND kschl = a003-kschl.
*
*        IF konp-kbetr NE 0.
*          cont = cont + 1.
*          IF NOT it_tax-tax_amount IS INITIAL.
*            i_bapiactx01-itemno_acc = cont.
*            i_bapiactx01-tax_code = it_tax-tax_code_sap.
*            i_bapiactx01-cond_key = konp-kschl.
*            APPEND i_bapiactx01.
*          ELSE.
*            it_tax-tax_amount = it_tax-item_amount * ( konp-kbetr / 1000 ).
*            i_bapiactx01-itemno_acc = cont.
*            i_bapiactx01-tax_code = it_tax-tax_code_sap.
*            i_bapiactx01-cond_key = konp-kschl.
*            APPEND i_bapiactx01.
*          ENDIF.
*          i_bapiaccr01-itemno_acc   = cont.
*          i_bapiaccr01-currency     = it_cab-currency.
*          IF it_cab-doctypesaphety NE 'NC' AND it_cab-doctypesaphety NE 'CREDITNOTE'.
*            i_bapiaccr01-amt_doccur   = it_tax-tax_amount.
*          ELSE.
*            i_bapiaccr01-amt_doccur   = it_tax-tax_amount * -1.
*          ENDIF.
*          i_bapiaccr01-exch_rate    = it_cab-exch_rate.
*          i_bapiaccr01-amt_base     = it_tax-item_amount.
*          APPEND i_bapiaccr01.
*        ENDIF.
*
*      ENDSELECT.
      ENDIF.
    ENDLOOP.


**    if itab_cab-cat_doc eq 'NC' and itab_cab-rebzg ne ' '.
**      t_extension-field1 = 'REBZG'.
**      t_extension-field2 = itab_cab-rebzg.
**      append t_extension.
**    endif.


* Nota 0000458392 - Withholding tax data
*  data: mont_irf1(15),
*        val_tot1(15).
*  if it_cab-cat_irf ne ' ' and it_cab-cod_irf ne ' '  and
*     it_cab-mont_irf ne 0.
*    t_extension-field1+0(6) = '000001'. "ACCWT-WT_KEY
*    t_extension-field1+6(2) = it_cab-cat_irf.     "ACCWT-WITHT
*    t_extension-field1+8(2) = it_cab-cod_irf.     "ACCWT-WT_WITHCD
*    val_tot1 = it_cab-gross_amount.
*    mont_irf1 = it_cab-mont_irf.
*    t_extension-field2 = val_tot1.                  "WT_QSSHB
*    t_extension-field3 = mont_irf1.                 "ACCWT-WT_QBUIHB
*
*    append t_extension.
*  endif.

*    IF it_cab-doctypesaphety EQ 'NC' OR it_cab-doctypesaphety EQ 'CREDITNOTE'
*                      OR it_cab-doctypesaphety EQ 'CN' OR it_cab-doctypesaphety EQ '4'.
*      "Verificar Nota de crédito dupla
*      DATA flag_bkpf_n(1).
*      DATA flag_aux_n(1).
*      CLEAR flag_aux_n.
*      TRANSLATE wa_header-ref_doc_no  TO UPPER CASE.
*      SELECT * FROM bsip WHERE
*                    bukrs = it_cab-comp_code AND
*                    lifnr = it_cab-vendor AND
*                    waers = it_cab-currency AND
*                    bldat = it_cab-doc_date AND
*                    xblnr = it_cab-ref_doc_no AND
*                    wrbtr = it_cab-gross_amount AND
*                    gjahr =  it_cab-pstng_date(4) AND
*                 shkzg = 'S'.
*        CLEAR flag_bkpf_n.
*        CLEAR bkpf.
*        flag_aux_n =  abap_true.
*        SELECT SINGLE * FROM bkpf WHERE belnr = bsip-belnr AND
*                                        bukrs = bsip-bukrs AND
*                                        gjahr = bsip-gjahr AND
*                                        xreversal EQ ' '.
*        IF sy-subrc = 0.
*          flag_bkpf_n = 'X'.
*        ELSE.
*          SELECT SINGLE  * FROM rbkp WHERE
*                bukrs = bkpf-bukrs AND
*                belnr = bkpf-awkey(10) AND
*                gjahr = bkpf-awkey+10(4) AND
*                stblg = ''.
*          IF sy-subrc NE 0.
*            fat_estornada_mm = 'X'.
*          ENDIF.
*        ENDIF.
*
*      ENDSELECT.
*      IF  ( fat_estornada_mm = ' ' AND flag_bkpf_n = 'X' ) OR ( fat_estornada_mm = ' ' AND flag_bkpf_n = ' '  AND flag_aux_n EQ abap_true )
*        OR ( fat_estornada_mm = 'X' AND flag_bkpf_n = 'X'  AND flag_aux_n EQ abap_true ).  "nao foi estornada nem em MM nem FI
*        MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr DISPLAY LIKE 'E'.
*        return.
*      ENDIF.
*      "Fim Verificar Nota de crédito dupla
*    ELSE.
** Valida Duplicados
*      DATA flag_bkpf(1).
*      DATA flag_aux(1).
*      CLEAR flag_aux.
*      SELECT * FROM bsip WHERE
*            bukrs = it_cab-comp_code AND
*            lifnr = it_cab-vendor AND
*            waers = it_cab-currency AND
*            bldat = it_cab-doc_date AND
*            xblnr = it_cab-ref_doc_no AND
*            wrbtr = it_cab-gross_amount AND
*            gjahr =  it_cab-pstng_date(4) AND
*            shkzg = 'H'.
*        CLEAR flag_bkpf.
*        CLEAR bkpf.
*        flag_aux =  abap_true.
*        SELECT SINGLE * FROM bkpf WHERE belnr = bsip-belnr AND
*                                        bukrs = bsip-bukrs AND
*                                        gjahr = bsip-gjahr AND
*                                        xreversal EQ ' '.
*        IF sy-subrc = 0.
*          flag_bkpf = 'X'.
*        ELSE.
*          SELECT SINGLE  * FROM rbkp WHERE
*                    bukrs = bkpf-bukrs AND
*                    belnr = bkpf-awkey(10) AND
*                    gjahr = bkpf-awkey+10(4) AND
*                    stblg = ''.
*          IF sy-subrc NE 0.
*            fat_estornada_mm = 'X'.
*          ENDIF.
*        ENDIF.
*      ENDSELECT.
*      IF  ( fat_estornada_mm = ' ' AND flag_bkpf = 'X' ) OR ( fat_estornada_mm = ' ' AND flag_bkpf = ' '  AND flag_aux EQ abap_true )
*        OR ( fat_estornada_mm = 'X' AND flag_bkpf = 'X'  AND flag_aux EQ abap_true ).  "nao foi estornada nem em MM nem FI
*        MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr DISPLAY LIKE 'E'.
*        RETURN.
*      ENDIF.
*    ENDIF.

      "Valida duplicados
      DATA: lv_return TYPE sy-subrc.
      CALL FUNCTION '/SBXC/ZCKP_VALIDA_DUPLICADO'
        EXPORTING
          doctypesaphety = it_cab-doctypesaphety
        IMPORTING
          return         = lv_return
          CHANGING
            cab            = it_cab
            ctrl           = ctrl.
      IF lv_return NE 4.

    CHECK bkpf-belnr IS INITIAL.

    CALL FUNCTION '/SBXC/ZCKP_INV_BAPI_FI'
      EXPORTING
        documentheader = i_bapiache03
      IMPORTING
        obj_key        = belnr
      TABLES
        accountpayable = i_bapiacap03
        accountgl      = i_bapiacgl03
        accounttax     = i_bapiactx01
        currencyamount = i_bapiaccr01
        extension1     = t_extension
        return         = return.



    IF belnr(10) CO ' 0123456789'.
*    move-corresponding t_fimsg to msg_ckp.
*      MESSAGE s018(/SBXC/ZCKP_COCKPIT) WITH belnr(10) it_cab-comp_code.
*      REFRESH return.
      UPDATE /sbxc/zckp_invh SET
             comp_code = it_cab-comp_code
             doc_fi = belnr(10)
             doc_estorno = ' '
             ano_lanc = it_cab-pstng_date(4)
             data_criacao = sy-datum
             enviado_saphety = ' '
       WHERE processo = it_cab-processo AND
             ano = it_cab-ano AND
             seqno = it_cab-seqno.

      UPDATE /sbxc/zckp_ctrl SET
      status1 = '4'
      WHERE processo = it_cab-processo AND
      ano = it_cab-ano AND
      seqno = it_cab-seqno.

      COMMIT WORK AND WAIT.

      it_cab-doc_fi = belnr(10).
      it_cab-ano_lanc = it_cab-pstng_date(4).
      it_cab-doc_estorno = ' '.
*    it_cab-status1 = '4'.
      MODIFY it_cab INDEX 1.
* Atualiza itens na tabela de cockpit
      DELETE FROM /sbxc/zckp_invi WHERE processo = it_cab-processo AND
             ano = it_cab-ano AND
             seqno = it_cab-seqno.
      MOVE-CORRESPONDING it_cab TO /sbxc/zckp_invi.
      SELECT buzei wrbtr gsber kostl hkont shkzg mwskz anln1 anln2 zuonr sgtxt FROM bseg
        INTO (/sbxc/zckp_invi-invoice_doc_item, /sbxc/zckp_invi-item_amount,
              /sbxc/zckp_invi-gsber, /sbxc/zckp_invi-costcenter,
              /sbxc/zckp_invi-gl_account, /sbxc/zckp_invi-db_cr_ind,   /sbxc/zckp_invi-tax_code_sap, /sbxc/zckp_invi-anln1, /sbxc/zckp_invi-anln2,
              /sbxc/zckp_invi-zuonr, /sbxc/zckp_invi-item_text)
        WHERE bukrs = it_cab-comp_code AND
              belnr = it_cab-doc_fi AND
              gjahr = it_cab-ano_lanc AND
              koart NE 'K' AND
              buzid NE 'T'
        ORDER BY buzei ASCENDING.
        REFRESH it_mwdat.
        CALL FUNCTION 'CALCULATE_TAX_FROM_NET_AMOUNT'
          EXPORTING
            i_bukrs           = it_cab-comp_code
            i_mwskz           = /sbxc/zckp_invi-tax_code_sap
            i_waers           = it_cab-currency
            i_wrbtr           = /sbxc/zckp_invi-item_amount
          TABLES
            t_mwdat           = it_mwdat
          EXCEPTIONS
            bukrs_not_found   = 1
            country_not_found = 2
            mwskz_not_defined = 3
            mwskz_not_valid   = 4
            ktosl_not_found   = 5
            kalsm_not_found   = 6
            parameter_error   = 7
            knumh_not_found   = 8
            kschl_not_found   = 9
            unknown_error     = 10
            account_not_found = 11
            txjcd_not_valid   = 12
            OTHERS            = 13.

        LOOP AT it_mwdat INTO wa_mwdat.
          /sbxc/zckp_invi-tax_imposto_sap = wa_mwdat-msatz.
          /sbxc/zckp_invi-tax_amount =  wa_mwdat-wmwst.
        ENDLOOP.
        INSERT  /sbxc/zckp_invi.
        CLEAR /sbxc/zckp_invi.
        MOVE-CORRESPONDING it_cab TO /sbxc/zckp_invi.
      ENDSELECT.

      COMMIT WORK.
* anexa doc. panagon ao doc de fi
*    perform anexos using it_cab-id_panagon it_cab-comp_code belnr(10) it_cab-pstng_date(4).

      IF it_cab-url IS NOT INITIAL.

        DATA: lv_belnr LIKE bkpf-belnr.

        lv_belnr = belnr(10).

        CALL FUNCTION '/SBXC/ZCKP_ANEXA_URL'
          EXPORTING
            bukrs = it_cab-comp_code
            belnr = lv_belnr
            gjahr = it_cab-pstng_date(4)
            url   = it_cab-url.
      ENDIF.

    ELSE.
      UPDATE /sbxc/zckp_ctrl SET
      status1 = '2'
      WHERE processo = it_cab-processo AND
      ano = it_cab-ano AND
      seqno = it_cab-seqno.

      COMMIT WORK AND WAIT.

      it_cab-doc_fi = ' '.
      it_cab-ano_lanc = ' '.
      it_cab-doc_estorno = ' '.
*    it_cab-status1 = '2'.
      MODIFY it_cab INDEX 1.

      MESSAGE id '/SBXC/ZCKP_COCKPIT' TYPE 'S' NUMBER 083 DISPLAY LIKE 'E'.
    ENDIF.

    LOOP AT return.
      CLEAR t_fimsg.
      t_fimsg-msgid = return-id.
      t_fimsg-msgty = return-type.
      t_fimsg-msgno = return-number.
      t_fimsg-msgv1 = return-message_v1.
      t_fimsg-msgv2 = return-message_v2.
      t_fimsg-msgv3 = return-message_v3.
      t_fimsg-msgv4 = return-message_v4.
      APPEND t_fimsg.
    ENDLOOP.

* apaga tabela de log
    PERFORM log_delete2 TABLES t_fimsg USING it_cab-processo it_cab-ano it_cab-seqno .
* Guarda msg na tabela de log
    PERFORM log2 TABLES t_fimsg USING it_cab-processo it_cab-ano it_cab-seqno .


    CLEAR: cab.

*  loop at it_cab.
    MOVE-CORRESPONDING it_cab TO cab.
*    append cab.
*  endloop.

  ENDIF.
  endif.
ENDFUNCTION.
