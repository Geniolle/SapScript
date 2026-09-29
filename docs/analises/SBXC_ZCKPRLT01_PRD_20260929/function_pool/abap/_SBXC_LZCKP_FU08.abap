FUNCTION /sbxc/zckp_fbr2.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) TYPE  SYST-UCOMM
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      COR
*"      LINHA
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------
  DATA: it_cab TYPE TABLE OF  /sbxc/zckp_invh WITH HEADER LINE.
  DATA: it_lin TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: l_bukrs LIKE bkpf-bukrs,
        l_belnr LIKE bkpf-belnr,
        l_gjahr LIKE bkpf-gjahr,
        objectkey LIKE swotobjid-objkey,
        dia(10),
        lv_bldat(8),
        lv_budat(8),
        lv_amount LIKE bbseg-wrbtr,
        lv_wrbtr LIKE bbseg-wrbtr,
        ls_nc.

    DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: ls_bim TYPE zsckp_to_bim,
        lv_errortext   TYPE string.
  DATA: lv_return TYPE sy-subrc.
  CONSTANTS: c_fi TYPE zsckp_to_bim-mm_fi VALUE 'FI'.


  DATA: lv_gjahr TYPE gjahr.


*  CHECK e_ucomm EQ 'FBR2' AND NOT cab[] IS INITIAL.
  CHECK e_ucomm EQ 'FBR2' AND NOT cab IS INITIAL.
  DATA doc_fi LIKE /sbxc/zckp_invh-doc_fi.
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO it_cab.
*    APPEND it_cab.
*  ENDLOOP.
  IF it_cab-em_tratamento = 'X' AND it_cab-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH it_cab-user_tratamento.
  ELSE.
    LOOP AT linha.
      MOVE-CORRESPONDING linha TO it_lin.
      APPEND it_lin.
    ENDLOOP.

* Obter documento modelo
    CALL FUNCTION 'ARCHIV_POPUP_OBJECT_KEY'
      EXPORTING
        objtype         = 'BKPF'
      IMPORTING
        objkey          = objectkey
      EXCEPTIONS
        error_parameter = 1
        user_cancel     = 2
        OTHERS          = 3.

    CHECK sy-subrc EQ 0.

*  LOOP AT it_cab.

    CLEAR status_ok.
*  PERFORM valida_status USING '0125'
*                              it_cab-processo
*                              it_cab-ano
*                              it_cab-seqno
*                              ctrl-status1
*                    CHANGING status_ok.
*
*  CHECK status_ok = 'X'.



    REFRESH: bdcdata, messtab.
    CLEAR bdcdata.

    WRITE it_cab-doc_date TO dia .
    TRANSLATE dia USING '. / - ' .
    CONDENSE dia NO-GAPS.
    lv_bldat = dia.

    WRITE it_cab-pstng_date TO dia .
    TRANSLATE dia USING '. / - ' .
    CONDENSE dia NO-GAPS.
    lv_budat = dia.

    WRITE it_cab-gross_amount TO  lv_amount DECIMALS 2 CURRENCY it_cab-currency.
    WRITE lv_amount TO  lv_wrbtr CURRENCY it_cab-currency .

    PERFORM inserir_dados USING:
      'SAPMF05A' '0104' 'X',
      'BDC_OKCODE' '/00' '',
      'BKPF-BELNR' objectkey+4(10) '',
      'BKPF-BUKRS' objectkey(4) '',
      'BKPF-GJAHR' objectkey+14(4) '',
      'RF05A-CPBET'  'X' '',
      'SAPMF05A' '0100' 'X',
      'BDC_OKCODE' '/00' '',
      'BKPF-BLDAT' lv_bldat '',
      'BKPF-BUDAT' lv_budat '',
      'BKPF-XBLNR' it_cab-ref_doc_no '',
      'RF05A-NEWKO' it_cab-vendor '',
      'SAPMF05A' '0302' 'X',
      'BDC_OKCODE' '/00' '',
      'BSEG-WRBTR' lv_wrbtr ''. "#EC CI_FLDEXT_OK[2610650]


    REFRESH messtab.
    CALL TRANSACTION 'FBR2' USING bdcdata MODE 'A' UPDATE 'S'
                                  MESSAGES INTO messtab.

    LOOP AT messtab WHERE msgid = 'F5' AND ( msgnr  = '312' OR
      msgnr	= '323' ).
    ENDLOOP.
    doc_fi = messtab-msgv1.
    UNPACK doc_fi TO doc_fi.
    IF sy-subrc EQ 0.
      REFRESH return.

        CLEAR lv_gjahr.
        GET PARAMETER ID 'GJR' FIELD lv_gjahr.

      UPDATE /sbxc/zckp_invh SET
             comp_code = messtab-msgv2
             doc_fi = doc_fi " messtab-msgv1
             doc_estorno = ' '
*             ano_lanc = it_cab-pstng_date(4)

               ano_lanc = lv_gjahr            "it_cab-pstng_date(4)

*           enviado_saphety = ' '
       WHERE processo = it_cab-processo AND
             ano = it_cab-ano AND
             seqno = it_cab-seqno.

      SELECT SINGLE      cpudt  cputm usnam FROM  bkpf
                                       INTO
                                       (bkpf-cpudt, bkpf-cputm, bkpf-usnam)
                                        WHERE bukrs =  messtab-msgv2 AND
                                                       belnr = doc_fi AND

                                                     gjahr = lv_gjahr.            "it_cab-pstng_date(4).

      UPDATE /sbxc/zckp_ctrl SET
      status1 = '4'
          data_chg_st1 = bkpf-cpudt
           hora_chg_st1 = bkpf-cputm
           user_chg_st1 = bkpf-usnam
      WHERE processo = it_cab-processo AND
      ano = it_cab-ano AND
      seqno = it_cab-seqno.

      COMMIT WORK AND WAIT.

*Associar doc
      it_cab-doc_fi = messtab-msgv1.
      UNPACK it_cab-doc_fi TO it_cab-doc_fi.

        it_cab-ano_lanc = lv_gjahr.            "it_cab-pstng_date(4).

      it_cab-doc_estorno = ' '.
*      it_cab-status1 = '4'.
*      modify it_cab.

**      Anexos
*      perform anexos using it_cab-id_panagon it_cab-comp_code it_cab-doc_fi it_cab-ano_lanc.
*
      IF it_cab-url IS NOT INITIAL.

        CALL FUNCTION '/SBXC/ZCKP_ANEXA_URL'
          EXPORTING
            bukrs = it_cab-comp_code
            belnr = it_cab-doc_fi
            gjahr = it_cab-ano_lanc
            url   = it_cab-url.
      ENDIF.

* Atualiza itens na tabela de cockpit
      DELETE FROM /sbxc/zckp_invi WHERE processo = it_cab-processo AND
             ano = it_cab-ano AND
             seqno = it_cab-seqno.
      MOVE-CORRESPONDING it_cab TO /sbxc/zckp_invi.
      SELECT buzei wrbtr gsber kostl hkont shkzg mwskz FROM bseg
        INTO (/sbxc/zckp_invi-invoice_doc_item, /sbxc/zckp_invi-item_amount,
              /sbxc/zckp_invi-gsber, /sbxc/zckp_invi-costcenter,
              /sbxc/zckp_invi-gl_account, /sbxc/zckp_invi-db_cr_ind,   /sbxc/zckp_invi-tax_code_sap)
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
      COMMIT WORK AND WAIT.

*  Se se tratar de uma NC, enviar email para o fornecedor
        CLEAR ls_nc.
        CALL FUNCTION '/SBXC/ZCKP_CHECK_NC'
          EXPORTING
            i_bukrs  = it_cab-comp_code
            i_doc_fi = it_cab-doc_fi
            i_gjahr  = it_cab-ano_lanc
          IMPORTING
            e_nc     = ls_nc.
        IF sy-subrc EQ 0 AND ls_nc EQ 'X'.
*               Enviar email para o fornecedor
          CALL FUNCTION '/SBXC/ZCKP_SENDMAIL_FORN'
            EXPORTING
              i_bukrs      = it_cab-comp_code
              i_ref_doc_no = it_cab-ref_doc_no
              i_gjahr      = it_cab-ano_lanc
              i_vendor     = it_cab-vendor
              i_doc_date   = it_cab-doc_date
              i_pstng_date = it_cab-pstng_date
              i_nc         = ls_nc.
        ENDIF.
*  Se se tratar de uma NC, enviar email para o fornecedor


          "Envia para workflow
            ls_bim-bukrs = it_cab-comp_code.
            ls_bim-belnr = it_cab-doc_fi.
            ls_bim-gjahr = it_cab-ano_lanc.
            ls_bim-mm_fi = c_fi.

            CREATE OBJECT lo_docinf.

            TRY.
                lo_docinf->send_to_bim_process( EXPORTING i_zsckp_to_bim = ls_bim
                                                IMPORTING e_wi_id = it_cab-wi_id
                                                          return_code = lv_return ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*           it_cab-erro = lv_errortext.
*

            IF lv_return eq '4'.
              it_cab-erro = text-010.
            ELSEIF lv_return eq '1'.
                it_cab-erro = text-011.
            ENDIF.

           UPDATE /sbxc/zckp_invh SET erro = it_cab-erro
           WHERE processo EQ it_cab-processo
            AND ano EQ it_cab-ano
            AND seqno EQ it_cab-seqno.
           commit WORK AND WAIT.
            ENDTRY.

            IF it_cab-wi_id IS NOT INITIAL .
              UPDATE /sbxc/zckp_invh SET wi_id = it_cab-wi_id
                WHERE processo EQ it_cab-processo
                AND ano EQ it_cab-ano
                AND seqno EQ it_cab-seqno.
              commit WORK AND WAIT.
            ENDIF.

            "Fim envio workflow
    ENDIF.

* apaga tabela de log
    PERFORM log_delete2 TABLES messtab
       USING it_cab-processo it_cab-ano it_cab-seqno .
* Guarda msg na tabela de log

    PERFORM log2 TABLES t_fimsg USING it_cab-processo it_cab-ano it_cab-seqno .

*  ENDLOOP.


*  REFRESH: cab.

*  LOOP AT it_cab.
    CLEAR cab.
    MOVE-CORRESPONDING it_cab TO cab.
*    APPEND cab.
*  ENDLOOP.
    COMMIT WORK AND WAIT.
    refresh = 'X'.
  ENDIF.
ENDFUNCTION.
