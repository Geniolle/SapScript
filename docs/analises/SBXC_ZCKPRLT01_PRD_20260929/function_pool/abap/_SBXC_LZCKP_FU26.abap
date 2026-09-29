FUNCTION /sbxc/zckp_mm_regista_fac_ckp.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) LIKE  SY-UCOMM
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------
* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.

  DATA: item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: simula_com LIKE sy-ucomm.

  DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: ls_bim       TYPE zsckp_to_bim,
        lv_errortext TYPE string.
  CONSTANTS: c_mm TYPE zsckp_to_bim-mm_fi VALUE 'MM'.

  DATA: it_cab TYPE TABLE OF  /sbxc/zckp_invh WITH HEADER LINE.

  DATA: lv_return TYPE sy-subrc.

  REFRESH msg_ckp.
  MOVE-CORRESPONDING cab TO it_cab.
  APPEND it_cab.
  READ TABLE it_cab INDEX 1.

*  DATA: ls_bsik TYPE bsik.
*  SELECT SINGLE * INTO ls_bsik
*    FROM bsik
*    WHERE lifnr EQ it_cab-vendor
*    AND bukrs EQ it_cab-comp_code
*    AND ( blart EQ 'TR' OR blart EQ 'BD' OR blart EQ 'BP' ).
*  IF sy-subrc EQ 0.
*    "Fornecedor & na empresa & com adiantamentos em conta corrente.
*    MESSAGE i057(/sbxc/zckp_cockpit) WITH it_cab-vendor it_cab-comp_code DISPLAY LIKE 'W'.
*  ENDIF.

  "Validar se a data de documento é posterior à data de lançamento
  IF it_cab-doc_date > it_cab-pstng_date.
    MESSAGE s064(/sbxc/zckp_cockpit) DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.


* CCF 21.01.2022
* Se a condição de pagamento tem o campo “Dia fixo” preenchido (T052 - ZFAEL), o sistema faz EOMONTH da “data base”.
  DATA: zfael    LIKE t052-zfael,
        lv_ldate TYPE sy-datum.

  SELECT SINGLE zfael INTO zfael FROM t052 WHERE zterm =  it_cab-zterm.
  IF zfael = '31'.
    CALL FUNCTION 'LAST_DAY_OF_MONTHS'
      EXPORTING
        day_in            = it_cab-baseline_date
      IMPORTING
        last_day_of_month = lv_ldate.
  ELSE. "MBOVO+ Se não for 31, a database é a do cockpit
    lv_ldate = it_cab-baseline_date.
  ENDIF.
  "FIM CCF
*   IF ctrl-status1 eq '9'."não permite avançar quando a fatura já está pré-editada.
*      MESSAGE s067(/sbxc/zckp_cockpit) DISPLAY LIKE 'E'.
*      exit.
*   endif.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    IF ( item-item_amount <> '0.00'  AND item-item_amount IS NOT INITIAL )
      OR ( item-quantity <> '0.000' AND item-quantity IS NOT INITIAL ).
      APPEND item.
    ENDIF.
  ENDLOOP.

  SORT item BY  po_number
                 po_item
                 ref_doc
                 ref_doc_year
                 ref_doc_item.

  DELETE ADJACENT DUPLICATES FROM item COMPARING po_number
                po_item
                ref_doc
                ref_doc_year
                ref_doc_item
                gl_account
                anln1
                anln2
                costcenter
                wbs_elem
                orderid
                tax_code_sap.
** Preencher estrutura cabecalho
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  APPEND header.


* Valida se o pedido do item pertence ao fornecedor que esta em cabeçalho
  LOOP AT item WHERE po_number IS NOT INITIAL.
    SELECT SINGLE bukrs lifnr waers FROM ekko INTO  (ekko-bukrs, ekko-lifnr, ekko-waers) WHERE ebeln = item-po_number.
    IF ekko-bukrs NE header-comp_code.

      MESSAGE i020(/sbxc/zckp_cockpit) WITH  header-comp_code.
*        exit.
*   Pedido Selecionado não pertence à empresa &
    ELSEIF ekko-lifnr NE header-vendor.
*      MESSAGE i049(/sbxc/zckp_cockpit) WITH  header-vendor.
      ""ODC - 31_03_2021
*        ELSEIF ekko-waers ne header-currency. "Validar se o pedido a associar está na mesma moeda do documento do cockpit
*          "ERRO: Moeda Cockpit & não corresponde à do Doc. &
*          MESSAGE i041(/sbxc/zckp_cockpit) WITH header-currency ekko-waers DISPLAY LIKE 'E'.
*          return.
      ""Fim ODC - 31_03_2021
    ENDIF.
  ENDLOOP.

  REFRESH: it_return.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
*  ENDLOOP.
* valida se forn tem IRF entao a factura tem de ser lançada via BI
    IF header-irf = 'X' AND e_ucomm = 'INT'.
      MESSAGE s010(/sbxc/zckp_cockpit).

    ELSE.
      LOOP AT header.

        CLEAR status_ok.
        PERFORM valida_status USING '0125'
                                    header-processo
                                    header-ano
                                    header-seqno
                                    ctrl-status1
                          CHANGING status_ok.

        CHECK status_ok = 'X'.


        indice = sy-tabix.
        REFRESH: item_f, it_return.
        CLEAR: invoicedocnumber, fiscalyear.
        LOOP AT item WHERE
          processo = header-processo AND
          ano      = header-ano      AND
          seqno    = header-seqno.

          MOVE-CORRESPONDING item TO item_f.
          APPEND item_f.
        ENDLOOP.
        MOVE-CORRESPONDING header TO wa_header.
        CLEAR simula_com.
        IF e_ucomm = 'SIMULA'.
          REFRESH:it_return.
          CALL FUNCTION '/SBXC/ZCKP_MM_SIMULA'
          "CCF 21.01.2022
            EXPORTING
              bline_date = lv_ldate
              "Fim
            IMPORTING
              ucomm      = simula_com
            TABLES
              itemdata   = item_f
              return     = it_return
            CHANGING
              headerdata = header
              ctrl       = ctrl.
          PERFORM grava_alteracoes TABLES item_f USING header.
        ENDIF.
        IF e_ucomm = 'INT' OR simula_com = 'POST'.
          REFRESH:it_return.
          CALL FUNCTION '/SBXC/ZCKP_MM_REGISTA_FACTURA'
          "CCF 21.01.2022
            EXPORTING
              bline_date       = lv_ldate
              "Fim
            IMPORTING
              invoicedocnumber = invoicedocnumber
              fiscalyear       = fiscalyear
*             ctrl             = ctrl
            TABLES
              itemdata         = item_f
            CHANGING
              headerdata       = header
              ctrl             = ctrl.
          PERFORM grava_alteracoes TABLES item_f USING header.
          refresh = 'X'.
        ELSEIF e_ucomm = 'PRE'.
          REFRESH:it_return.
          "ODC - 15_06_2020
          CALL FUNCTION '/SBXC/ZCKP_MM_REGIST_FAC_PREED'
          "CCF 21.01.2022
            EXPORTING
              bline_date = lv_ldate
              "Fim
*        IMPORTING
*             INVOICEDOCNUMBER       =
*             FISCALYEAR =
*             ctrl       = ctrl
            TABLES
              itemdata   = item_f
              return     = it_return
            CHANGING
              headerdata = header
              ctrl       = ctrl.

          IF header-doc_lo IS NOT INITIAL.
            invoicedocnumber = header-doc_lo.
          ENDIF.
        ELSEIF e_ucomm = 'INT2'.
          CALL FUNCTION '/SBXC/ZCKP_CALL_MIRO_HEADER'
          "CCF 21.01.2022
            EXPORTING
              bline_date = lv_ldate
              "Fim
*           IMPORTING
*             INVOICEDOCNUMBER       =
*             FISCALYEAR =
            TABLES
              itemdata   = item_f
              return     = it_return
            CHANGING
              headerdata = header
              ctrl       = ctrl.

          "Fim ODC - 15_06_2020
          PERFORM grava_alteracoes TABLES item_f USING header.
          refresh = 'X'.
*          READ TABLE it_return WITH KEY type = 'S'.
*          IF sy-subrc EQ 0.
*            invoicedocnumber = it_return-message_v1.
*          ENDIF.
          IF header-doc_lo IS NOT INITIAL.
            invoicedocnumber = header-doc_lo.
          ENDIF.
          "ODC - 16_06_2020
        ELSEIF e_ucomm = 'MIRO'.
          REFRESH:it_return.

          CALL FUNCTION '/SBXC/ZCKP_MM_REGIST_FAC_PREED'
          "CCF 21.01.2022
            EXPORTING
              bline_date = lv_ldate
              "Fim
            TABLES
              itemdata   = item_f
              return     = it_return
            CHANGING
              headerdata = header
              ctrl       = ctrl.
          PERFORM grava_alteracoes TABLES item_f USING header.
          refresh = 'X'.

*          READ TABLE it_return WITH KEY type = 'S'.
*          IF sy-subrc EQ 0.
*            invoicedocnumber = it_return-message_v1.
*          ENDIF.
          IF header-doc_lo IS NOT INITIAL.
            invoicedocnumber = header-doc_lo.
          ENDIF.
          "Fim ODC - 16_06_2020
        ENDIF.

        CHECK e_ucomm NE 'SIMULA'.
        MODIFY header.
* 'E' - Processado com exito

*    IF header-doc_fi <> '' AND e_ucomm = 'INT' AND header-status1 EQ 'E'.
*      header-status1 = '3'.
*    ELSEIF header-doc_fi <> '' AND e_ucomm = 'INT2' AND header-status1 NE 'E' AND header-status1 NE '5'.
*      header-status1 = '9'.
*    ELSEIF header-doc_fi <> '' AND e_ucomm = 'INT2' AND header-status1 EQ 'E'.
*      header-status1 = '3'.
*    ENDIF.
*    MODIFY header.

      ENDLOOP.
    ENDIF.
* Preencher estrutura cabecalho
*  move-corresponding cab to header.
    IF e_ucomm NE 'SIMULA'.
      REFRESH: linha.
      LOOP AT header.
        MOVE-CORRESPONDING header TO cab.
*    APPEND cab.
      ENDLOOP.
      LOOP AT item_f.
        MOVE-CORRESPONDING item_f TO linha.
        APPEND linha.
      ENDLOOP.
    ENDIF.
* Exibir msg de erro
    LOOP AT it_return.
      IF it_return-id = 'M8' AND it_return-number = '607'.
        it_return-type = 'S'.
        it_return-id = '/SBXC/ZCKP_COCKPIT'.
        it_return-number = '025'.
*      IT_RETURN-message = 'Entrar montante'.
        MODIFY it_return INDEX sy-tabix TRANSPORTING id number type.
      ENDIF.
    ENDLOOP.

    IF invoicedocnumber IS NOT INITIAL AND ( ctrl-status1 EQ '3' OR ctrl-status1 EQ 'E' ).
      it_return-type = 'S'.
      it_return-id = 'M8'.
      it_return-number = '60'.
      it_return-message_v1 = invoicedocnumber.
      APPEND it_return.


      "Envia para workflow
      ls_bim-bukrs = header-comp_code.
      ls_bim-belnr = header-doc_fi.
      ls_bim-gjahr = header-ano_lanc.
      ls_bim-doc_lo = header-doc_lo.
      ls_bim-mm_fi = c_mm.
      CREATE OBJECT lo_docinf.

      TRY.
          lo_docinf->send_to_bim_process( EXPORTING i_zsckp_to_bim = ls_bim
                                          IMPORTING e_wi_id = header-wi_id
                                                    return_code = lv_return ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*            ADD 1 TO wa_msg-lineno.
*            wa_msg-msgty = 'I'.
*            wa_msg-msgv1 =  lv_errortext.
*            APPEND wa_msg TO msg_ckp .

      ENDTRY.

      IF lv_return EQ '4'.
        wa_header-erro = TEXT-010.
      ELSEIF lv_return EQ '1'.
        wa_header-erro = TEXT-011.
      ENDIF.

      IF header-wi_id IS NOT INITIAL.
        UPDATE /sbxc/zckp_invh SET wi_id = header-wi_id
         WHERE processo EQ header-processo
          AND ano EQ header-ano
          AND seqno EQ header-seqno.
        COMMIT WORK AND WAIT.
      ELSEIF lv_return EQ '1' OR lv_return EQ '4'.
        UPDATE /sbxc/zckp_invh SET erro = header-erro
       WHERE processo EQ header-processo
        AND ano EQ header-ano
        AND seqno EQ header-seqno.
        COMMIT WORK AND WAIT.

      ENDIF.
      "Fim envio workflow
    ENDIF.

* Altera msg de erro para informação para não sair do cockpit
    LOOP AT  it_return.
      IF   it_return-type = 'E'.
        it_return-type = 'I'.
        MODIFY it_return TRANSPORTING type.
      ENDIF.
    ENDLOOP.
    CALL FUNCTION 'C14ALD_BAPIRET2_SHOW'
      TABLES
        i_bapiret2_tab = it_return.

*  READ TABLE it_return WITH KEY type = 'E'.
*  IF sy-subrc EQ 0.
*    MESSAGE ID it_return-id TYPE 'S' NUMBER it_return-number WITH it_return-message_v1
*    it_return-message_v2 it_return-message_v3 it_return-message_v4 DISPLAY LIKE 'E'.
*  ENDIF.
  ENDIF.
ENDFUNCTION.
