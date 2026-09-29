FUNCTION /sbxc/zckp_rfbibl00.
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
  DATA: item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: simula_com LIKE sy-ucomm.
  DATA: ls_cab TYPE /sbxc/zckp_invh.
  DATA: lv_status1 TYPE /sbxc/zckp_ctrl-status1. "ODC - 22_06_2020

  DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: ls_bim       TYPE zsckp_to_bim,
        lv_errortext TYPE string.
  CONSTANTS: c_fi TYPE zsckp_to_bim-mm_fi VALUE 'FI'.
  DATA: lv_return TYPE sy-subrc.

  DATA: l_message   TYPE string,
        o_exception TYPE REF TO cx_root.

  DATA: BEGIN OF xbltab OCCURS 1.
          INCLUDE STRUCTURE blntab.
  DATA: END OF xbltab.

  MOVE-CORRESPONDING cab TO ls_cab.
  MOVE-CORRESPONDING cab TO it_cab.
  APPEND it_cab.

* CCF 21.01.2022
* Se a condição de pagamento tem o campo “Dia fixo” preenchido (T052 - ZFAEL), o sistema faz EOMONTH da “data base”.
  DATA: zfael    LIKE t052-zfael,
        lv_ldate TYPE sy-datum.

  SELECT SINGLE zfael INTO zfael FROM t052 WHERE zterm = ls_cab-zterm.
  IF zfael = '31'.
    CALL FUNCTION 'LAST_DAY_OF_MONTHS'
      EXPORTING
        day_in            = ls_cab-baseline_date
      IMPORTING
        last_day_of_month = lv_ldate
*  EXCEPTIONS
*       DAY_IN_NO_DATE    = 1
*       OTHERS            = 2
      .
*    IF sy-subrc = 0.
*      ls_cab-baseline_date = lv_ldate.
*    ENDIF.
  else.
    lv_ldate = ls_cab-baseline_date.
  ENDIF.
  "FIM CCF
  lv_status1 = ctrl-status1. "ODC - 22_06_2020

  "Validar se a data de documento é posterior à data de lançamento
  IF it_cab-doc_date > it_cab-pstng_date.
    MESSAGE s064(/sbxc/zckp_cockpit) DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  IF it_cab-em_tratamento = 'X' AND it_cab-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH it_cab-user_tratamento.
  ELSE.
    LOOP AT linha.
      MOVE-CORRESPONDING linha TO it_lin.
      APPEND it_lin.
    ENDLOOP.
    PERFORM inicializa_estruturas USING c_nodata.

    DATA prog(7).
    prog = 'COCKPIT'.
    EXPORT prog TO MEMORY ID  'PROG'.

    CONCATENATE 'COCKPIT_DOC' sy-uname sy-datum sy-uzeit INTO ds_name.
*    DATA mess(60).

    TRY .

        OPEN DATASET ds_name FOR OUTPUT IN TEXT MODE ENCODING DEFAULT MESSAGE l_message.
        IF sy-subrc IS NOT INITIAL.
          MESSAGE l_message
             TYPE 'S' DISPLAY LIKE 'E'.
        ENDIF.
      CATCH cx_root
         INTO o_exception.

*     Gets error message
        CALL METHOD o_exception->if_message~get_text
          RECEIVING
            result = l_message.

        MESSAGE l_message
           TYPE 'S' DISPLAY LIKE 'E'.

    ENDTRY.

    bgr00-stype = '0'.
    bgr00-mandt = sy-mandt.
    bgr00-group = 'COCKPIT'.
    bgr00-usnam  = sy-uname.
    bgr00-nodata = c_nodata.
    bgr00-xkeep = 'X'.
    TRANSFER bgr00 TO ds_name.

    LOOP AT it_cab.
      IF it_cab-comp_code IS INITIAL.
        MESSAGE s021(/sbxc/zckp_cockpit).
      ELSEIF   it_cab-doc_type IS INITIAL.
        MESSAGE s022(/sbxc/zckp_cockpit).
      ELSE.
        IF ctrl-status1 = '6'.
          MESSAGE s018(/sbxc/zckp_cockpit).
        ELSEIF ctrl-status1 = '4'.
          MESSAGE s012(/sbxc/zckp_cockpit).
        ELSE.
          REFRESH: item_f, it_return.
          LOOP AT it_lin WHERE
            processo = it_cab-processo AND
            ano      = it_cab-ano      AND
            seqno    = it_cab-seqno.
            MOVE-CORRESPONDING it_lin TO item_f.
            APPEND item_f.
          ENDLOOP.
          IF e_ucomm = 'SIMULAFI'.
            REFRESH: it_return.
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
                headerdata = it_cab
                ctrl       = ctrl.
            MOVE-CORRESPONDING it_cab TO ls_cab.
            PERFORM grava_alteracoes TABLES it_lin USING ls_cab."ODC - 24_07_2018
          ENDIF.
          IF e_ucomm <> 'SIMULAFI' OR simula_com = 'POST'.
            PERFORM init_bbkpf USING bbkpf.
            PERFORM preenche_cab USING e_ucomm it_cab.
            PERFORM bbseg_forn USING e_ucomm  it_cab lv_ldate.


            LOOP AT it_lin WHERE processo = it_cab-processo AND
                                 ano      = it_cab-ano AND
                                 seqno    = it_cab-seqno.
              IF it_lin-gl_account IS INITIAL AND ( it_lin-anln1 IS INITIAL OR it_lin-anln2 IS INITIAL ).
                MESSAGE s058(/sbxc/zckp_cockpit).
                EXIT.
              ELSEIF it_lin-db_cr_ind IS INITIAL.
                MESSAGE s024(/sbxc/zckp_cockpit).
                EXIT.
              ENDIF.
              PERFORM bbseg_lin USING e_ucomm it_lin.
            ENDLOOP.
            CLOSE DATASET ds_name.

*Abrir imagem
            SELECT SINGLE * FROM /sbxc/zckp_img
              WHERE processo = it_cab-processo
              AND seqno = it_cab-seqno
              AND ano = it_cab-ano.
            IF sy-subrc EQ 0.

              CALL FUNCTION 'ARCHIVOBJECT_DISPLAY'
                EXPORTING
                  archiv_doc_id = /sbxc/zckp_img-arc_doc_id
                  archiv_id     = /sbxc/zckp_img-archiv_id.

            ENDIF.
            DATA: l_callmode, l_anzmode.

            IF e_ucomm EQ 'F-41B' OR e_ucomm EQ 'F-43B' OR e_ucomm EQ 'FB01B'.
              l_callmode = 'C'.
              l_anzmode = 'N'.
            ELSE.
              l_callmode = 'C'.
              l_anzmode = 'A'.
            ENDIF.

            REFRESH: xbltab, t_fimsg.
            EXPORT xbltab TO MEMORY ID 'FI_XBLTAB'.
            EXPORT t_fimsg TO MEMORY ID 'ZFE_FIMSG'.
            FREE MEMORY ID 'FI_XBLTAB'.
            FREE MEMORY ID 'ZFE_FIMSG'.
            CLEAR bkpf.

            SUBMIT /sbxc/rfbibl00
                    WITH anz_mode = l_anzmode
                  WITH callmode = l_callmode
                  WITH ds_name = ds_name
                  WITH update = 'S'
                  WITH xinf = 'X'
            AND RETURN.

            IMPORT xbltab FROM  MEMORY ID 'FI_XBLTAB'.
            READ TABLE xbltab INDEX 1.

            bkpf-belnr = xbltab-belnr.
            bkpf-bukrs = xbltab-bukrs.
            bkpf-gjahr = xbltab-gjahr.

            IMPORT t_fimsg FROM MEMORY ID 'ZFE_FIMSG'.

            IF bkpf-belnr IS NOT INITIAL.

              REFRESH return.
              UPDATE /sbxc/zckp_invh SET
                     comp_code = bbkpf-bukrs
                     doc_fi = bkpf-belnr
                     doc_lo = ' '
                     doc_estorno = ' '
                     ano_lanc = xbltab-gjahr
                     data_criacao = sy-datum
               WHERE processo = it_cab-processo AND
                     ano = it_cab-ano AND
                     seqno = it_cab-seqno.

              SELECT SINGLE      cpudt  cputm usnam FROM  bkpf
                               INTO (bkpf-cpudt, bkpf-cputm, bkpf-usnam)
                               WHERE bukrs =   bbkpf-bukrs AND
                                     belnr = bkpf-belnr AND
                                     gjahr =  xbltab-gjahr.

              UPDATE /sbxc/zckp_ctrl SET
              status1 = '4'
               data_chg_st1 = bkpf-cpudt
           hora_chg_st1 = bkpf-cputm
           user_chg_st1 = bkpf-usnam
              WHERE processo = it_cab-processo AND
              ano = it_cab-ano AND
              seqno = it_cab-seqno.

              COMMIT WORK AND WAIT.

              it_cab-doc_fi = bkpf-belnr.
              it_cab-ano_lanc = bkpf-gjahr."it_cab-pstng_date(4).
              it_cab-doc_estorno = ' '.
              ctrl-status1 = '4'.
              MOVE-CORRESPONDING it_cab TO ls_cab.
              MODIFY it_cab.
              MOVE-CORRESPONDING it_cab TO ls_cab.
              PERFORM grava_alteracoes TABLES it_lin USING ls_cab.

              "Envia para workflow
              ls_bim-bukrs = ls_cab-comp_code.
              ls_bim-belnr = ls_cab-doc_fi.
              ls_bim-gjahr = ls_cab-ano_lanc.
              ls_bim-mm_fi = c_fi.

              CREATE OBJECT lo_docinf.

              TRY.
                  lo_docinf->send_to_bim_process( EXPORTING i_zsckp_to_bim = ls_bim
                                                  IMPORTING e_wi_id = ls_cab-wi_id
                                                            return_code = lv_return ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*           ls_cab-erro = lv_errortext.

                  IF lv_return EQ '4'.
                    ls_cab-erro = TEXT-010.
                  ELSEIF lv_return EQ '1'.
                    ls_cab-erro = TEXT-011.
                  ENDIF.

                  UPDATE /sbxc/zckp_invh SET erro = ls_cab-erro
                  WHERE processo EQ ls_cab-processo
                   AND ano EQ ls_cab-ano
                   AND seqno EQ ls_cab-seqno.
                  COMMIT WORK AND WAIT.
              ENDTRY.

              IF ls_cab-wi_id IS NOT INITIAL .
                UPDATE /sbxc/zckp_invh SET wi_id = ls_cab-wi_id
                  WHERE processo EQ ls_cab-processo
                  AND ano EQ ls_cab-ano
                  AND seqno EQ ls_cab-seqno.
                COMMIT WORK AND WAIT.
              ENDIF.

              "Fim envio workflow


              "Envia status para Saphety
              CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
                EXPORTING
                  cab    = ls_cab
                  status = 'ACCOUNTED'.
              "Fim envio

* Atualiza itens na tabela de cockpit
              DELETE FROM /sbxc/zckp_invi WHERE processo = it_cab-processo AND
                     ano = it_cab-ano AND
                     seqno = it_cab-seqno.
              MOVE-CORRESPONDING it_cab TO /sbxc/zckp_invi.
              SELECT buzei wrbtr gsber kostl matnr hkont shkzg mwskz anln1 anln2 zuonr sgtxt FROM bseg
              INTO (/sbxc/zckp_invi-invoice_doc_item, /sbxc/zckp_invi-item_amount,
              /sbxc/zckp_invi-gsber, /sbxc/zckp_invi-costcenter, /sbxc/zckp_invi-matnr,
              /sbxc/zckp_invi-gl_account, /sbxc/zckp_invi-db_cr_ind,   /sbxc/zckp_invi-tax_code_sap, /sbxc/zckp_invi-anln1, /sbxc/zckp_invi-anln2,
              /sbxc/zckp_invi-zuonr, /sbxc/zckp_invi-item_text)
              WHERE bukrs = bkpf-bukrs AND
              belnr = bkpf-belnr AND
              gjahr = bkpf-gjahr AND
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
              DATA: ls_nc.
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



            ELSE.
              "ODC - 20_06_2020
              "Verificar se ocorreu algum erro ou se apenas foi cancelado o processo
              "Se ocorrer algum erro coloca o status 2 senao mantém o status inicial.
              READ TABLE t_fimsg WITH KEY msgty = 'E'.
              IF sy-subrc EQ 0.
                UPDATE /sbxc/zckp_ctrl SET
                  status1 = '2'
                  WHERE processo = it_cab-processo AND
                  ano = it_cab-ano AND
                  seqno = it_cab-seqno.
                ctrl-status1 = '2'.
              ELSE.
                UPDATE /sbxc/zckp_ctrl SET
                  status1 = lv_status1
                  WHERE processo = it_cab-processo AND
                  ano = it_cab-ano AND
                  seqno = it_cab-seqno.
                ctrl-status1 = lv_status1.
              ENDIF.
*              UPDATE /sbxc/zckp_ctrl SET
*              status1 = '2'
*              WHERE processo = it_cab-processo AND
*              ano = it_cab-ano AND
*              seqno = it_cab-seqno.
*              ctrl-status1 = '2'.
              "Fim ODC - 20_06_2020
              COMMIT WORK AND WAIT.
              PERFORM grava_alteracoes TABLES it_lin USING ls_cab.


              READ TABLE t_fimsg WITH KEY msgty = 'E'.
              IF sy-subrc EQ 0.
                MESSAGE ID t_fimsg-msgid TYPE 'S' NUMBER t_fimsg-msgno WITH t_fimsg-msgv1
                t_fimsg-msgv2 t_fimsg-msgv3 t_fimsg-msgv4 DISPLAY LIKE 'E'.
              ELSE.
                MESSAGE ID '/SBXC/ZCKP_COCKPIT' TYPE 'S' NUMBER 082 DISPLAY LIKE 'E'.
              ENDIF.

            ENDIF.
          ELSEIF e_ucomm = 'SIMULAFI'.
            LOOP AT it_return.
              t_fimsg-msgid = it_return-id.
              t_fimsg-msgty = it_return-type .
              t_fimsg-msgno = it_return-number.
              t_fimsg-msgv1 = it_return-message_v1.
              t_fimsg-msgv2 = it_return-message_v2.
              t_fimsg-msgv3 = it_return-message_v3.
              t_fimsg-msgv4 = it_return-message_v4.
              APPEND t_fimsg.
            ENDLOOP.
            CALL FUNCTION 'C14ALD_BAPIRET2_SHOW'
              TABLES
                i_bapiret2_tab = it_return.
          ENDIF.
* apaga tabela de log
          PERFORM log_delete2 TABLES t_fimsg USING it_cab-processo it_cab-ano it_cab-seqno .
* Guarda msg na tabela de log

          PERFORM log2 TABLES t_fimsg USING it_cab-processo it_cab-ano it_cab-seqno .
        ENDIF.
      ENDIF.
      REFRESH: xbltab, t_fimsg.
      EXPORT xbltab TO MEMORY ID 'FI_XBLTAB'.
      EXPORT t_fimsg TO MEMORY ID 'ZFE_FIMSG'.
      FREE MEMORY ID 'FI_XBLTAB'.
      FREE MEMORY ID 'ZFE_FIMSG'.
      CLEAR bkpf.
    ENDLOOP.

    CLEAR cab.
    LOOP AT it_cab.
      MOVE-CORRESPONDING it_cab TO cab.
    ENDLOOP.

    refresh = 'X'.
  ENDIF.
ENDFUNCTION.
