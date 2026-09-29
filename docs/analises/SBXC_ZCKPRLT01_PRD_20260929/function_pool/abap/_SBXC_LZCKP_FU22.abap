FUNCTION /sbxc/zckp_mm_liga_ulterior.
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
  DATA: header    TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.
  DATA: item      TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: item_f    TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: l_status TYPE /sbxc/zckp_ctrl-status1,
        l_doc    TYPE rbkp-belnr.

  DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: ls_bim TYPE zsckp_to_bim,
        lv_errortext   TYPE string.
  CONSTANTS: c_fi TYPE zsckp_to_bim-mm_fi VALUE 'FI',
             c_mm TYPE zsckp_to_bim-mm_fi VALUE 'MM'.
DATA: lv_return TYPE sy-subrc.

  REFRESH msg_cockpit.

  TABLES bseg.

* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.

    IF  ctrl-status1 = '9' OR  ctrl-status1 = '3' OR  ctrl-status1 = '4' OR ctrl-status1 = '6'.
      "Status do documento não permite esta operação
      MESSAGE s026(/sbxc/zckp_cockpit).
    ELSE.

    LOOP AT linha.
      MOVE-CORRESPONDING linha TO item.
      APPEND item.
    ENDLOOP.

    SET PARAMETER ID 'BUK' FIELD header-comp_code.

    CALL FUNCTION 'ARCHIV_POPUP_OBJECT_KEY'
      EXPORTING
        objtype         = 'BKPF'
      IMPORTING
        objkey          = obj
      EXCEPTIONS
        error_parameter = 1
        user_cancel     = 2
        OTHERS          = 3.

    CHECK sy-subrc = 0.

* Verificar se documento existe
    SELECT SINGLE belnr bukrs gjahr cpudt cputm usnam awtyp awkey bldat stblg xblnr waers
    INTO (bkpf-belnr, bkpf-bukrs, bkpf-gjahr, bkpf-cpudt,  bkpf-cputm, bkpf-usnam, bkpf-awtyp, bkpf-awkey, bkpf-bldat, bkpf-stblg, bkpf-xblnr, bkpf-waers )
    FROM bkpf
    WHERE bukrs = obj(4)     AND
    belnr = obj+4(10)  AND
    gjahr = obj+14(4).


    IF sy-subrc NE 0.
      MESSAGE i019(/sbxc/zckp_cockpit).
*   Documento para Empresa/Nº documento/Exercício inexistente
    ELSE.

      IF bkpf-awtyp EQ 'RMRP'.
        l_status = '3'. "Integrado via logistica
        l_doc = bkpf-awkey(10).
      ELSE.
        l_status = '4'. "Integrado via financeira
      ENDIF.

*Verificar se é doc de fornecedor
      DATA: l_koart LIKE bseg-koart.
      CLEAR l_koart.
      SELECT SINGLE koart INTO l_koart
        FROM bseg
        WHERE bukrs = obj(4)     AND
              belnr = obj+4(10)  AND
              gjahr = obj+14(4)  AND
              koart = 'K'.
      IF sy-subrc NE 0.
        CALL FUNCTION 'POPUP_TO_CONFIRM'
          EXPORTING
            text_question         = TEXT-035 "'Documento não é de fornecedor. Pretende continuar com a associação?'
            text_button_1         = 'Sim'(001)
            text_button_2         = 'Não'(002)
            display_cancel_button = ' '
          IMPORTING
            answer                = ret
          EXCEPTIONS
            text_not_found        = 1
            OTHERS                = 2.

        CHECK ret = '1'.
      ENDIF.

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

      READ TABLE header INDEX 1.
      MOVE-CORRESPONDING header TO wa_header.

      wa_header-doc_fi = obj+4(10).
      wa_header-doc_lo = l_doc.
      wa_header-ano_lanc = obj+14(4).

      MODIFY header FROM wa_header INDEX 1.

* verifica se documento já associado
      SELECT SINGLE * FROM  /sbxc/zckp_invh  "#EC CI_NOORDER
          WHERE
              comp_code = obj(4) AND
             doc_fi   = wa_header-doc_fi AND
             ano_lanc = wa_header-ano_lanc.
      IF sy-subrc = 0.
        MESSAGE i038(/sbxc/zckp_cockpit) WITH obj(4) wa_header-doc_fi  wa_header-ano_lanc /sbxc/zckp_invh-seqno.
      ELSE.

* valida campos fornecedor/empresa/atribuiçã/data documento/moeda/bstat/fornecedor
*        SELECT SINGLE lifnr wrbtr
*        INTO (bseg-lifnr, bseg-wrbtr) FROM bseg
*        WHERE bukrs = bkpf-bukrs
*          AND belnr = bkpf-belnr
*          AND gjahr = bkpf-gjahr
*          AND koart = 'K'
*          AND lifnr = wa_header-vendor.

        DATA: lv_wrbtr TYPE bseg-wrbtr.
        CLEAR lv_wrbtr.
        SELECT lifnr wrbtr
       INTO (bseg-lifnr, bseg-wrbtr) FROM bseg
       WHERE bukrs = bkpf-bukrs
         AND belnr = bkpf-belnr
         AND gjahr = bkpf-gjahr
         AND koart = 'K'
         AND lifnr = wa_header-vendor
          ORDER BY buzei ASCENDING.
          lv_wrbtr = lv_wrbtr + bseg-wrbtr.
        ENDSELECT.
        bseg-wrbtr = lv_wrbtr.

        IF wa_header-comp_code <> obj(4).
* Erro empresa do documento diferente da empresa da linha seleccionada
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '002'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-bukrs.
          wa_msg-msgv2 = obj(4).
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF bkpf-xblnr <> wa_header-ref_doc_no.
*          Associação Impossível: Referência Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '053'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-xblnr.
          wa_msg-msgv2 = wa_header-ref_doc_no.
          APPEND wa_msg TO msg_cockpit .

        ENDIF.

        IF   bkpf-bldat <> wa_header-doc_date.
*          Associação Impossível: DataDoc Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '040'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-bldat.
          wa_msg-msgv2 = wa_header-doc_date.
          APPEND wa_msg TO msg_cockpit .

        ENDIF.

        IF bkpf-waers <> wa_header-currency.
*          Associação Impossível: Moeda Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '041'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-waers.
          wa_msg-msgv2 = wa_header-currency.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF   bkpf-stblg <> ' '.
*          Associação Impossível: Documento & Estornado
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '042'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-belnr.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF bseg-lifnr <> wa_header-vendor.
*          Associação Impossível: Fornecedor Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '043'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bseg-lifnr.
          wa_msg-msgv2 = wa_header-vendor.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF bseg-wrbtr <> wa_header-gross_amount.
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '046'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bseg-wrbtr.
          wa_msg-msgv2 = wa_header-gross_amount.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF  msg_cockpit[] IS NOT INITIAL .
          CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
            TABLES
              i_message_tab = msg_cockpit.

          RETURN.
        ELSE.


          UPDATE /sbxc/zckp_invh
             SET doc_fi   = wa_header-doc_fi
                 doc_lo   = wa_header-doc_lo
                 ano_lanc = wa_header-ano_lanc
                 doc_estorno = '          '
                 data_criacao = bkpf-cpudt
           WHERE processo = header-processo AND
                 ano = header-ano AND
                 seqno = header-seqno.

          COMMIT WORK.

*  Se se tratar de uma NC, enviar email para o fornecedor
          DATA: ls_nc.
          CLEAR ls_nc.
          CALL FUNCTION '/SBXC/ZCKP_CHECK_NC'
            EXPORTING
              i_bukrs  = wa_header-comp_code
              i_doc_fi = wa_header-doc_fi
              i_gjahr  = wa_header-ano_lanc
            IMPORTING
              e_nc     = ls_nc.
          IF sy-subrc EQ 0 AND ls_nc EQ 'X'.
*             Enviar email para o fornecedor
            CALL FUNCTION '/SBXC/ZCKP_SENDMAIL_FORN'
              EXPORTING
                i_bukrs      = wa_header-comp_code
                i_ref_doc_no = wa_header-ref_doc_no
                i_gjahr      = wa_header-ano_lanc
                i_vendor     = wa_header-vendor
                i_doc_date   = wa_header-doc_date
                i_pstng_date = wa_header-pstng_date
                i_nc         = ls_nc.
          ENDIF.
*  Se se tratar de uma NC, enviar email para o fornecedor


          UPDATE /sbxc/zckp_ctrl SET
             status1 = l_status
             status2 = 'M' "Actualizado manualmente
                            data_chg_st1 = bkpf-cpudt
           hora_chg_st1 = bkpf-cputm
           user_chg_st1 = bkpf-usnam
             WHERE processo = header-processo AND
              ano = header-ano AND
              seqno = header-seqno.

          "Envia status para Saphety
          CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
            EXPORTING
              cab    = header
              status = 'ACCOUNTED'.

          COMMIT WORK.

          IF header-url IS NOT INITIAL.


            CALL FUNCTION '/SBXC/ZCKP_ANEXA_URL'
              EXPORTING
                bukrs = header-comp_code
                belnr = wa_header-doc_fi
                gjahr = header-pstng_date(4)
                url   = header-url.
        ENDIF.

           "Envia para workflow
           "Verificar se existe algum bloqueio de pagamento
           SELECT SINGLE zlspr INTO @DATA(lv_ZLSPR)
             FROM bseg
             WHERE bukrs EQ @wa_header-comp_code AND
                   belnr EQ @wa_header-doc_fi AND
                   gjahr EQ @wa_header-ano_lanc AND
                   koart EQ 'K'.
             IF lv_zlspr IS NOT INITIAL.

            ls_bim-bukrs = wa_header-comp_code.
            ls_bim-belnr = wa_header-doc_fi.
            ls_bim-gjahr = wa_header-ano_lanc.
            CASE l_status.
              WHEN '3'.
                ls_bim-doc_lo = wa_header-doc_lo.
                ls_bim-mm_fi = c_mm.
              WHEN '4'.
                ls_bim-mm_fi = c_fi.
            ENDCASE.


            CREATE OBJECT lo_docinf.

            TRY.
                lo_docinf->send_to_bim_process( EXPORTING i_zsckp_to_bim = ls_bim
                                                IMPORTING e_wi_id = wa_header-wi_id
                                                          return_code = lv_return ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*           wa_header-erro = lv_errortext.
*
            IF lv_return EQ '4'.
              wa_header-erro = TEXT-010.
              ELSEIF lv_return EQ '1'.
                wa_header-erro = TEXT-011.
            ENDIF.

               UPDATE /sbxc/zckp_invh SET erro = wa_header-erro
               WHERE processo EQ wa_header-processo
                AND ano EQ wa_header-ano
                AND seqno EQ wa_header-seqno.
               COMMIT WORK AND WAIT.
            ENDTRY.

            IF wa_header-wi_id IS NOT INITIAL .
              UPDATE /sbxc/zckp_invh SET wi_id = wa_header-wi_id
                WHERE processo EQ wa_header-processo
                AND ano EQ wa_header-ano
                AND seqno EQ wa_header-seqno.
              COMMIT WORK AND WAIT.
            ENDIF.

             ENDIF.
            "Fim envio workflow

        ENDIF.
      ENDIF.
      IF wa_header-doc_fi <> ''.
        REFRESH: cor_tab[], cor[].
        cor[] = cor_tab[].
      ENDIF.

* Preencher estrutura cabecalho
      REFRESH: linha.
      CLEAR cab.

      READ TABLE header INDEX 1.
      MOVE-CORRESPONDING header TO cab.

      LOOP AT item.
        MOVE-CORRESPONDING item TO linha.
        APPEND linha.
      ENDLOOP.

      CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
        TABLES
          i_message_tab = msg_cockpit.

      COMMIT WORK AND WAIT.

      refresh = 'X'.
    ENDIF.
    ENDIF.
  ENDIF.
ENDFUNCTION.
