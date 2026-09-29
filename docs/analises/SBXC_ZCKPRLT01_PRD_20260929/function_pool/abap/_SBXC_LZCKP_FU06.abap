FUNCTION /sbxc/zckp_cancel_doc_lo.
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
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh,
        ano_doc_est LIKE header-ano_lanc.

  refresh = 'X'.
  REFRESH msg_cockpit.

  MOVE-CORRESPONDING cab TO header.

  IF header-doc_lo IS NOT INITIAL.

    REFRESH: lt_sval, t_messtab, bdcdata.

    lt_sval-tabname = 'BAPI_INCINV_FLD'.
    lt_sval-fieldname = 'REASON_REV'.
    lt_sval-field_obl = 'X'.
    APPEND lt_sval. CLEAR lt_sval.

    lt_sval-tabname = 'BAPI_INCINV_FLD'.
    lt_sval-fieldname = 'PSTNG_DATE'.

    APPEND lt_sval. CLEAR lt_sval.


    CALL FUNCTION 'POPUP_GET_VALUES'
      EXPORTING
*       NO_VALUE_CHECK  = ' '
        popup_title     = text-028 "'Dados Para Lançamento'
        start_column    = '10'
        start_row       = '5'
      IMPORTING
        returncode      = l_returncode
      TABLES
        fields          = lt_sval
      EXCEPTIONS
        error_in_fields = 1
        OTHERS          = 2.
    CHECK l_returncode NE 'A'.

    LOOP AT lt_sval.
      IF  lt_sval-fieldname = 'REASON_REV'.
        reason_rev = lt_sval-value.
      ELSEIF  lt_sval-fieldname = 'PSTNG_DATE'.
        data_lanc = lt_sval-value.
      ENDIF.
    ENDLOOP.

    IF data_lanc IS INITIAL OR data_lanc = ' '.


      CALL FUNCTION 'BAPI_INCOMINGINVOICE_CANCEL'
        EXPORTING
          invoicedocnumber          = header-doc_lo
          fiscalyear                = header-ano_lanc
          reasonreversal            = reason_rev
*         postingdate               = data_lanc
        IMPORTING
          invoicedocnumber_reversal = invoicedocnumber
          fiscalyear_reversal       = ano_doc_est
        TABLES
          return                    = return.

    ELSE.

      CALL FUNCTION 'BAPI_INCOMINGINVOICE_CANCEL'
        EXPORTING
          invoicedocnumber          = header-doc_lo
          fiscalyear                = header-ano_lanc
          reasonreversal            = reason_rev
          postingdate               = data_lanc
        IMPORTING
          invoicedocnumber_reversal = invoicedocnumber
          fiscalyear_reversal       = ano_doc_est
        TABLES
          return                    = return.
    ENDIF.
** chama função comensaçao
*    call function 'BAPI_TRANSACTION_COMMIT'
*     exporting
*       wait          = 'X'
**     IMPORTING
**       RETURN        =
*              .
*
*    if invoicedocnumber is not initial.
*       call function 'ZFI_MR8M_COMP'
*         exporting
*           belnr         = header-doc_lo
*           gjahr         = header-ano_lanc
*           belnr_e       = invoicedocnumber
*           gjahr_e       = ano_doc_est
*           lifnr         = header-vendor
*           bukrs         = header-comp_code
*           xblnr         = header-ref_doc_no.
*
*    endif.
    LOOP AT return.
      ADD 1 TO wa_msg-lineno.
      wa_msg-msgid = return-id.
      wa_msg-msgno = return-number.
      wa_msg-msgty = return-type.

      wa_msg-msgv1 =  return-message_v1.
      wa_msg-msgv2 =  return-message_v2.
      wa_msg-msgv3 =  return-message_v3.
      wa_msg-msgv4 =  return-message_v4.

      MOVE-CORRESPONDING wa_msg TO it_fimsg.
      APPEND it_fimsg.
      APPEND wa_msg TO msg_ckp.
*
    ENDLOOP.

    IF invoicedocnumber IS NOT INITIAL.
      SELECT SINGLE      cpudt  cputm usnam FROM  bkpf
                        INTO (bkpf-cpudt, bkpf-cputm, bkpf-usnam)
                        WHERE bukrs =  header-comp_code AND
                              belnr = invoicedocnumber AND
                              gjahr =  header-ano_lanc.
      UPDATE /sbxc/zckp_invh SET
      doc_estorno = invoicedocnumber
      ano_lanc = ' '
      doc_fi = ' '
      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

* verifica se existe processo anterior, se sim, actualiza também os dados do anterior
      SELECT SINGLE processo_ant ano_ant seqno_ant FROM /sbxc/zckp_ctrl
        INTO CORRESPONDING FIELDS OF /sbxc/zckp_ctrl WHERE
              processo = header-processo AND
             ano = header-ano AND
             seqno = header-seqno.

      IF  /sbxc/zckp_ctrl-processo_ant IS NOT INITIAL.
        UPDATE /sbxc/zckp_invh SET
            doc_estorno = invoicedocnumber
            ano_lanc = ' '
            doc_fi = ' '
            WHERE processo = /sbxc/zckp_ctrl-processo_ant AND
              ano = /sbxc/zckp_ctrl-ano_ant AND
              seqno = /sbxc/zckp_ctrl-seqno_ant.
      ENDIF.


      UPDATE /sbxc/zckp_ctrl SET
      status1 = '5'
           data_chg_st1 = bkpf-cpudt
           hora_chg_st1 = bkpf-cputm
           user_chg_st1 = bkpf-usnam
      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

      header-doc_estorno = invoicedocnumber.
*      header-status1 = '5'.

      "Envia status para Saphety
          CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
             EXPORTING
               cab            = header
               status         = 'CANCEL'.
          "Fim envio

      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          wait = 'X'.
*   IMPORTING
*     RETURN        =

      ADD 1 TO wa_msg-lineno.
      wa_msg-msgid = 'M8'.
      wa_msg-msgno = '060'.
      wa_msg-msgty = 'S'.
      wa_msg-msgv1 = invoicedocnumber  .
      MOVE-CORRESPONDING wa_msg TO it_fimsg.
      APPEND it_fimsg.
      APPEND wa_msg TO msg_ckp .

*      delete from zckp_inv_item where
*             processo = header-processo and
*             ano = header-ano and
*            seqno = header-seqno.
*
*
*      select * from zckp_hist_item into zckp_inv_item  where
*                    processo = header-processo and
*                   ano = header-ano and
*                   seqno = header-seqno.
*        insert zckp_inv_item.
*      endselect.

* se a factura tiver tido lançamentos de diferimentos é necessario proceder ao seu estorno
** actualiza com ultimo documento
*      SELECT SINGLE doc_fi_diferimen FROM /sbxc/zckp_invh INTO header-doc_fi_diferimen WHERE
*              processo = header-processo AND
*              ano = header-ano AND
*              seqno = header-seqno.
*
*      IF header-doc_fi_diferimen IS NOT INITIAL.
*        PERFORM preenche_bdcdata USING header-doc_fi_diferimen header-comp_code header-ano_lanc reason_rev.
*      ENDIF.
*
*      LOOP AT t_messtab INTO wa_messtab.
*        IF wa_messtab-msgid EQ 'F5' AND
*                      wa_messtab-msgnr  EQ '312'.
*
*          UPDATE /sbxc/zckp_invh SET
*                 doc_estorno_dife = wa_messtab-msgv1
*                 WHERE processo = header-processo AND
*                 ano = header-ano AND
*                 seqno = header-seqno.
*
*          ADD 1 TO wa_msg-lineno.
*
*          MOVE-CORRESPONDING wa_messtab TO wa_msg.
*          wa_msg-msgno = wa_messtab-msgnr.
*          wa_msg-msgid = 'F5'.
**        wa_msg-msgno = '312'.
*          wa_msg-msgty = wa_messtab-msgtyp.
**        wa_msg-msgv1 = wa_messtab-msgv1  .
*          APPEND wa_msg TO msg_ckp .
*          MOVE-CORRESPONDING wa_msg TO it_fimsg.
*          APPEND it_fimsg.
*        ENDIF.
*      ENDLOOP.


* Guarda msg na tabela de log
      PERFORM log_delete2 TABLES it_fimsg  USING header-processo header-ano header-seqno.
      PERFORM log2 TABLES it_fimsg USING header-processo header-ano header-seqno .

    ENDIF.
* Preencher estrutura cabecalho

    REFRESH:   linha.
*    header-ano_lanc = ' '.
    header-doc_estorno = invoicedocnumber.
    MOVE-CORRESPONDING header TO cab.

  ENDIF.

ENDFUNCTION.
