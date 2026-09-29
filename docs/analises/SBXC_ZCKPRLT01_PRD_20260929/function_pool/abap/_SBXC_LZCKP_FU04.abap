FUNCTION /sbxc/zckp_cancel_doc_fi.
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

  DATA: wa_header TYPE /sbxc/zckp_invh, wa_msg LIKE LINE OF msg_ckp.
  REFRESH msg_cockpit.

  refresh = 'X'.
  MOVE-CORRESPONDING cab TO header.


  IF header-doc_fi IS NOT INITIAL.

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
        popup_title     = TEXT-028 "'Dados Para Lançamento'
        start_column    = '10'
        start_row       = '5'
      IMPORTING
        returncode      = l_returncode
      TABLES
        fields          = lt_sval
      EXCEPTIONS
        error_in_fields = 1
        OTHERS          = 2.

    IF l_returncode NE 'A'.

      LOOP AT lt_sval.
        IF  lt_sval-fieldname = 'REASON_REV'.
          reason_rev = lt_sval-value.
        ELSEIF  lt_sval-fieldname = 'PSTNG_DATE'.
          data_lanc = lt_sval-value.
        ENDIF.
      ENDLOOP.

      PERFORM preenche_bdcdata USING  header-doc_fi header-comp_code header-ano_lanc reason_rev.

      LOOP AT t_messtab INTO wa_messtab.
        IF wa_messtab-msgid EQ 'F5' AND
                     wa_messtab-msgnr  EQ '312'.

          UPDATE /sbxc/zckp_invh SET
                 doc_estorno = wa_messtab-msgv1
                 WHERE processo = header-processo AND
                 ano = header-ano AND
                 seqno = header-seqno.

          SELECT SINGLE  cpudt  cputm usnam FROM  bkpf
                               INTO (bkpf-cpudt, bkpf-cputm, bkpf-usnam)
                               WHERE bukrs =  header-comp_code AND
                                     belnr = wa_messtab-msgv1 AND
                                     gjahr =  header-ano_lanc.

          UPDATE /sbxc/zckp_ctrl SET
          status1 = '5'
             data_chg_st1 = bkpf-cpudt
             hora_chg_st1 = bkpf-cputm
             user_chg_st1 = bkpf-usnam
          WHERE processo = header-processo AND
          ano = header-ano AND
          seqno = header-seqno.

          header-doc_estorno = wa_messtab-msgv1.

          "Envia status para Saphety
          CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
            EXPORTING
              cab    = header
              status = 'CANCEL'.
          "Fim envio

        ENDIF.
        ADD 1 TO wa_msg-lineno.

        MOVE-CORRESPONDING wa_messtab TO wa_msg.
        wa_msg-msgno = wa_messtab-msgnr.
        wa_msg-msgid = 'F5'.
        wa_msg-msgty = wa_messtab-msgtyp.
        APPEND wa_msg TO msg_ckp .

      ENDLOOP.

      COMMIT WORK AND WAIT.

* Preencher estrutura cabecalho
      REFRESH:  linha.

      header-doc_estorno = invoicedocnumber.
      header-ano_lanc = ' '.
      MOVE-CORRESPONDING header TO cab.
    ENDIF.
  ENDIF.


ENDFUNCTION.
