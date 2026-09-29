FUNCTION /sbxc/zckp_sendsaphety.
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
*{   INSERT         DEVK938786                                        1

  DATA: header    TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        lv_cod    TYPE /sbxc/zckp_tab09-cod_mes,
        lv_return TYPE sy-subrc.
* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.

  CASE ctrl-status1.
    WHEN '3' OR '4'.
      "Envia status para Saphety
      CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
        EXPORTING
          cab    = header
          status = 'ACCOUNTED'
        IMPORTING
          return = lv_return.
      "Fim envio
      IF lv_return EQ 0.
        MESSAGE s072(/sbxc/zckp_cockpit) .
      ENDIF.
    WHEN '6'.
      CLEAR lv_cod.
      SELECT SINGLE cod_mes INTO lv_cod
        FROM /sbxc/zckp_tab09
        WHERE processo EQ header-processo
        AND ano EQ header-ano
        AND seqno EQ header-seqno.
      IF lv_cod IS NOT INITIAL.

        CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
          EXPORTING
            cab     = header
            cod_mes = lv_cod
            status  = 'REJECTED'
            accao   = header-mot_n_contab
          IMPORTING
            return  = lv_return.
        IF lv_return EQ 0.
          MESSAGE s072(/sbxc/zckp_cockpit) .

*          UPDATE /sbxc/zckp_ctrl SET
*            status3 = '8'
*            data_chg_st3 = sy-datum
*            hora_chg_st3 = sy-uzeit
*            user_chg_st3 = sy-uname
*            WHERE processo = header-processo AND
*            ano = header-ano AND
*            seqno = header-seqno.
        ELSE.
          MESSAGE s073(/sbxc/zckp_cockpit) .
        ENDIF.
      ENDIF.
    WHEN OTHERS.
      MESSAGE s071(/sbxc/zckp_cockpit) .
  ENDCASE.


*}   INSERT
ENDFUNCTION.
