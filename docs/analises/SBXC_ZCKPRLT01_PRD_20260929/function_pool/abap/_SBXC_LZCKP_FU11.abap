FUNCTION /sbxc/zckp_rej2.
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
  DATA: wa_header TYPE /sbxc/zckp_invh.

  DATA: item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.


  DATA: l_returncode,
        lt_sval LIKE sval OCCURS 0 WITH HEADER LINE.

  REFRESH msg_cockpit.

  DATA: BEGIN OF mensagens OCCURS 0,
          cod_mes TYPE /sbxc/zckp_tab06-cod_mes,
          split,
          mensagem TYPE /sbxc/zckp_tab07-mensagem,
        END OF mensagens.

  DATA: est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem.
  DATA: processo_ant LIKE /sbxc/zckp_ctrl-processo.


* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.


  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  IF opcao GE 1.

* Ler parametrização da mensagem seleccionada
    READ TABLE mensagens INDEX opcao.
* Verificar se todas as linhas seleccionadas
    SELECT SINGLE *
      FROM /sbxc/zckp_tab06 WHERE est_mensagem = est_mensagem AND
                            cod_mes = mensagens-cod_mes.




    /sbxc/zckp_tab09-processo = header-processo.
    /sbxc/zckp_tab09-ano = header-ano.
    /sbxc/zckp_tab09-seqno = header-seqno.
    /sbxc/zckp_tab09-est_mensagem = /sbxc/zckp_tab06-est_mensagem.
    /sbxc/zckp_tab09-cod_mes = /sbxc/zckp_tab06-cod_mes.
    CONCATENATE /sbxc/zckp_tab06-prefixo mensagens-mensagem INTO /sbxc/zckp_tab09-mensagem.
    /sbxc/zckp_tab09-processo_seg = /sbxc/zckp_tab06-destino.
    INSERT /sbxc/zckp_tab09.



    ctrl-status1 = '6'.
*   Se mudar de processo
    IF /sbxc/zckp_tab06-destino NE space AND
      /sbxc/zckp_tab06-destino <> header-processo.

      processo_ant = header-processo.
      header-processo = /sbxc/zckp_tab06-destino.


      CALL FUNCTION '/SBXC/ZCKP_MM_CRIA_FACTURAS'
        EXPORTING
         headerdata             = header
         processo_ant           = processo_ant
         ano_ant                = header-ano
         seqno_ant              = header-seqno
*          IMPORTING
*            INVOICEDOCNUMBER       =
*            FISCALYEAR             =
        TABLES
          itemdata               = item
          return                 = return
                .

      UPDATE /sbxc/zckp_ctrl SET
      status1 = '6'
      data_chg_st1 = sy-datum
      hora_chg_st1 = sy-uzeit
      user_chg_st1 = sy-uname

      processo_post = processo_ant
      ano_post = header-ano
      seqno_post = header-seqno

      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

    ELSE.

      UPDATE /sbxc/zckp_ctrl SET
       status1 = '6'
       data_chg_st1 = sy-datum
       hora_chg_st1 = sy-uzeit
       user_chg_st1 = sy-uname
       WHERE processo = header-processo AND
       ano = header-ano AND
       seqno = header-seqno.

    ENDIF.

    MOVE-CORRESPONDING header TO cab.

    LOOP AT item.
      MOVE-CORRESPONDING item TO linha.
      APPEND linha.
    ENDLOOP.


  ENDIF.

  COMMIT WORK AND WAIT.

  refresh = 'X'.

  CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
    TABLES
      i_message_tab = msg_cockpit.


ENDFUNCTION.
