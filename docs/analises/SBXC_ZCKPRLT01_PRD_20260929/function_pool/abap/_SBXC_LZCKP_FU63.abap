FUNCTION /SBXC/ZCKP_REJ_MULTI.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) LIKE  SY-UCOMM
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA OPTIONAL
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"      CAB
*"      CTRL STRUCTURE  /SBXC/ZCKP_CTRL
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        wa_header TYPE /sbxc/zckp_invh,
        item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        l_returncode,
        lt_sval LIKE sval OCCURS 0 WITH HEADER LINE,
        lv_pro  TYPE  /sbxc/zckp_processo,
        lv_ano  TYPE  gjahr,
        lv_seq  TYPE  /sbxc/zckp_seqno,
        opcao TYPE i,
        est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem,
        processo_pos TYPE /sbxc/zckp_processo,
        ano_pos TYPE gjahr,
        seqno_pos TYPE /sbxc/zckp_seqno,
        gt_outtab TYPE TABLE OF /sbxc/zckp_tab07 WITH HEADER LINE,
        gs_private TYPE slis_data_caller_exit,
        gs_selfield TYPE slis_selfield,
        g_exit(1) TYPE c.


  DATA: BEGIN OF mensagens OCCURS 0,
          cod_mes TYPE /sbxc/zckp_tab06-cod_mes,
          split,
          mensagem TYPE /sbxc/zckp_tab07-mensagem,
        END OF mensagens.

  REFRESH msg_cockpit.
* Preencher estrutura cabecalho


  loop at cab.

  MOVE-CORRESPONDING cab TO header.
  APPEND header.

  ENDLOOP.

  IF  ctrl-status1 = '9' or  ctrl-status1 = '3' or  ctrl-status1 = '4'.
    MESSAGE s026(/sbxc/zckp_cockpit).
  else.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.

* Ler ref1 na ekko
    IF item-po_number NE space.
      SELECT SINGLE ihrez INTO item-ref_1
        FROM ekko
        WHERE ebeln EQ item-po_number.

      CLEAR:  item-ref_doc,
              item-ref_doc_year,
              item-ref_doc_item.
    ENDIF.
    APPEND item.
  ENDLOOP.

  READ TABLE header INDEX 1.
**  select * from /SBXC/ZCKP_INVI
**      appending corresponding fields of table item
**      where processo = header-processo
**          and ano = header-ano
**          and seqno = header-seqno.

* Verificar se processos das mensagens seleccionadas têm o mesmo grupo
* de mensagens de rejeição
  est_mensagem = space.
  LOOP AT header.

    SELECT SINGLE est_mensagem INTO /sbxc/zckp_tab00-est_mensagem
    FROM /sbxc/zckp_tab00 WHERE processo = header-processo.

    AT FIRST.
      est_mensagem =  /sbxc/zckp_tab00-est_mensagem.
    ENDAT.

    IF NOT est_mensagem =  /sbxc/zckp_tab00-est_mensagem.
      MESSAGE e004(/sbxc/zckp_cockpit).
*   Seleccionar registos com a mesma estrutura de mensagens
    ENDIF.

  ENDLOOP.


  SELECT  * FROM  /sbxc/zckp_tab06
         WHERE  est_mensagem  = est_mensagem.

    SELECT SINGLE * FROM  /sbxc/zckp_tab07 WHERE
                  est_mensagem  = /sbxc/zckp_tab06-est_mensagem
           AND    cod_mes   = /sbxc/zckp_tab06-cod_mes
           AND    spras     = sy-langu.
    mensagens-mensagem = /sbxc/zckp_tab07-mensagem.
    mensagens-cod_mes = /sbxc/zckp_tab07-cod_mes.
    MOVE-CORRESPONDING /sbxc/zckp_tab07 TO gt_outtab.
    APPEND gt_outtab.
    APPEND mensagens.
  ENDSELECT.


  CLEAR opcao.
*  CHECK NOT mensagens[] IS INITIAL.
*  SORT mensagens BY cod_mes.
*  CALL FUNCTION 'POPUP_WITH_TABLE_DISPLAY'
*    EXPORTING
*      endpos_col   = 50
*      endpos_row   = 20
*      startpos_col = 1
*      startpos_row = 1
*      titletext    = 'Motivo rejeição'
*    IMPORTING
*      choise       = opcao
*    TABLES
*      valuetab     = mensagens
*    EXCEPTIONS
*      break_off    = 1
*      OTHERS       = 2.
*  IF sy-subrc <> 0.
** MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
**         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
*  ENDIF.


  CALL FUNCTION 'REUSE_ALV_POPUP_TO_SELECT'
    EXPORTING
      i_title          = text-051 "'Motivo rejeição'
      i_tabname        = '1'
      i_structure_name = '/SBXC/ZCKP_TAB07'
      is_private       = gs_private
    IMPORTING
      es_selfield      = gs_selfield
      e_exit           = g_exit
    TABLES
      t_outtab         = gt_outtab
    EXCEPTIONS
      program_error    = 1
      OTHERS           = 2.


  IF g_exit NE 'X'.

    opcao = gs_selfield-tabindex.

* Ler parametrização da mensagem seleccionada
    READ TABLE mensagens INDEX opcao.
* Verificar se todas as linhas seleccionadas
    SELECT SINGLE *
      FROM /sbxc/zckp_tab06 WHERE est_mensagem = est_mensagem AND
                            cod_mes = mensagens-cod_mes.


* Preencher estrutura cabecalho
    REFRESH:  linha.
    LOOP AT header.

      refresh return.
      clear return.

      /sbxc/zckp_tab09-processo = header-processo.
      /sbxc/zckp_tab09-ano = header-ano.
      /sbxc/zckp_tab09-seqno = header-seqno.
      /sbxc/zckp_tab09-est_mensagem = /sbxc/zckp_tab06-est_mensagem.
      /sbxc/zckp_tab09-cod_mes = /sbxc/zckp_tab06-cod_mes.
      CONCATENATE /sbxc/zckp_tab06-prefixo mensagens-mensagem INTO /sbxc/zckp_tab09-mensagem.
      /sbxc/zckp_tab09-processo_seg = /sbxc/zckp_tab06-destino.
      INSERT /sbxc/zckp_tab09.


      UPDATE /sbxc/zckp_ctrl SET
      status1 = '6'
      data_chg_st1 = sy-datum
      hora_chg_st1 = sy-uzeit
      user_chg_st1 = sy-uname
      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

      ctrl-status1 = '6'.
*   Se mudar de processo
      IF /sbxc/zckp_tab06-destino NE space AND
        /sbxc/zckp_tab06-destino <> header-processo.

        lv_pro = header-processo.
        lv_ano = header-ano.
        lv_seq = header-seqno.

        header-processo = /sbxc/zckp_tab06-destino.


        CALL FUNCTION '/SBXC/ZCKP_MM_CRIA_FACTURAS'
          EXPORTING
            headerdata   = header
            processo_ant = lv_pro
            ano_ant      = lv_ano
            seqno_ant    = lv_seq
          TABLES
            itemdata     = item
            return       = return.

        READ TABLE return INDEX 1.
        SPLIT return-message AT '/' INTO processo_pos ano_pos seqno_pos.



        UPDATE /sbxc/zckp_ctrl SET
        processo_post  = processo_pos
        ano_post       = ano_pos
        seqno_post     = seqno_pos
*        status1 = '6'
        WHERE processo = lv_pro AND
        ano            = lv_ano AND
        seqno          = lv_seq.

        UPDATE /sbxc/zckp_ctrl
        SET
        processo_ant  = lv_pro
        ano_ant       = lv_ano
        seqno_ant     = lv_seq
*        status1 =     ' '
        WHERE
        processo = processo_pos AND
        ano            = ano_pos AND
        seqno          = seqno_pos.
      ENDIF.


      MOVE-CORRESPONDING header TO cab.

    ENDLOOP.


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

endif.
ENDFUNCTION.
