FUNCTION /sbxc/zckp_rej.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) LIKE  SY-UCOMM
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
  DATA: header       TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        wa_header    TYPE /sbxc/zckp_invh,
        item         TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        item_f       TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        l_returncode,
        lt_sval      LIKE sval OCCURS 0 WITH HEADER LINE,
        lv_pro       TYPE  /sbxc/zckp_processo,
        lv_ano       TYPE  gjahr,
        lv_seq       TYPE  /sbxc/zckp_seqno,
        opcao        TYPE i,
        est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem,
        processo_pos TYPE /sbxc/zckp_processo,
        ano_pos      TYPE gjahr,
        seqno_pos    TYPE /sbxc/zckp_seqno,
        gt_outtab    TYPE TABLE OF /sbxc/zckp_tab07 WITH HEADER LINE,
        gs_private   TYPE slis_data_caller_exit,
        gs_selfield  TYPE slis_selfield,
        g_exit(1)    TYPE c.


  DATA: BEGIN OF mensagens OCCURS 0,
          cod_mes  TYPE /sbxc/zckp_tab06-cod_mes,
          split,
          mensagem TYPE /sbxc/zckp_tab07-mensagem,
        END OF mensagens.

  TYPES: BEGIN OF ty_rej_msg,
           processo TYPE /sbxc/zckp_tab09-processo,
           mensagem TYPE /sbxc/zckp_tab09-mensagem,
           cod_mes  TYPE /sbxc/zckp_tab09-cod_mes,
           resposta TYPE char1,
         END OF ty_rej_msg.

  DATA: it_rej_msg TYPE STANDARD TABLE OF ty_rej_msg,
        wa_rej_msg TYPE ty_rej_msg,
        lv_msg_idx TYPE sy-tabix.

  DATA: lv_spras TYPE lfa1-spras.
  DATA: n_lines TYPE i.

  REFRESH msg_cockpit.
* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.

  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    "Processo em tratamento pelo utilizador &
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
IF  ctrl-status1 = '9' OR  ctrl-status1 = '3' OR  ctrl-status1 = '4' OR ctrl-status1 = '6'.
      "Status do documento não permite esta operação
      MESSAGE s026(/sbxc/zckp_cockpit).
    ELSE.

      LOOP AT linha.
        MOVE-CORRESPONDING linha TO item.
* Ler ref1 na ekko
        IF item-po_number NE space.
          IF ITEM-REF_DOC IS INITIAL. "ODC - 30_11_2020

          SELECT SINGLE ihrez INTO item-ref_1
            FROM ekko
            WHERE ebeln EQ item-po_number.

          CLEAR:  item-ref_doc,
                  item-ref_doc_year,
                  item-ref_doc_item.

          ENDIF.
        ENDIF.
        APPEND item.
      ENDLOOP.
      READ TABLE header INDEX 1.

      "Selecionar idioma do fornecedor
      CLEAR: lv_spras.
      SELECT SINGLE spras INTO lv_spras
        FROM lfa1
        WHERE lifnr EQ header-vendor.
      IF lv_spras IS INITIAL.
        lv_spras = 'P'.
      ELSEIF lv_spras NE 'P'.
        lv_spras = 'E'.
      ENDIF.
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
          EXIT.
*   Seleccionar registos com a mesma estrutura de mensagens
        ENDIF.
      ENDLOOP.
      CLEAR n_lines.
      IMPORT n_lines TO n_lines FROM MEMORY ID 'N_LINES'.
      REFRESH it_rej_msg.
      IMPORT it_rej_msg FROM MEMORY ID 'CKP_REJ'.

      READ TABLE it_rej_msg INTO wa_rej_msg
                            WITH KEY processo = header-processo
                                     resposta = 'X'.
      lv_msg_idx = sy-tabix.
      IF sy-subrc NE 0.
        IF v_cod_mes IS NOT INITIAL AND v_mensagem IS NOT INITIAL.
          gt_outtab-est_mensagem = v_cod_mes.
          gt_outtab-cod_mes = v_mensagem.
        ELSE.
          IF e_ucomm EQ 'REJ2'.
            SELECT  * FROM  /sbxc/zckp_tab06
                 WHERE  est_mensagem  = est_mensagem
                  AND destino NE 'SAPHETY'.

              SELECT SINGLE * FROM  /sbxc/zckp_tab07 WHERE
                           est_mensagem  = /sbxc/zckp_tab06-est_mensagem
                    AND    cod_mes   = /sbxc/zckp_tab06-cod_mes
                    AND    spras     = lv_spras. "sy-langu.
              mensagens-mensagem = /sbxc/zckp_tab07-mensagem.
              mensagens-cod_mes = /sbxc/zckp_tab07-cod_mes.
              MOVE-CORRESPONDING /sbxc/zckp_tab07 TO gt_outtab.
              APPEND gt_outtab.
              APPEND mensagens.
            ENDSELECT.
          ELSE.
            SELECT  * FROM  /sbxc/zckp_tab06
                   WHERE  est_mensagem  = est_mensagem
                    AND destino EQ 'SAPHETY'.

              SELECT SINGLE * FROM  /sbxc/zckp_tab07 WHERE
                           est_mensagem  = /sbxc/zckp_tab06-est_mensagem
                    AND    cod_mes   = /sbxc/zckp_tab06-cod_mes
                    AND    spras     = lv_spras. "sy-langu.
              mensagens-mensagem = /sbxc/zckp_tab07-mensagem.
              mensagens-cod_mes = /sbxc/zckp_tab07-cod_mes.
              MOVE-CORRESPONDING /sbxc/zckp_tab07 TO gt_outtab.
              APPEND gt_outtab.
              APPEND mensagens.
            ENDSELECT.
          ENDIF.
          CLEAR opcao.
          SORT gt_outtab BY est_mensagem cod_mes.
          CALL FUNCTION 'REUSE_ALV_POPUP_TO_SELECT'
            EXPORTING
              i_title          = TEXT-029 "'Motivo rejeição'
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
        ENDIF.
      ENDIF.
      IF g_exit NE 'X'.

        SELECT SINGLE destino INTO processo_pos
          FROM /sbxc/zckp_tab06
          WHERE est_mensagem EQ gt_outtab-est_mensagem
            AND cod_mes EQ gt_outtab-cod_mes.
        IF processo_pos = 'SAPHETY'.

          IF header-mot_n_contab IS INITIAL.
            v_cod_mes = gt_outtab-est_mensagem.
            v_mensagem = gt_outtab-cod_mes.
*      "Inserir código da acção - campo "C.Mensag.'
            MESSAGE s039(/sbxc/zckp_cockpit) DISPLAY LIKE 'E'.
            RETURN.
          ELSE.
            IF v_cod_mes IS NOT INITIAL AND v_mensagem IS NOT INITIAL.
              gt_outtab-est_mensagem = v_cod_mes.
              gt_outtab-cod_mes = v_mensagem.
              SELECT SINGLE mensagem INTO gt_outtab-mensagem
                FROM /sbxc/zckp_tab07
                WHERE cod_mes EQ v_mensagem
                AND est_mensagem EQ /sbxc/zckp_tab06-est_mensagem
                AND spras EQ lv_spras.
              CLEAR :  v_cod_mes, v_mensagem.
            ENDIF.
          ENDIF.
        ENDIF.
        CLEAR processo_pos.
        CLEAR :  v_cod_mes, v_mensagem.
        IF wa_rej_msg-resposta NE 'X'.
          opcao = gs_selfield-tabindex.
* Ler parametrização da mensagem seleccionada
          READ TABLE gt_outtab INDEX opcao.

          IF gt_outtab-mensagem IS INITIAL.
            SELECT SINGLE mensagem INTO gt_outtab-mensagem
                 FROM /sbxc/zckp_tab07
                 WHERE cod_mes EQ gt_outtab-cod_mes
                 AND est_mensagem EQ /sbxc/zckp_tab06-est_mensagem
                 AND spras EQ lv_spras.
          ENDIF.

          wa_rej_msg-processo = header-processo.
          wa_rej_msg-mensagem = gt_outtab-mensagem.
          wa_rej_msg-cod_mes  = gt_outtab-cod_mes.
          mensagens-cod_mes  = wa_rej_msg-cod_mes.
          mensagens-mensagem = wa_rej_msg-mensagem.
          DATA: lv_answer(1).

          IF n_lines > 1.

            CALL FUNCTION 'POPUP_TO_CONFIRM'
              EXPORTING
*              TITLEBAR       = ' '
*              DIAGNOSE_OBJECT             = ' '
              text_question  = TEXT-030 "'Aplicar o mesmo motivo a todos os registos do processo?'
              text_button_1  = 'Sim'(001)
*              ICON_BUTTON_1  = ' '
              text_button_2  = 'Não'(002)
              IMPORTING
              answer         = lv_answer
              EXCEPTIONS
              text_not_found = 1
              OTHERS         = 2.
            IF sy-subrc EQ 0 AND lv_answer EQ '1'.
              wa_rej_msg-resposta = 'X'.
            ENDIF.
          ENDIF.
          CHECK lv_answer NE 'A'.
            IF lv_msg_idx GT 0.
              MODIFY it_rej_msg FROM wa_rej_msg INDEX lv_msg_idx.
            ELSE.
              APPEND wa_rej_msg TO it_rej_msg .
            ENDIF.

        ELSE.
          mensagens-cod_mes  = wa_rej_msg-cod_mes.
          mensagens-mensagem = wa_rej_msg-mensagem.
        ENDIF.
* Verificar se todas as linhas seleccionadas
        SELECT SINGLE *
          FROM /sbxc/zckp_tab06 WHERE est_mensagem = est_mensagem AND
                                cod_mes = mensagens-cod_mes.

* Preencher estrutura cabecalho
        REFRESH:  linha.
        LOOP AT header.

          REFRESH return.
          CLEAR return.
          IF /sbxc/zckp_tab06-destino NE space AND
                /sbxc/zckp_tab06-destino <> header-processo.

            /sbxc/zckp_tab09-processo = header-processo.
            /sbxc/zckp_tab09-ano = header-ano.
            /sbxc/zckp_tab09-seqno = header-seqno.
          /sbxc/zckp_tab09-est_mensagem = /sbxc/zckp_tab06-est_mensagem.
            /sbxc/zckp_tab09-cod_mes = /sbxc/zckp_tab06-cod_mes.
CONCATENATE /sbxc/zckp_tab06-prefixo mensagens-mensagem INTO /sbxc/zckp_tab09-mensagem.
            /sbxc/zckp_tab09-processo_seg = /sbxc/zckp_tab06-destino.
*            /sbxc/zckp_tab09-data_dev = header-data_dev.
            INSERT /sbxc/zckp_tab09.
            header-mensagem = /sbxc/zckp_tab09-mensagem.


            UPDATE /sbxc/zckp_ctrl SET
            status1 = '6'
            data_chg_st1 = sy-datum
            hora_chg_st1 = sy-uzeit
            user_chg_st1 = sy-uname
            WHERE processo = header-processo AND
            ano = header-ano AND
            seqno = header-seqno.

            UPDATE /sbxc/zckp_invh SET
                  mot_n_contab = header-mot_n_contab
                  mensagem     = /sbxc/zckp_tab09-mensagem
            WHERE processo = header-processo AND
                  ano = header-ano AND
                  seqno = header-seqno.

            ctrl-status1 = '6'.
            lv_pro = header-processo.
            lv_ano = header-ano.
            lv_seq = header-seqno.

  IF /sbxc/zckp_tab06-destino = 'LO' OR /sbxc/zckp_tab06-destino = 'FI'.
              CLEAR header-processo.
            ELSE.
              header-processo = /sbxc/zckp_tab06-destino.
            ENDIF.

            CALL FUNCTION '/SBXC/ZCKP_MM_CRIA_FACTURAS'
              EXPORTING
                headerdata   = header
                processo_ant = lv_pro
                ano_ant      = lv_ano
                seqno_ant    = lv_seq
              TABLES
                itemdata     = item
                return       = return.

*  se /sbxc/zckp_tab06-send_mail = 'X', enviar email para o fornecedor
            IF /sbxc/zckp_tab06-send_mail EQ 'X'.
*             Enviar email para o fornecedor
              CALL FUNCTION '/SBXC/ZCKP_SENDMAIL_FORN'
                EXPORTING
                  i_bukrs      = header-comp_code
                  i_ref_doc_no = header-ref_doc_no
                  i_gjahr      = header-ano
                  i_vendor     = header-vendor
                  i_doc_date   = header-doc_date
                  i_pstng_date = header-pstng_date
                  i_bcc        = 'X'
                  i_reject     = 'X'
                  i_est_mens   = /sbxc/zckp_tab06-est_mensagem
                  i_cod_mens   = /sbxc/zckp_tab06-cod_mes.
            ENDIF.
*  se /sbxc/zckp_tab06-send_mail = 'X', enviar email para o fornecedor

            READ TABLE return INDEX 1.
        SPLIT return-message AT '\' INTO processo_pos ano_pos seqno_pos.

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
*            ENDIF.

            MOVE-CORRESPONDING header TO cab.
          ELSE.

IF /sbxc/zckp_tab06-send_mail EQ 'X' AND /sbxc/zckp_tab06-destino EQ space.
*             Enviar email para o fornecedor
              CALL FUNCTION '/SBXC/ZCKP_SENDMAIL_FORN'
                EXPORTING
                  i_bukrs      = header-comp_code
                  i_ref_doc_no = header-ref_doc_no
                  i_gjahr      = header-ano
                  i_vendor     = header-vendor
                  i_doc_date   = header-doc_date
                  i_pstng_date = header-pstng_date
                  i_bcc        = 'X'
                  i_reject     = 'X'
                  i_est_mens   = /sbxc/zckp_tab06-est_mensagem
                  i_cod_mens   = /sbxc/zckp_tab06-cod_mes.
            ELSE.

              "ERRO: Processo de destino idêntico ao processo de origem
              MESSAGE s051(/sbxc/zckp_cockpit) .
            ENDIF.

          ENDIF.
* Envia informação para PI para posteriro envio para Saphety
          IF       processo_pos = 'SAPHETY'.
            CLEAR :  v_cod_mes, v_mensagem.

            CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
              EXPORTING
                cab     = header
                cod_mes = /sbxc/zckp_tab06-cod_mes
                status  = 'REJECTED'
                accao   = header-mot_n_contab.

          ENDIF.
        ENDLOOP.
        LOOP AT item.
          MOVE-CORRESPONDING item TO linha.
          APPEND linha.
        ENDLOOP.
      ENDIF.
      COMMIT WORK AND WAIT.

      EXPORT it_rej_msg TO MEMORY ID 'CKP_REJ'.

      refresh = 'X'.

      CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
        TABLES
          i_message_tab = msg_cockpit.
*endif.
    ENDIF.
  ENDIF.

ENDFUNCTION.
