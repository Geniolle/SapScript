FUNCTION /sbxc/zckp_canc_liga_ulterior.
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
  DATA: l_status  TYPE /sbxc/zckp_ctrl-status1,
        l_doc     TYPE rbkp-belnr.

* DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: lv_errortext   TYPE string.

  REFRESH msg_cockpit.


* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
    LOOP AT linha.
      MOVE-CORRESPONDING linha TO item.
      APPEND item.
    ENDLOOP.

* verifica se a associação tinha sido feita de forma manual - so neste caso podemos eliminar
* a associação
    SELECT SINGLE status2 FROM /sbxc/zckp_ctrl INTO /sbxc/zckp_ctrl-status2
      WHERE processo = header-processo AND
             ano = header-ano AND
             seqno = header-seqno AND
             status2 = 'M'.

    IF sy-subrc <> 0.
      CALL FUNCTION 'POPUP_TO_INFORM'
        EXPORTING
          titel         = text-053 "'Processo Inválido'
          txt1          = text-054 "'O processo selecionado não pode ser alterado '
          txt2          = text-055 "'Associação não foi efetuada de forma manual'
*         TXT3          = ' '
*         TXT4          = ' '
                .
    ELSE.



      CALL FUNCTION 'POPUP_TO_CONFIRM'
        EXPORTING
          text_question         = text-056 "'Pretende eliminar a associação feita anteriormente?'
          text_button_1         = 'Sim'(001)
          display_cancel_button = ' '
        IMPORTING
          answer                = ret
        EXCEPTIONS
          text_not_found        = 1
          OTHERS                = 2.

      CHECK ret = '1'.


        UPDATE /sbxc/zckp_invh
           SET doc_fi   = ' '
               doc_lo   = ' '
               ano_lanc = ' '
               doc_estorno = ' '
               data_criacao = ' '
         WHERE processo = header-processo AND
               ano = header-ano AND
               seqno = header-seqno.

        COMMIT WORK.

        UPDATE /sbxc/zckp_ctrl SET
           status1 = '0'
           status2 = ' ' "Actualizado manualmente
           WHERE processo = header-processo AND
            ano = header-ano AND
            seqno = header-seqno.

        COMMIT WORK.

*        "Envia reversão para workflow
*      IF header-doc_estorno IS NOT INITIAL.
*
*            CREATE OBJECT lo_docinf.
*
*            TRY.
*                lo_docinf->reverse_bim_process( exporting i_zsckp_to_bim = ls_reverse
*                                                          doc_estorno = header-doc_estorno ).
**
***              CATCH cx_ai_system_fault INTO lo_systemfault.
***                lv_errortext = lo_systemfault->errortext.
**
***           header-erro = lv_errortext.
***
***           UPDATE /sbxc/zckp_invh SET erro = header-erro
***           WHERE processo EQ header-processo
***            AND ano EQ header-ano
***            AND seqno EQ header-seqno.
***           commit WORK AND WAIT.
*            ENDTRY.
*
*      ENDIF.
*  "Fim envio de reversão para workflow

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


ENDFUNCTION.
