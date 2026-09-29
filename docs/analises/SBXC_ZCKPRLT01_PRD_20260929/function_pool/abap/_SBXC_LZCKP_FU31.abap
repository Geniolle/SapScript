FUNCTION /sbxc/zckp_ver_doc_fi.
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

  REFRESH msg_cockpit.

* Preencher estrutura cabecalho
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
*  ENDLOOP.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  LOOP AT header.
    CLEAR status_ok.
*    PERFORM valida_status USING '34'
*                                header-processo
*                                header-ano
*                                header-seqno
*                                ctrl-status1
*                      CHANGING status_ok.
*
*    CHECK status_ok = 'X'.

* Verificar se documento existe
    SET PARAMETER ID: 'BLN' FIELD header-doc_fi,
                      'BUK' FIELD header-comp_code,
                      'GJR' FIELD header-ano_lanc.


    SELECT SINGLE bukrs INTO bkpf-bukrs FROM bkpf
     WHERE bukrs = header-comp_code     AND
           belnr = header-doc_fi  AND
           gjahr = header-ano_lanc.


    IF sy-subrc <> 0.
* documento não encontrado
      wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
      wa_msg-msgno = '003'.
      wa_msg-msgty = 'W'.
      wa_msg-msgv1 =  header-comp_code .
      wa_msg-msgv1 =  header-doc_fi  .
      wa_msg-msgv1 =  header-ano_lanc .
      APPEND wa_msg TO msg_cockpit .

    ELSE.

      CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.


    ENDIF.

  ENDLOOP.


  LOOP AT header.
    indice = sy-tabix.
    REFRESH: item_f, it_return.

    MOVE-CORRESPONDING header TO wa_header.

    IF wa_header-doc_fi <> ''.

      REFRESH: cor_tab[], cor[].
      cor[] = cor_tab[].
    ENDIF.

  ENDLOOP.


* Preencher estrutura cabecalho
  REFRESH: linha.
  CLEAR cab.
  LOOP AT header.
    MOVE-CORRESPONDING header TO cab.
    MOVE-CORRESPONDING ctrl to cab. "ODC - 04_11_2020
*    APPEND cab.
  ENDLOOP.


  LOOP AT item.
    MOVE-CORRESPONDING item TO linha.
    APPEND linha.
  ENDLOOP.

  CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
    TABLES
      i_message_tab = msg_cockpit.

ENDFUNCTION.
