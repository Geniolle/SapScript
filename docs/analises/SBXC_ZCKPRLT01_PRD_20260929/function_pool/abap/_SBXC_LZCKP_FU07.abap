FUNCTION /sbxc/zckp_copia.
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

  DATA: item     TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.

  DATA: l_returncode,
        lt_sval LIKE sval OCCURS 0 WITH HEADER LINE.

  REFRESH msg_cockpit.


  DATA erro(1).

  DATA ret TYPE TABLE OF bapiret2.

  DATA: processo_novo LIKE header-processo,
        seqno_novo LIKE header-seqno,
        ano_novo LIKE header-ano.

*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.

  IF ctrl-status1 = '0'.

* se move processo de tax para processo normal então tem de actualizar campo estado
*    IF header-status IS INITIAL.
*      MESSAGE e008(/sbxc/zckp_cockpit).
**   Preencher tipo de condição de espera - Campo Estado
*
*      erro = 'X'.
**   Preencher tipo de condição de espera - Campo Status
*    ELSE.

    APPEND header.
*    ENDIF.
  ELSE.
    MESSAGE i009(/sbxc/zckp_cockpit).
    erro = 'X'.
  ENDIF.
*  ENDLOOP.

*  CHECK erro IS INITIAL.
IF erro IS INITIAL.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.


  LOOP AT header.

    CALL FUNCTION 'NUMBER_GET_NEXT'
     EXPORTING
       nr_range_nr                   = '01'
       object                        = 'ZCKP_COCKP'
      quantity                      = '1'
*      SUBOBJECT                     = ' '
      toyear                        = sy-datum(4)
*      IGNORE_BUFFER                 = ' '
    IMPORTING
      number                        = seqno_novo
*      QUANTITY                      =
*      RETURNCODE                    =
*    EXCEPTIONS
*      INTERVAL_NOT_FOUND            = 1
*      NUMBER_RANGE_NOT_INTERN       = 2
*      OBJECT_NOT_FOUND              = 3
*      QUANTITY_IS_0                 = 4
*      OTHERS                        = 5
  .

    CALL FUNCTION '/SBXC/ZCKP_DETERMINA_PROCESSO'
     EXPORTING
       headerdata       =  header
       tax              = ' '
    IMPORTING
*     REFRESH          =
      processo         =  processo_novo
     TABLES
       itemdata         = item
       return           = ret.


    UPDATE /sbxc/zckp_ctrl SET
       status1 = 'T'
       processo_post = processo_novo
       ano_post = sy-datum(4)
       seqno_post = seqno_novo
       WHERE processo = header-processo AND
       ano = header-ano AND
       seqno = header-seqno.

    /sbxc/zckp_ctrl-processo = processo_novo.
    /sbxc/zckp_ctrl-seqno = seqno_novo.
    /sbxc/zckp_ctrl-ano = sy-datum(4).
    /sbxc/zckp_ctrl-processo_ant = header-processo.
    /sbxc/zckp_ctrl-ano_ant = header-ano.
    /sbxc/zckp_ctrl-seqno_ant = header-seqno.
    /sbxc/zckp_ctrl-data_in = sy-datum.
    /sbxc/zckp_ctrl-hora_in = sy-uzeit.
    INSERT /sbxc/zckp_ctrl.



    MOVE-CORRESPONDING header TO /sbxc/zckp_invh.
    MOVE-CORRESPONDING /sbxc/zckp_ctrl TO /sbxc/zckp_invh.
*    /sbxc/zckp_invh-status = header-status.
    INSERT  /sbxc/zckp_invh.


    LOOP AT item WHERE processo = header-processo AND
                       ano = header-ano AND
                       seqno = header-seqno.

      MOVE-CORRESPONDING item TO /sbxc/zckp_invi.
      MOVE-CORRESPONDING header TO /sbxc/zckp_invi.
      MOVE-CORRESPONDING /sbxc/zckp_ctrl TO /sbxc/zckp_invi.
      INSERT /sbxc/zckp_invi.

    ENDLOOP.

  ENDLOOP.

  COMMIT WORK AND WAIT.

  wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
  wa_msg-msgno = '007'.
  wa_msg-msgty = 'S'.
  wa_msg-msgv1 =  processo_novo.
  wa_msg-msgv2 = seqno_novo.
  APPEND wa_msg TO msg_cockpit.

  CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
    TABLES
      i_message_tab = msg_cockpit.

ENDIF.

ENDFUNCTION.
