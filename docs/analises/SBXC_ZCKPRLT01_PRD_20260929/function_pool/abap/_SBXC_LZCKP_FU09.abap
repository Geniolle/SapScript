FUNCTION /SBXC/ZCKP_LOG_MESS .
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
  data: header type table of /sbxc/zckp_invh with header line.
  data: wa_header type /sbxc/zckp_invh.


* Preencher estrutura cabecalho

*  LOOP AT cab.
  move-corresponding cab to header.
  append header.
*  ENDLOOP.


  refresh msg_cockpit.
  loop at header.


    select * from /sbxc/zckp_tab08 where
         processo = header-processo and
         ano = header-ano and
         seqno = header-seqno.
      wa_msg-msgid = /sbxc/zckp_tab08-msgid.
      wa_msg-msgty = /sbxc/zckp_tab08-msgty.
      wa_msg-msgno = /sbxc/zckp_tab08-msgno.
      wa_msg-msgv1 = /sbxc/zckp_tab08-msgv1.
      wa_msg-msgv2 = /sbxc/zckp_tab08-msgv2.
      wa_msg-msgv3 = /sbxc/zckp_tab08-msgv3.
      wa_msg-msgv4 = /sbxc/zckp_tab08-msgv4.
      append wa_msg to msg_cockpit .


    endselect.
  endloop.

  call function 'C14Z_MESSAGES_SHOW_AS_POPUP'
    tables
      i_message_tab = msg_cockpit.


endfunction.
