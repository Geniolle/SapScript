FUNCTION /sbxc/zckp_guarda_alt_trat.
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


* Preencher estrutura cabecalho
  refresh = 'X'.
*  LOOP AT cab.
*    MOVE-CORRESPONDING cab TO zckp_inv_header.
  MOVE-CORRESPONDING cab TO /sbxc/zckp_invh.

  IF e_ucomm = 'TRAT_A'.
    IF /sbxc/zckp_invh-user_tratamento IS INITIAL.
    /sbxc/zckp_invh-user_tratamento = sy-uname.
    /sbxc/zckp_invh-data_tratamento = sy-datum.
    /sbxc/zckp_invh-em_tratamento = 'X'.
    ELSE.
      MESSAGE s036(/sbxc/zckp_cockpit) WITH /sbxc/zckp_invh-user_tratamento.
    ENDIF.
  ELSE.
    IF /sbxc/zckp_invh-user_tratamento = sy-uname.
    CLEAR: /sbxc/zckp_invh-user_tratamento,
    /sbxc/zckp_invh-data_tratamento,
    /sbxc/zckp_invh-em_tratamento.
    ELSE.
      MESSAGE s036(/sbxc/zckp_cockpit) WITH /sbxc/zckp_invh-user_tratamento.
    ENDIF.
  ENDIF.
  MOVE-CORRESPONDING /sbxc/zckp_invh TO cab.


  UPDATE /sbxc/zckp_invh.

  COMMIT WORK AND WAIT.

ENDFUNCTION.
