FUNCTION /sbxc/zckp_inverter.
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
*"      STATUS STRUCTURE  /SBXC/ZCKP_TAB_STATUS OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------

  DATA: wa_status TYPE /sbxc/zckp_tab_status.

  LOOP AT status INTO wa_status.
    IF wa_status-box = 'X'.
      CLEAR wa_status-box.
    ELSEIF wa_status-box IS INITIAL.
      wa_status-box = 'X'.
    ENDIF.
    MODIFY status FROM wa_status INDEX sy-tabix.
  ENDLOOP.

  REFRESH linha.
  refresh = 'I'.


ENDFUNCTION.
