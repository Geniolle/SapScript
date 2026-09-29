FUNCTION /sbxc/zckp_m_tudo.
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

  wa_status-box = 'X'.
  MODIFY status FROM wa_status TRANSPORTING box
  WHERE box IS INITIAL.

  refresh = 'X'.


ENDFUNCTION.
