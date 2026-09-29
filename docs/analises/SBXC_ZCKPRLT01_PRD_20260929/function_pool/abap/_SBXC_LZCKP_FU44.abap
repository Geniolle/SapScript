FUNCTION /sbxc/zckp_dm_tudo.
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

  wa_status-box = ''.
  MODIFY status FROM wa_status TRANSPORTING box
  WHERE box = 'X'.

  REFRESH linha.

  refresh = 'D'.


ENDFUNCTION.
