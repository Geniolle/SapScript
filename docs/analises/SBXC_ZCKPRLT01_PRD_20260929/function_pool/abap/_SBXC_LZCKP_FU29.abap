function /sbxc/zckp_refresh.
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

  call function '/SBXC/ZCKP_ITEM_FAT'
*    EXPORTING
*      cab   = cab
      IMPORTING
         refresh = refresh
    tables
      linha = linha
      cor   = cor
    changing cab = cab .

*  refresh = 'X'.
endfunction.
