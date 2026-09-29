FUNCTION /sbxc/zckp_rej_pos.
*"--------------------------------------------------------------------
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
*"--------------------------------------------------------------------

  TYPES: BEGIN OF ty_rej_msg,
          processo TYPE /sbxc/zckp_tab09-processo,
          mensagem TYPE /sbxc/zckp_tab09-mensagem,
          cod_mes  TYPE /sbxc/zckp_tab09-cod_mes,
          resposta TYPE char1,
         END OF ty_rej_msg.

  DATA: it_rej_msg TYPE STANDARD TABLE OF ty_rej_msg.

  IMPORT it_rej_msg FROM MEMORY ID 'CKP_REJ'.

  REFRESH it_rej_msg.

  EXPORT it_rej_msg TO MEMORY ID 'CKP_REJ'.



ENDFUNCTION.
