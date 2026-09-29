FUNCTION /sbxc/zckp_guarda_alt_adianta.
*"--------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(CAB)
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  CHANGING
*"     REFERENCE(LINHA)
*"--------------------------------------------------------------------
* Guarda Alterações na tabela de adiantamentos

  MOVE-CORRESPONDING cab TO /sbxc/zckp_adnt.

  UPDATE /sbxc/zckp_adnt.


ENDFUNCTION.
