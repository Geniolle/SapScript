FUNCTION ZFI_CHANGE_ACC_TYPE.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  CHANGING
*"     REFERENCE(ET_FIAPPTD_BALANCE) TYPE  FIAPPT_T_BALANCE
*"----------------------------------------------------------------------

LOOP AT ET_FIAPPTD_BALANCE ASSIGNING FIELD-SYMBOL(<FS>) WHERE BANK_ACC_TYPE <> 'X'.
  <FS>-BANK_ACC_TYPE = 'X'.
ENDLOOP.

ENDFUNCTION.