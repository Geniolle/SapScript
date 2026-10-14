FUNCTION ZFI_CHANGE_BANKID.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(I_BANK_ACCOUNT_TYPE) TYPE  FIAPPT_BANK_ACC_TYPE
*"  CHANGING
*"     REFERENCE(CT_FIAPPTD_TRANS) TYPE  FIAPPT_T_TRANS
*"     REFERENCE(CT_XML_ITEMS) TYPE  FIAPPT_T_CBR_XML_ITEMS
*"----------------------------------------------------------------------

" **********************************************************************
" USER: NSO - Nuno Santos Oliveira - Abaco Consulting
" DATE: 16.06.2023 10:41:53
" DESCRIPTION: Clear bank ID for "fictitious" banks (non real banks)
" Function module used in the implicit enhancements 'ZFI_ENH_CHANGE_BANK_ID'
" and 'ZFI_ENH_CHANGE_BANK_ID_2'.
" Used in:
"   - Class/Method: CL_FIAPPT_CBR/PROCESS_PAY_WO_INV
"   - Class/Method: CL_FIAPPT_CBR/PROCESS_AUTO_PAYMENTS
" **********************************************************************
" {

  DATA: LR_BANK_ID TYPE RANGE OF FIAPPT_BANKID.

  FIELD-SYMBOLS: <FS_FIAPPTD_TRANS> TYPE FIAPPTD_TRANS,
                 <FS_XML_ITEMS> TYPE FIAPPT_S_CBR_XML_ITEMS.

  " Check bank account type
  IF I_BANK_ACCOUNT_TYPE = 'I'.

    " Get bank IDs to clear
    CLEAR LR_BANK_ID.
    ZCLCA_FIXEDVALS=>GET_CONS_RAN(
      EXPORTING
        IV_BUKRS = SPACE
        IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_FI
        IV_PROCE = 'COPE_CLEAR_BANK_ID'
        IV_FNAME = 'BANK_ID'
      IMPORTING
        ET_RANGE = LR_BANK_ID
      EXCEPTIONS
        NO_DATA  = 1  ).

    IF LR_BANK_ID[] IS NOT INITIAL.

      " Get all bank IDs to clear
      LOOP AT CT_FIAPPTD_TRANS ASSIGNING <FS_FIAPPTD_TRANS> WHERE BANK_ID IN LR_BANK_ID.

        " Clear bank ID
        CLEAR <FS_FIAPPTD_TRANS>-BANK_ID.

      ENDLOOP.

      " Get all bank IDs to clear
      LOOP AT CT_XML_ITEMS ASSIGNING <FS_XML_ITEMS> WHERE BANK_ID IN LR_BANK_ID.

        " Clear bank ID
        CLEAR <FS_XML_ITEMS>-BANK_ID.

      ENDLOOP.

    ENDIF.

  ENDIF.

" }
" **********************************************************************

ENDFUNCTION.