FUNCTION /SBXC/ZCKP_INV_BAPI_FI1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(DOCUMENTHEADER) TYPE  BAPIACHE03
*"     VALUE(CUSTOMERCPD) TYPE  BAPIACPA00 OPTIONAL
*"  EXPORTING
*"     VALUE(OBJ_TYPE) TYPE  BAPIACHE03-OBJ_TYPE
*"     VALUE(OBJ_KEY) TYPE  BAPIACHE03-OBJ_KEY
*"     VALUE(OBJ_SYS) TYPE  BAPIACHE03-OBJ_SYS
*"  TABLES
*"      ACCOUNTPAYABLE STRUCTURE  BAPIACAP03
*"      ACCOUNTGL STRUCTURE  BAPIACGL03
*"      ACCOUNTTAX STRUCTURE  BAPIACTX01
*"      CURRENCYAMOUNT STRUCTURE  BAPIACCR01
*"      PURCHASEORDER STRUCTURE  BAPIACPO00 OPTIONAL
*"      PURCHASEAMOUNT STRUCTURE  BAPIACCRPO OPTIONAL
*"      RETURN STRUCTURE  BAPIRET2
*"      CRITERIA STRUCTURE  BAPIACKECR OPTIONAL
*"      VALUEFIELD STRUCTURE  BAPIACKEVA OPTIONAL
*"      EXTENSION1 STRUCTURE  BAPIEXTC OPTIONAL
*"----------------------------------------------------------------------


* apenas para testes SBX preencher estrutura CO-PA
CRITERIA-ITEMNO_ACC = '0000000002'.
CRITERIA-FIELDNAME = 'PRCTR'.
*CRITERIA-CHARACTER = 'DUMMY'.
CRITERIA-CHARACTER = '0000001900'.
append criteria.
*
 call function 'BAPI_ACC_INVOICE_RECEIPT_POST' "#EC CI_USAGE_OK[2438131]
    exporting
      documentheader = documentheader
      customercpd    = customercpd
    importing
      obj_type       = obj_type
      obj_key        = obj_key
      obj_sys        = obj_sys
    tables
      accountpayable = accountpayable
      accountgl      = accountgl "#EC CI_USAGE_OK[2628704]
      accounttax     = accounttax
      currencyamount = currencyamount
      purchaseorder  = purchaseorder "#EC CI_USAGE_OK[2628704]
      purchaseamount = purchaseamount
      return         = return
      criteria       = criteria
      valuefield     = valuefield
      extension1     = extension1.

  call function 'BAPI_TRANSACTION_COMMIT'
    EXPORTING
      WAIT          = 'X'.




ENDFUNCTION.
