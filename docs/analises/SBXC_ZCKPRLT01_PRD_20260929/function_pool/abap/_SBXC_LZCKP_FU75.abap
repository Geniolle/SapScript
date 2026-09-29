FUNCTION /sbxc/zckp_check_nc.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(I_BUKRS) TYPE  BSEG-BUKRS
*"     REFERENCE(I_DOC_FI) TYPE  BSEG-BELNR
*"     REFERENCE(I_GJAHR) TYPE  BSEG-GJAHR
*"  EXPORTING
*"     REFERENCE(E_NC) TYPE  CHAR1
*"----------------------------------------------------------------------

*  Validar se se trata de uma nota de crédito

  DATA: l_belnr LIKE bseg-belnr.

  CLEAR l_belnr.
  SELECT SINGLE belnr FROM bseg INTO l_belnr  "#EC CI_NOORDER
  WHERE  bukrs = i_bukrs AND belnr = i_doc_fi AND
         gjahr = i_gjahr AND koart = 'K' AND shkzg = 'S'.
  IF sy-subrc EQ 0 AND NOT l_belnr IS INITIAL.
    e_nc = 'X'.
  ENDIF.

ENDFUNCTION.
