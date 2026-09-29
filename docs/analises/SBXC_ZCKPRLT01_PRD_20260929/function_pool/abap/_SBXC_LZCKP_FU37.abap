FUNCTION /sbxc/zckp_status_oc.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(DATA) TYPE  DATS
*"  TABLES
*"      DOC STRUCTURE  /SBXC/ZCKP_STATUS_OC_FAT_PAG
*"----------------------------------------------------------------------
*Função para:
* - Informar Saphety de pedidos de compra do tipo ZI
* aprovados num determinado dia
* - Informar Saphety de faturas criadas pelo cockpit num determinado
* dia
* - Informar Saohety de Facturas pagas num determinado dia
************************************************************************

  PERFORM envia_faturas_criadas  TABLES doc
          USING data.

  PERFORM envia_pagamentos_criadas  TABLES doc
          USING data.


  PERFORM act_tab_cockpit TABLES doc.

ENDFUNCTION.
