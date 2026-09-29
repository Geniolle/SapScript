FUNCTION /sbxc/zckp_mm_cancela_pedido.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(POCHANGE) TYPE  /SBXC/ZBAPIPOCHANGE
*"     VALUE(TESTRUN) TYPE  BAPIFLAG-BAPIFLAG
*"  TABLES
*"      RETURN STRUCTURE  BAPIRET2
*"----------------------------------------------------------------------
* Criado por Cláudia Fernandes @ SBXC
* Função para criar eliminar itens do pedido de compra
************************************************************************

*"----------------------------------------------------------------------
* Funções
*"----------------------------------------------------------------------
*

  PERFORM deriva_campos_po_change
               USING pochange.

  REFRESH: return, lt_poitem, lt_poitemx.

  PERFORM preenche_estruturas_po_change
                USING pochange.

*"----------------------------------------------------------------------
* BAPI
*"----------------------------------------------------------------------
  CALL FUNCTION 'BAPI_PO_CHANGE'
     EXPORTING
       purchaseorder                = pedido
*   POHEADER                     =
*   POHEADERX                    =
*   POADDRVENDOR                 =
*   TESTRUN                      =
*   MEMORY_UNCOMPLETE            =
*   MEMORY_COMPLETE              =
*   POEXPIMPHEADER               =
*   POEXPIMPHEADERX              =
*   VERSIONS                     =
*   NO_MESSAGING                 =
*   NO_MESSAGE_REQ               =
*   NO_AUTHORITY                 =
*   NO_PRICE_FROM_PO             =
*   PARK_UNCOMPLETE              =
*   PARK_COMPLETE                =
* IMPORTING
*   EXPHEADER                    =
*   EXPPOEXPIMPHEADER            =
 TABLES
      return                       = return
      poitem                       = lt_poitem "#EC CI_USAGE_OK[2438131]
      poitemx                      = lt_poitemx. "#EC CI_USAGE_OK[2438131]
*   POADDRDELIVERY               =
*   POSCHEDULE                   =
*   POSCHEDULEX                  =
*   POACCOUNT                    =
*   POACCOUNTPROFITSEGMENT       =
*   POACCOUNTX                   =
*   POCONDHEADER                 =
*   POCONDHEADERX                =
*   POCOND                       =
*   POCONDX                      =
*   POLIMITS                     =
*   POCONTRACTLIMITS             =
*   POSERVICES                   =
*   POSRVACCESSVALUES            =
*   POSERVICESTEXT               =
*   EXTENSIONIN                  =
*   EXTENSIONOUT                 =
*   POEXPIMPITEM                 =
*   POEXPIMPITEMX                =
*   POTEXTHEADER                 =
*   POTEXTITEM                   =
*   ALLVERSIONS                  =
*   POPARTNER                    =
*   POCOMPONENTS                 =
*   POCOMPONENTSX                =
*   POSHIPPING                   =
*   POSHIPPINGX                  =
*   POSHIPPINGEXP                =
*   POHISTORY                    =
*   POHISTORY_TOTALS             =
*   POCONFIRMATION               =
*   SERIALNUMBER                 =
*   SERIALNUMBERX                =
*   INVPLANHEADER                =
*   INVPLANHEADERX               =
*   INVPLANITEM                  =
*   INVPLANITEMX                 =
*   POHISTORY_MA                 =

  CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'.

ENDFUNCTION.
