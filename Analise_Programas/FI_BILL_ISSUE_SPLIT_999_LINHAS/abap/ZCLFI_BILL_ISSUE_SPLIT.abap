*----------------------------------------------------------------------*
* Classe: ZCLFI_BILL_ISSUE_SPLIT
* Objetivo: Implementação da BAdI standard FI_BILL_ISSUE_SPLIT para
*           particionamento contábil de faturas SD com >999 linhas.
* Interface: IF_EX_FI_BILL_ISSUE_SPLIT
*----------------------------------------------------------------------*

CLASS ZCLFI_BILL_ISSUE_SPLIT DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES IF_BADI_INTERFACE .
    INTERFACES IF_EX_FI_BILL_ISSUE_SPLIT .

  PROTECTED SECTION.

  PRIVATE SECTION.

ENDCLASS.



CLASS ZCLFI_BILL_ISSUE_SPLIT IMPLEMENTATION.

*----------------------------------------------------------------------*
* Método: ACTIVATE_AUTOMATIC_SPLIT
* Descrição: Ativa o particionamento contábil automático.
*----------------------------------------------------------------------*
  METHOD IF_EX_FI_BILL_ISSUE_SPLIT~ACTIVATE_AUTOMATIC_SPLIT.

    E_AUTOMATIC_SPLIT = 'X'.

  ENDMETHOD.


*----------------------------------------------------------------------*
* Método: SET_NUMBER_OF_INVOICE_ITEMS
* Descrição: Define a quantidade de itens por lote de split.
* NOTA TÉCNICA: A linha e_number_of_invoice_items = 900 encontra-se
*               comentada no ambiente produtivo SAP PRD.
*----------------------------------------------------------------------*
  METHOD IF_EX_FI_BILL_ISSUE_SPLIT~SET_NUMBER_OF_INVOICE_ITEMS.

*    e_number_of_invoice_items = 900.

  ENDMETHOD.


*----------------------------------------------------------------------*
* Método: SET_DOCUMENT_TYPE_SUBSEQ
* Descrição: Define o tipo de documento contábil (BLART) para os
*           documentos subsequentes gerados pela quebra.
*           Documento principal = RH, Documentos subsequentes = ZV.
*----------------------------------------------------------------------*
  METHOD IF_EX_FI_BILL_ISSUE_SPLIT~SET_DOCUMENT_TYPE_SUBSEQ.

    E_DOCUMENT_TYPE_SUBSEQ = 'ZV'.

  ENDMETHOD.

ENDCLASS.
