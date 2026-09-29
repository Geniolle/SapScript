FUNCTION /SBXC/ZCKP_MM_EST_REC.
*"--------------------------------------------------------------------
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
*"--------------------------------------------------------------------
* Estruturas de cabeçalho e linha
  DATA: header    TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        item      TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE,
        item_f    TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE,
        wa_header TYPE /sbxc/zckp_invh,
        materialdocument  TYPE  bapi2017_gm_head_ret-mat_doc,
        matdocumentyear TYPE  bapi2017_gm_head_ret-doc_year.
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.
  SORT item BY  po_number
                 po_item
                 ref_doc
                 ref_doc_year
                 ref_doc_item.

  DELETE ADJACENT DUPLICATES FROM item COMPARING po_number
                po_item
                ref_doc
                ref_doc_year
                ref_doc_item.
** Preencher estrutura cabecalho
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
*  ENDLOOP.
  LOOP AT header.

    CLEAR status_ok.
    PERFORM valida_status USING '0125'
                                header-processo
                                header-ano
                                header-seqno
                                ctrl-status1
                      CHANGING status_ok.

    CHECK status_ok = 'X'.


    indice = sy-tabix.
    REFRESH: item_f, it_return.
    LOOP AT item WHERE
      processo = header-processo AND
      ano      = header-ano      AND
      seqno    = header-seqno.

      MOVE-CORRESPONDING item TO item_f.
      APPEND item_f.
    ENDLOOP.


    MOVE-CORRESPONDING header TO wa_header.
    IF e_ucomm = 'ESTREC'.
      CALL FUNCTION '/SBXC/ZCKP_MM_ESTORNA_RECEPCAO'
        IMPORTING
          materialdocument = materialdocument
          matdocumentyear  = matdocumentyear
        TABLES
          itemdata         = item_f
        CHANGING
          headerdata       = header
          ctrl             = ctrl.
      refresh = 'X'.
    ENDIF.
  ENDLOOP.
* Preencher estrutura cabecalho
  REFRESH: linha.
  LOOP AT header.
    MOVE-CORRESPONDING header TO cab.
  ENDLOOP.
  LOOP AT item_f.
    MOVE-CORRESPONDING item_f TO linha.
    APPEND linha.
  ENDLOOP.
  IF invoicedocnumber IS NOT INITIAL.
    it_return-id = 'MIGO'.
    it_return-number = '012'.
    it_return-type = 'S'.
    it_return-message_v1 = materialdocument .
    APPEND it_return.
  ENDIF.
  CALL FUNCTION 'C14ALD_BAPIRET2_SHOW'
    TABLES
      i_bapiret2_tab = it_return.
ENDFUNCTION.
