FUNCTION /sbxc/zckp_guarda_alt_fat.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(CAB)
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"----------------------------------------------------------------------

  DATA: item  TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.


  DATA invoice_doc_item LIKE /sbxc/zckp_invi-invoice_doc_item.
* Preencher estrutura cabecalho

  MOVE-CORRESPONDING cab TO /sbxc/zckp_invh.
  if /sbxc/zckp_invh-em_tratamento is not initial and /sbxc/zckp_invh-user_tratamento is initial.
   /sbxc/zckp_invh-user_tratamento = sy-uname.
    /sbxc/zckp_invh-data_tratamento = sy-datum.
  endif.
*  move-corresponding /sbxc/zckp_invh to cab.
  UPDATE /sbxc/zckp_invh.


  SORT linha  DESCENDING.


  LOOP AT linha.


    MOVE-CORRESPONDING linha TO  /sbxc/zckp_invi.
    IF /sbxc/zckp_invi-processo =  /sbxc/zckp_invh-processo AND
           /sbxc/zckp_invi-ano = /sbxc/zckp_invh-ano AND
            /sbxc/zckp_invi-seqno = /sbxc/zckp_invh-seqno.

      SELECT MAX( invoice_doc_item ) INTO  invoice_doc_item
      FROM   /sbxc/zckp_invi
      WHERE
            processo =  /sbxc/zckp_invi-processo AND
            ano =  /sbxc/zckp_invi-ano AND
            seqno =   /sbxc/zckp_invi-seqno.
      IF /sbxc/zckp_invi-invoice_doc_item IS INITIAL.
        ADD 1 TO invoice_doc_item.
        /sbxc/zckp_invi-invoice_doc_item = invoice_doc_item.
      ENDIF.

      MODIFY /sbxc/zckp_invi.
      commit work.
    ENDIF.
    MOVE-CORRESPONDING linha TO  item.
    item-invoice_doc_item = /sbxc/zckp_invi-invoice_doc_item.
    APPEND item.
  ENDLOOP.

  COMMIT WORK.
*
  IF item[] IS NOT INITIAL.
** elimina da tabela de itens linhas entretanto apagadas
    SELECT * FROM  /sbxc/zckp_invi WHERE
         processo =  /sbxc/zckp_invi-processo AND
         ano =  /sbxc/zckp_invi-ano AND
        seqno =   /sbxc/zckp_invi-seqno.
*
      READ TABLE item WITH KEY
            processo = /sbxc/zckp_invi-processo
            ano = /sbxc/zckp_invi-ano
            seqno = /sbxc/zckp_invi-seqno
            invoice_doc_item = /sbxc/zckp_invi-invoice_doc_item.
*
      IF sy-subrc <> 0.
        DELETE /sbxc/zckp_invi FROM /sbxc/zckp_invi.

      ENDIF.
    ENDSELECT.

    REFRESH linha.
    LOOP AT item.
      MOVE-CORRESPONDING item TO linha.
      APPEND linha.
    ENDLOOP.

  ENDIF.
  SORT linha ASCENDING.

  COMMIT WORK AND WAIT.

ENDFUNCTION.
