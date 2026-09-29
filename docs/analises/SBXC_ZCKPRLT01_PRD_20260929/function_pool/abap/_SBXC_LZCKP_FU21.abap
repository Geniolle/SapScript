FUNCTION /SBXC/ZCKP_MM_CRIA_FACTURAS_AN .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(HEADERDATA) TYPE  /SBXC/ZCKP_INVH
*"     VALUE(PROCESSO_ANT) TYPE  /SBXC/ZCKP_PROCESSO OPTIONAL
*"     VALUE(ANO_ANT) TYPE  GJAHR OPTIONAL
*"     VALUE(SEQNO_ANT) TYPE  /SBXC/ZCKP_SEQNO OPTIONAL
*"  EXPORTING
*"     VALUE(BELNR) TYPE  BELNR_D
*"     VALUE(PREDIT) TYPE  CHAR1
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI
*"      RETURN STRUCTURE  BAPIRET2
*"----------------------------------------------------------------------
* Função para registar dados das facturas na tabela de suporte ao
* cockpit
*********************************************************************

  MOVE-CORRESPONDING headerdata TO /sbxc/zckp_invh.
  DATA: item LIKE  /sbxc/zckp_invi-invoice_doc_item,
        org_cmp LIKE ekko-ekorg.

  DATA pedido_msg LIKE pedido.

  CLEAR pedido.


  UNPACK  headerdata-vendor TO forn.

* valida indice a inserir na tabela
  CALL FUNCTION 'NUMBER_GET_NEXT'
    EXPORTING
      nr_range_nr                   = '01'
      object                        = 'ZCKP_COCKP'
     quantity                      = '1'
*   SUBOBJECT                     = ' '
     toyear                        = sy-datum(4)
*   IGNORE_BUFFER                 = ' '
   IMPORTING
     number                        =  /sbxc/zckp_invh-seqno
*   QUANTITY                      =
*   RETURNCODE                    =
 EXCEPTIONS
   INTERVAL_NOT_FOUND            = 1
   NUMBER_RANGE_NOT_INTERN       = 2
   OBJECT_NOT_FOUND              = 3
   QUANTITY_IS_0                 = 4
   OTHERS                        = 5
            .
  IF sy-subrc <> 0.
 MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

  headerdata-seqno = /sbxc/zckp_invh-seqno.

  READ TABLE itemdata INDEX 1.
  IF  itemdata-po_number IS NOT INITIAL.
    pedido_msg = itemdata-po_number .
    SELECT SINGLE ebeln zterm ekorg  FROM ekko
          INTO (pedido, /sbxc/zckp_invh-zterm, org_cmp)
            WHERE
               ebeln = itemdata-po_number AND
               lifnr = forn AND
               bukrs = headerdata-comp_code.

  ENDIF.

  IF  itemdata-ref_1 IS NOT INITIAL.
    SELECT SINGLE ebeln zterm ekorg  FROM ekko  "#EC CI_NOORDER
          INTO (pedido, /sbxc/zckp_invh-zterm, org_cmp)
            WHERE
               lifnr = forn AND
               bukrs = headerdata-comp_code AND
               ihrez = itemdata-ref_1.
  ENDIF.
* verifica condição pagamento do pedido se não encontrar procura no Mestre Forn
  IF /sbxc/zckp_invh-zterm IS INITIAL.
    SELECT SINGLE zterm INTO /sbxc/zckp_invh-zterm FROM lfm1
      WHERE lifnr = forn
            AND ekorg = org_cmp.
  ENDIF.

  IF /sbxc/zckp_invh-ano IS INITIAL.
    /sbxc/zckp_invh-ano = sy-datum(4).
  ENDIF.

  IF headerdata-processo EQ space.

    CALL FUNCTION '/SBXC/ZCKP_DETERMINA_PROCESSO'
      EXPORTING
        headerdata = headerdata
        tax        = flag_tax
      IMPORTING
*       REFRESH    =
        processo   = /sbxc/zckp_invh-processo
        regra      = /sbxc/zckp_ctrl-regra_entrada
      TABLES
        itemdata   = itemdata
        return     = return.
    headerdata-processo = /sbxc/zckp_invh-processo.
  ENDIF.

* Valida nome do fornecedor
  SELECT SINGLE name1 FROM lfa1 INTO /sbxc/zckp_invh-name1 WHERE
         lifnr = forn.

  IF /sbxc/zckp_invh-ano IS INITIAL.
    /sbxc/zckp_ctrl-ano = sy-datum(4).
    /sbxc/zckp_invh-ano = sy-datum(4).
    headerdata-ano = sy-datum(4).

  ENDIF.

  INSERT  /sbxc/zckp_invh.

* Guardar tabela de histórico - cabeçalho
  MOVE-CORRESPONDING /sbxc/zckp_invh TO /sbxc/zckp_hinvh.

  INSERT /sbxc/zckp_hinvh.

  MOVE-CORRESPONDING /sbxc/zckp_invh TO /sbxc/zckp_ctrl.
  /sbxc/zckp_ctrl-data_in = sy-datum.
  /sbxc/zckp_ctrl-hora_in = sy-uzeit.
  /sbxc/zckp_ctrl-user_in = sy-uname.
  /sbxc/zckp_ctrl-status1 = '0'.

  SELECT SINGLE icon_status FROM /sbxc/zckp_tab10
    INTO /sbxc/zckp_ctrl-icon_status
  WHERE processo = /sbxc/zckp_invh-processo
    AND status = '0'.

  /sbxc/zckp_ctrl-processo_ant = processo_ant.
  /sbxc/zckp_ctrl-ano_ant = ano_ant.
  /sbxc/zckp_ctrl-seqno_ant = seqno_ant.

  IF /sbxc/zckp_invh-ano IS INITIAL.
    /sbxc/zckp_ctrl-ano = sy-datum(4).
  ENDIF.

  INSERT  /sbxc/zckp_ctrl.

  CONCATENATE /sbxc/zckp_ctrl-processo '/' /sbxc/zckp_ctrl-ano '/' /sbxc/zckp_ctrl-seqno INTO
   return-message.
  INSERT return INDEX 1.

  item = 0.


  LOOP AT itemdata.
    IF itemdata-po_number IS NOT INITIAL.
      SELECT SINGLE ebeln zterm ekorg  FROM ekko
                INTO (pedido, /sbxc/zckp_invh-zterm, org_cmp)
                  WHERE
                     ebeln = itemdata-po_number AND
                     lifnr = forn AND
                     bukrs = headerdata-comp_code.
    ENDIF.

    IF itemdata-ref_1 IS NOT INITIAL.
      SELECT SINGLE ebeln FROM ekko  "#EC CI_NOORDER
             INTO pedido
               WHERE
                  lifnr = forn AND
                  bukrs = headerdata-comp_code AND
                  ihrez = itemdata-ref_1.
    ENDIF.
    itemdata-processo = /sbxc/zckp_invh-processo.
    CLEAR /sbxc/zckp_invi.

    MOVE-CORRESPONDING itemdata TO /sbxc/zckp_invi.
    MOVE-CORRESPONDING /sbxc/zckp_invh TO /sbxc/zckp_invi.
    /sbxc/zckp_invi-po_number = pedido.

* Verifica nº de item pedido e item utilizando nº pedido e item saphety
    IF pedido IS NOT INITIAL.
      SELECT SINGLE ebelp meins FROM ekpo INTO  "#EC CI_NOORDER
           (/sbxc/zckp_invi-po_item, /sbxc/zckp_invi-po_unit)
      WHERE
             ebeln = pedido AND
             bednr = itemdata-trackingno.
    ENDIF.
    ADD 1 TO item.
    /sbxc/zckp_invi-invoice_doc_item = item.

* verifica a conta do serviço relacionada com o item do pedido de
*compra
    IF pedido IS NOT INITIAL.
      SELECT SINGLE sakto INTO /sbxc/zckp_invi-gl_account FROM ekkn WHERE  "#EC CI_NOORDER
           ebeln = pedido  AND
           ebelp = /sbxc/zckp_invi-po_item.
    ENDIF.
    /sbxc/zckp_invi-item_amount_forn = /sbxc/zckp_invi-item_amount.
    INSERT  /sbxc/zckp_invi.

*   Guardar tabela de histórico - linhas
    MOVE-CORRESPONDING /sbxc/zckp_invi TO /sbxc/zckp_hinvi.
    /sbxc/zckp_hinvi-po_number =  pedido_msg.
    INSERT /sbxc/zckp_hinvi.

    MODIFY itemdata.
  ENDLOOP.

  COMMIT WORK AND WAIT.

  CLEAR: belnr, predit.
  CALL FUNCTION '/SBXC/ZCKP_CONTABILIZA_AUT'
    EXPORTING
      processo = /sbxc/zckp_invh-processo
      seqno    = /sbxc/zckp_invh-seqno
      ano      = /sbxc/zckp_invh-ano
    IMPORTING
      belnr    = belnr
      predit   = predit.
  IF belnr NE space.
    CONCATENATE text-032 belnr
        INTO return-message SEPARATED BY space.
  ELSEIF predit NE space.
    CONCATENATE text-033 belnr
        INTO return-message SEPARATED BY space.
  ELSE.
    CONCATENATE text-034 /sbxc/zckp_invh-seqno
        INTO return-message SEPARATED BY space.
  ENDIF.

  return-type = 'S'.
  APPEND return.


ENDFUNCTION.
