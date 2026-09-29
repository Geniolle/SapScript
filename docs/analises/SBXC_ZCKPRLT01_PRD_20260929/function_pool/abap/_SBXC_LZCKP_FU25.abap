FUNCTION /SBXC/ZCKP_MM_REGISTA_FACTURA1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  EXPORTING
*"     VALUE(INVOICEDOCNUMBER) LIKE  BAPI_INCINV_FLD-INV_DOC_NO
*"     VALUE(FISCALYEAR) LIKE  BAPI_INCINV_FLD-FISC_YEAR
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI
*"  CHANGING
*"     VALUE(HEADERDATA) TYPE  /SBXC/ZCKP_INVH
*"     VALUE(CTRL) TYPE  /SBXC/ZCKP_CTRL
*"----------------------------------------------------------------------

* Função para criar facturas
*********************************************************************

** Preenche estruturas BAPI de facturas
*  REFRESH: it_itemdata,  it_return, it_glaccountdata, return, msg_ckp, it_accountingdata.
*
*  CLEAR wa_header.
*
*
*  READ TABLE  itemdata WITH KEY po_number = ' '.
*
*** Verifica se a factura ainda se mantem por registar
**  select single status1 from zckp_ctrl into zckp_ctrl-status1 where
**       processo = headerdata-processo and
**       seqno = headerdata-seqno and
**       ano = headerdata-ano.

** Asteriscado por causa da eliminação do status no header
**  IF /sbxc/zckp_ctrl-status1 = '3' AND ( /sbxc/zckp_ctrl-status1 NE headerdata-status1 ).
**    MESSAGE s016(/sbxc/zckp_cockpit).
**  ELSEIF  /sbxc/zckp_ctrl-status1 = '9' AND ( /sbxc/zckp_ctrl-status1 NE headerdata-status1 ).
**    MESSAGE s017(/sbxc/zckp_cockpit).

*  IF  ctrl-status1 = '6'.
*    MESSAGE s018(/sbxc/zckp_cockpit).
*  ELSE.
*    IF  ctrl-status1 EQ '9'.
*      MESSAGE s011(/sbxc/zckp_cockpit).
**   Doc. já está pré-editado, continuar processamento com opção "Via MIRO
*
*    ELSEIF  ctrl-status1 EQ '3'.
*      MESSAGE s012(/sbxc/zckp_cockpit).
*
*    ELSE.
*
*      PERFORM preenche_estruturas_inv
*                TABLES itemdata
*                USING headerdata.
*
*
**"----------------------------------------------------------------
** BAPI
**"----------------------------------------------------------------
** verifica existencia de factura dupla
*      DATA flag_bkpf(1).
*      COMMIT WORK AND WAIT.
*      TRANSLATE wa_header-ref_doc_no TO UPPER CASE.
*      SELECT * FROM bsip WHERE
*          bukrs = wa_header-comp_code AND
*          lifnr = headerdata-vendor AND
*          waers = headerdata-currency AND
*          bldat = headerdata-doc_date AND
*          xblnr = wa_header-ref_doc_no AND
*          wrbtr = headerdata-gross_amount AND
*          gjahr = headerdata-ano_lanc AND
*          shkzg = 'H'.
*        CLEAR flag_bkpf.
*        SELECT SINGLE * FROM bkpf WHERE belnr = bsip-belnr AND
*                                              gjahr = bsip-gjahr AND
*                                              xreversal EQ ' '.
*        IF sy-subrc = 0.
*          flag_bkpf = 'X'.
*        ENDIF.
*        SELECT SINGLE  * FROM rbkp WHERE
*                  bukrs = bkpf-bukrs AND
*                  belnr = bkpf-awkey(10) AND
*                  gjahr = bkpf-awkey+10(4) and
*                  stblg = ''.
*
*      ENDSELECT.
*
*     if (  sy-subrc = 0 ) or ( sy-subrc ne 0 and flag_bkpf = 'X' ).
*        MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr.
*
*
*
**   Verificar se fatura já foi registrada sob documento contábil & &
*      ELSE.
*        SORT it_itemdata BY invoice_doc_item.
*        SORT  it_glaccountdata BY invoice_doc_item.
*        SORT it_accountingdata BY invoice_doc_item ASCENDING serial_no ASCENDING.
*        REFRESH: it_return, msg_ckp .
*
*        CALL FUNCTION 'BAPI_INCOMINGINVOICE_CREATE'
*       EXPORTING
*         headerdata                = wa_header
**   ADDRESSDATA               =
*       IMPORTING
*         invoicedocnumber          = invoicedocnumber
*         fiscalyear                =  fiscalyear
*       TABLES
*         itemdata                  = it_itemdata
*         accountingdata            = it_accountingdata
*         glaccountdata             = it_glaccountdata
**   MATERIALDATA              =
**   TAXDATA                   =
**   WITHTAXDATA               =
**   VENDORITEMSPLITDATA       =
*         return                    = it_return
**   EXTENSIONIN               =
*               .
*        return[] = it_return[].
*        IF invoicedocnumber IS NOT INITIAL.
*          return-id = 'M8'.
*          return-number = '060'.
*          return-type = 'S'.
*          return-message_v1 = invoicedocnumber .
*          APPEND return.
*        ELSE.
*          CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'
**     IMPORTING
**       RETURN        =
*                    .
*        ENDIF.
*
*        LOOP AT return.
*          ADD 1 TO wa_msg-lineno.
*          IF return-id IS NOT INITIAL.
*            wa_msg-msgid = return-id.
*            wa_msg-msgno = return-number.
*
*          ENDIF.
*          wa_msg-msgty = return-type.
*          wa_msg-msgv1 =  return-message_v1.
*          wa_msg-msgv2 =  return-message_v2.
*          wa_msg-msgv3 =  return-message_v3.
*          wa_msg-msgv4 =  return-message_v4.
*
*
*          APPEND wa_msg TO msg_ckp .
**
*        ENDLOOP.
*
** apaga tabela de log
*        PERFORM log_delete TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .
** Guarda msg na tabela de log
*        PERFORM log TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .
*
*        IF invoicedocnumber IS NOT INITIAL.
*          CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
*            EXPORTING
*              wait = 'X'.
**   IMPORTING
**     RETURN        =
*
** guarda na tabela do cockpit o doc. de FI gerado
*          MOVE fiscalyear TO ld_aworg.
*
*          CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
*            EXPORTING
*              i_awtyp      = 'RMRP'
*              i_awref      = invoicedocnumber
*              i_aworg      = ld_aworg
*            TABLES
*              e_accdn      = lt_accdn
*            EXCEPTIONS
*              no_doc_found = 1
*              OTHERS       = 2.
*
*
*          IF sy-subrc = 0.
*            LOOP AT lt_accdn INTO ls_accdn.
*              CHECK ls_accdn-belnr IS NOT INITIAL.
*              headerdata-doc_lo = invoicedocnumber.
*              headerdata-doc_fi = ls_accdn-belnr.
*              headerdata-ano_lanc = ls_accdn-gjahr.
*              headerdata-doc_estorno = ' '.
*              headerdata-data_criacao = sy-datum.
*
*
*              PERFORM actualiza_bd TABLES itemdata USING headerdata invoicedocnumber headerdata-ano_lanc
*                                                CHANGING ctrl.
*
** anexa doc. panagon ao doc de fi
**              perform anexos using headerdata-id_panagon ls_accdn-bukrs ls_accdn-belnr ls_accdn-gjahr.
** anexa URL
*              IF headerdata-url IS NOT INITIAL.
*
*                CALL FUNCTION '/SBXC/ZCKP_ANEXA_URL'
*                  EXPORTING
*                    bukrs = ls_accdn-bukrs
*                    belnr = ls_accdn-belnr
*                    gjahr = ls_accdn-gjahr
*                    url   = headerdata-url.
*              ENDIF.
*
*
** verifica se existe processo anterior, se sim, actualiza também os dados do anterior

** Select desnecessário porque a tabela de control passa a vir como parâmetro Importação
**              SELECT SINGLE processo_ant ano_ant seqno_ant FROM /sbxc/zckp_ctrl
**                INTO CORRESPONDING FIELDS OF /sbxc/zckp_ctrl WHERE
**                      processo = headerdata-processo AND
**                     ano = headerdata-ano AND
**                     seqno = headerdata-seqno.

*
*              IF  ctrl-processo_ant IS NOT INITIAL.
*                UPDATE zckp_inv_header SET
*                    doc_fi = ls_accdn-belnr
*                    doc_lo = invoicedocnumber
*                    doc_estorno = ' '
*                    ano_lanc = fiscalyear
*                    enviado_saphety = ' '
*                    data_criacao = sy-datum
*                    WHERE processo = ctrl-processo_ant AND
*                      ano = ctrl-ano_ant AND
*                      seqno = ctrl-seqno_ant.
*              ENDIF.
*
*            ENDLOOP.
*          ENDIF.
*
*        ENDIF.
*
**        CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
**          TABLES
**            i_message_tab = msg_ckp.
*
*        REFRESH msg_ckp. CLEAR wa_msg-lineno.
*
*      ENDIF.
*
*    ENDIF.
*  ENDIF.

ENDFUNCTION.
