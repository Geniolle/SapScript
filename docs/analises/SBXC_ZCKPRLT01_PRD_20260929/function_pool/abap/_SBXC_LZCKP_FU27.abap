FUNCTION /sbxc/zckp_mm_regist_fac_pree1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  EXPORTING
*"     VALUE(INVOICEDOCNUMBER) LIKE  BAPI_INCINV_FLD-INV_DOC_NO
*"     VALUE(FISCALYEAR) LIKE  BAPI_INCINV_FLD-FISC_YEAR
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI
*"      RETURN STRUCTURE  BAPIRET2
*"  CHANGING
*"     VALUE(HEADERDATA) TYPE  /SBXC/ZCKP_INVH
*"     VALUE(CTRL) TYPE  /SBXC/ZCKP_CTRL
*"----------------------------------------------------------------------

* Função para criar facturas
*********************************************************************

* Preenche estruturas BAPI de facturas
  REFRESH: it_itemdata,  it_return, it_glaccountdata, return,
   msg_ckp,  bdcdata, it_accountingdata.
  CLEAR: wa_header, it_return, msg_ckp.

  READ TABLE itemdata WITH KEY po_number = ' '.

  IF  ctrl-status1 EQ '3'.
    ctrl-status1 = 'E'. "processado com exito

    MESSAGE s012(/sbxc/zckp_cockpit).

  ELSEIF  ctrl-status1 NE '9'.
    PERFORM preenche_estruturas_inv
              TABLES itemdata
              USING headerdata.

    CALL FUNCTION 'MESSAGES_INITIALIZE'.
*"----------------------------------------------------------------
* BAPI
*"----------------------------------------------------------------
* se ainda não existe factura pre editada
    SORT it_itemdata BY invoice_doc_item.
    SORT  it_glaccountdata BY invoice_doc_item.
*    SORT it_accountingdata BY invoice_doc_item ASCENDING serial_no ASCENDING.

    REFRESH: it_return, msg_ckp .
    CALL FUNCTION 'BAPI_INCOMINGINVOICE_PARK' "#EC CI_USAGE_OK[2438131]
      EXPORTING
        headerdata          = wa_header
*       ADDRESSDATA         =
      IMPORTING
        invoicedocnumber    = invoicedocnumber
        fiscalyear          = fiscalyear
      TABLES
        itemdata            = it_itemdata
        accountingdata      = it_accountingdata
        glaccountdata       = it_glaccountdata
*       MATERIALDATA        = "#EC CI_USAGE_OK[2438131]
*       TAXDATA             =
*       WITHTAXDATA         =
*       VENDORITEMSPLITDATA =
        return              = it_return
*       EXTENSIONIN         =
      .


    LOOP AT return.
      ADD 1 TO wa_msg-lineno.
      IF return-id IS NOT INITIAL.
        wa_msg-msgid = return-id.
        wa_msg-msgno = return-number.

      ENDIF.
      wa_msg-msgty = return-type.
      wa_msg-msgv1 =  return-message_v1.
      wa_msg-msgv2 =  return-message_v2.
      wa_msg-msgv3 =  return-message_v3.
      wa_msg-msgv4 =  return-message_v4.
      IF wa_msg-msgid = 'M8' AND wa_msg-msgty = 'E' AND
               wa_msg-msgno = '496'.
* valida PEP bloqueados
        LOOP AT itemdata.
          SELECT SINGLE objnr FROM prps INTO prps-objnr
            WHERE pspnr = itemdata-wbs_elem.
          SELECT SINGLE * FROM jest WHERE
                 objnr =  prps-objnr
                 AND stat = 'I0065' AND
                 inact = ' '.
          IF sy-subrc = 0.
            wa_msg-msgid = '/sbxc/zckp_cockpit'.
            wa_msg-msgno = '029'.

            wa_msg-msgv1 = itemdata-po_number.
            wa_msg-msgv2 = itemdata-po_item.

            APPEND wa_msg TO msg_ckp .
          ENDIF.
        ENDLOOP.

        APPEND wa_msg TO msg_ckp .
      ENDIF.
    ENDLOOP.

** apaga tabela de log
    PERFORM log_delete TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .
** Guarda msg na tabela de log
    PERFORM log TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .

    IF invoicedocnumber IS NOT INITIAL.
      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          wait = 'X'.
*   IMPORTING
*     RETURN        =


* guarda na tabela do cockpit o doc. de FI gerado
      MOVE fiscalyear TO ld_aworg.
      CLEAR lt_accdn. REFRESH lt_accdn.
      CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
        EXPORTING
          i_awtyp      = 'RMRP'
          i_awref      = invoicedocnumber
          i_aworg      = ld_aworg
        TABLES
          e_accdn      = lt_accdn
        EXCEPTIONS
          no_doc_found = 1
          OTHERS       = 2.


      IF sy-subrc = 0.
        LOOP AT lt_accdn INTO ls_accdn
          WHERE LDGRP IS INITIAL. "ODC - 04_11_2020

          UPDATE /sbxc/zckp_invh SET
          doc_fi = ls_accdn-belnr
          doc_lo = invoicedocnumber
          doc_estorno = ' '
          ano_lanc = ls_accdn-gjahr
*          ano_lanc = fiscalyear
          data_criacao = sy-datum
*          enviado_saphety = ' '
          WHERE processo = headerdata-processo AND
                     ano = headerdata-ano AND
                   seqno = headerdata-seqno.


** verifica se existe processo anterior, se sim, actualiza também os dados do anterior
* Select desnecessário porque a tabela de control passa a vir como parâmetro Importação
*          SELECT SINGLE processo_ant ano_ant seqno_ant FROM /sbxc/zckp_ctrl
*            INTO CORRESPONDING FIELDS OF /sbxc/zckp_ctrl WHERE
*                  processo = headerdata-processo AND
*                 ano = headerdata-ano AND
*                 seqno = headerdata-seqno.


          IF  ctrl-processo_ant IS NOT INITIAL.
            UPDATE /sbxc/zckp_invh SET
                doc_fi = ls_accdn-belnr
                doc_lo = invoicedocnumber
                doc_estorno = ' '
                ano_lanc = fiscalyear
*                enviado_saphety = ' '
                WHERE processo = ctrl-processo_ant AND
                  ano = ctrl-ano_ant AND
                  seqno = ctrl-seqno_ant.
          ENDIF.


          UPDATE /sbxc/zckp_ctrl SET
          status1 = '9'
          WHERE processo = headerdata-processo AND
          ano = headerdata-ano AND
          seqno = headerdata-seqno.

          COMMIT WORK AND WAIT.
        ENDLOOP.

      ENDIF.
      headerdata-doc_lo = invoicedocnumber.
      headerdata-doc_fi = ls_accdn-belnr.
      headerdata-ano_lanc = fiscalyear.
      headerdata-doc_estorno = ' '.
      headerdata-data_criacao = sy-datum.

      IF ctrl-aut_man NE 'A'.
        PERFORM bi_mir4 USING invoicedocnumber fiscalyear.
      ENDIF.
      COMMIT WORK AND WAIT.
      LOOP AT messtab.
        IF messtab-msgid = 'M8' AND messtab-msgnr = '418'.
          headerdata-doc_lo = messtab-msgv1.
          headerdata-ano_lanc = messtab-msgv2.
          UNPACK  headerdata-doc_lo TO   headerdata-doc_lo.
* anexa doc. panagon ao doc de fi
          CLEAR lt_accdn. REFRESH lt_accdn.
          MOVE  fiscalyear TO ld_aworg.
          CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
            EXPORTING
              i_awtyp      = 'RMRP'
              i_awref      = headerdata-doc_lo
              i_aworg      = ld_aworg
            TABLES
              e_accdn      = lt_accdn
            EXCEPTIONS
              no_doc_found = 1
              OTHERS       = 2.


          IF sy-subrc = 0.
            LOOP AT lt_accdn INTO ls_accdn
              WHERE LDGRP IS INITIAL. "ODC - 04_11_2020
              headerdata-doc_fi = ls_accdn-belnr.
              headerdata-comp_code = ls_accdn-bukrs.
              headerdata-ano_lanc = ls_accdn-gjahr.
            ENDLOOP.
          ENDIF.
        ENDIF.
        IF messtab-msgid = 'M8' AND messtab-msgnr = '660'.
          headerdata-doc_lo = messtab-msgv2(10).
          headerdata-ano_lanc = messtab-msgv2+11(4).
          UNPACK  headerdata-doc_lo TO   headerdata-doc_lo.
* anexa doc. panagon ao doc de fi
          CLEAR lt_accdn. REFRESH lt_accdn.
          MOVE  fiscalyear TO ld_aworg.
          CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
            EXPORTING
              i_awtyp      = 'RMRP'
              i_awref      = headerdata-doc_lo
              i_aworg      = ld_aworg
            TABLES
              e_accdn      = lt_accdn
            EXCEPTIONS
              no_doc_found = 1
              OTHERS       = 2.


          IF sy-subrc = 0.
            LOOP AT lt_accdn INTO ls_accdn
              WHERE LDGRP IS INITIAL. "ODC - 04_11_2020
              headerdata-doc_fi = ls_accdn-belnr.
              headerdata-comp_code = ls_accdn-bukrs.
              headerdata-ano_lanc = ls_accdn-gjahr.
            ENDLOOP.
          ENDIF.

        ENDIF.
      ENDLOOP.
      PERFORM actualiza_bd TABLES itemdata USING headerdata invoicedocnumber fiscalyear
                                        CHANGING ctrl.


    else.
      PERFORM grava_alteracoes TABLES itemdata USING headerdata.
    ENDIF.

  ELSE.
    IF ctrl-aut_man NE 'A'.
      PERFORM bi_mir4 USING headerdata-doc_lo headerdata-ano_lanc.
    ENDIF.
    COMMIT WORK AND WAIT.
    LOOP AT messtab.
      IF messtab-msgid = 'M8' AND messtab-msgnr = '418'.
        headerdata-doc_lo = messtab-msgv1.
        headerdata-ano_lanc = messtab-msgv2.
        UNPACK  headerdata-doc_lo TO   headerdata-doc_lo.
        CLEAR lt_accdn. REFRESH lt_accdn.
        MOVE  fiscalyear TO ld_aworg.
        CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
          EXPORTING
            i_awtyp      = 'RMRP'
            i_awref      = headerdata-doc_lo
            i_aworg      = ld_aworg
          TABLES
            e_accdn      = lt_accdn
          EXCEPTIONS
            no_doc_found = 1
            OTHERS       = 2.


        IF sy-subrc = 0.
          LOOP AT lt_accdn INTO ls_accdn
            WHERE LDGRP IS INITIAL. "ODC - 04_11_2020
            headerdata-doc_fi = ls_accdn-belnr.
            headerdata-comp_code = ls_accdn-bukrs.
            headerdata-ano_lanc = ls_accdn-gjahr.
          ENDLOOP.
        ENDIF.

      ENDIF.


      IF messtab-msgid = 'M8' AND messtab-msgnr = '660'.
        headerdata-doc_lo = messtab-msgv2(10).
        headerdata-ano_lanc = messtab-msgv2+11(4).
        UNPACK  headerdata-doc_lo TO   headerdata-doc_lo.
* anexa doc. panagon ao doc de fi
        CLEAR lt_accdn. REFRESH lt_accdn.
        MOVE  fiscalyear TO ld_aworg.
        CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
          EXPORTING
            i_awtyp      = 'RMRP'
            i_awref      = headerdata-doc_lo
            i_aworg      = ld_aworg
          TABLES
            e_accdn      = lt_accdn
          EXCEPTIONS
            no_doc_found = 1
            OTHERS       = 2.


        IF sy-subrc = 0.
          LOOP AT lt_accdn INTO ls_accdn
            WHERE LDGRP IS INITIAL. "ODC - 04_11_2020
            headerdata-doc_fi = ls_accdn-belnr.
            headerdata-comp_code = ls_accdn-bukrs.
            headerdata-ano_lanc = ls_accdn-gjahr.
          ENDLOOP.
        ENDIF.
      ENDIF.
    ENDLOOP.


    PERFORM actualiza_bd TABLES itemdata USING headerdata headerdata-doc_lo headerdata-ano_lanc
                                      CHANGING ctrl.


  ENDIF.

  CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
    TABLES
      i_message_tab = msg_ckp.

  REFRESH msg_ckp. CLEAR wa_msg-lineno.

ENDFUNCTION.
