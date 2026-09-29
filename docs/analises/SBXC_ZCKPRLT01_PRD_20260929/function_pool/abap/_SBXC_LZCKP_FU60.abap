FUNCTION /sbxc/zckp_call_miro.
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

 IF  ctrl-status1 EQ '3'.
      ctrl-status1 = 'E'. "processado com exito
      MESSAGE s012(/sbxc/zckp_cockpit).
    ELSEIF  ctrl-status1 EQ '6'.
      MESSAGE s018(/sbxc/zckp_cockpit).
ELSE.
*        PERFORM preenche_estruturas_inv
*                TABLES itemdata
*                USING headerdata.

"Valida duplicados
      DATA: lv_return TYPE sy-subrc.
      CALL FUNCTION '/SBXC/ZCKP_VALIDA_DUPLICADO'
        EXPORTING
          doctypesaphety = headerdata-doctypesaphety
        IMPORTING
          return         = lv_return
          CHANGING
            cab            = headerdata
            ctrl           = ctrl.
      IF lv_return NE 4.
        PERFORM bi_miro USING headerdata.

        COMMIT WORK AND WAIT.
        LOOP AT messtab.
          IF messtab-msgid = 'M8' AND ( messtab-msgnr = '418' or messtab-msgnr = '060' )..
            headerdata-doc_lo = messtab-msgv1.
            headerdata-ano_lanc = messtab-msgv2.
            UNPACK  headerdata-doc_lo TO   headerdata-doc_lo.

            invoicedocnumber = headerdata-doc_lo.
            fiscalyear = headerdata-ano_lanc.
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

            invoicedocnumber = headerdata-doc_lo.
            fiscalyear = headerdata-ano_lanc.
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

"ODC - 15_06_2020
*PERFORM actualiza_bd IN PROGRAM /sbxc/lversao_basef01 TABLES itemdata
*        USING invoicedocnumber headerdata-ano_lanc
*        CHANGING headerdata ctrl.
PERFORM actualiza_bd TABLES itemdata USING headerdata invoicedocnumber fiscalyear
                                          CHANGING ctrl.
"Fim ODC - 15_06_2020
endif.
ENDIF.

ENDFUNCTION.
