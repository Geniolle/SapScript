  METHOD PROCESS_PMNT_TRANS.

    DATA: LT_BSEG_PAY_GL LIKE IT_BSEG_PAY_GL,
          LT_BSEG_PAY_BANK LIKE IT_BSEG_PAY_GL,
          LT_BSEG_PAY_OTHERS LIKE IT_BSEG_PAY_GL,
          LS_BSEG_PAY_GL LIKE LINE OF IT_BSEG_PAY_GL,
          LS_BSEG_PAY_VC LIKE LINE OF IT_BSEG_PAY_VC,
          LS_BSEG_PAY_GL_CURR_DOC TYPE FIAPPT_S_TRANS_CBR,
          LT_PAID_ITEMS_CURR_DOC LIKE IT_PAID_ITEMS,
          LS_PAID_ITEMS_CURR_DOC LIKE LINE OF LT_PAID_ITEMS_CURR_DOC,
          LS_LFA1 LIKE LINE OF IT_LFA1,
          LS_KNA1 LIKE LINE OF IT_KNA1,
          LS_FIAPPTD_TRANS LIKE LINE OF ET_FIAPPTD_TRANS.
    DATA: LT_EX_FIAPPTD_TRANS LIKE ET_FIAPPTD_TRANS,
          LS_EX_FIAPPTD_TRANS LIKE LINE OF LT_EX_FIAPPTD_TRANS,
          LS_EX_TRANS LIKE LINE OF LT_EX_FIAPPTD_TRANS.
    DATA: LT_FIAPPTD_TRANS_EXIT LIKE ET_FIAPPTD_TRANS,
          LS_FIAPPTD_TRANS_EXIT LIKE LINE OF LT_FIAPPTD_TRANS_EXIT.
    DATA: LS_XML_ITEMS LIKE LINE OF ET_XML_ITEMS.
    DATA: LV_LIFNR, LV_KUNNR.
    DATA: LV_MULTIPLE_BANKS TYPE FLAG,
          LV_GL_PAYMENTS TYPE FLAG.
    DATA: LV_LAND TYPE LAND1,
          LV_LAND_ISO TYPE INTCA3.
    DATA: OBJ_CBR_HANDLE TYPE REF TO BADI_RFFOPT_CBR.
    DATA: LV_REGISTER_EXISTS TYPE FLAG,
          LV_REGISTER_ID TYPE FIAPPT_REG_ID.
    DATA: LV_ORIGIN_RUN_ID TYPE FIAPPT_RUN_ID.
    DATA: LV_FOREIGN_TRANS TYPE FLAG,
          LT_BSEG_PARTIAL TYPE  FIAPPT_T_TRANS_CBR,
          LT_BSEG_OTHER_GL TYPE FIAPPT_T_TRANS_CBR,
          LT_BSEG_REM_ALL_WO_INV   TYPE FIAPPT_T_TRANS_CBR, "it_bseg_rem_all without invoice reference
          LS_BSEG_REM_ALL_WO_INV   TYPE FIAPPT_S_TRANS_CBR,
          LT_BSEG_REM_ALL_WITH_INV TYPE FIAPPT_T_TRANS_CBR, "it_bseg_rem_all with invoice reference
          LS_BSEG_REM_ALL_WITH_INV TYPE FIAPPT_S_TRANS_CBR,
          LS_BSEG_REM_ALL          TYPE FIAPPT_S_TRANS_CBR,
          LS_PAID_ITEMS            TYPE FIAPPT_S_TRANS_CBR,
          LV_PAY_WO_INV            TYPE FLAG,
          LV_PAY_WITH_INV          TYPE FLAG,
          LS_TCURC TYPE TCURC.

*    DATA: lv_index TYPE sy-index VALUE 2500.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID,
          FLAG_DIFF_PERIOD TYPE FLAG.

    CONSTANTS: CV_CURR TYPE WAERS VALUE 'EUR'.

**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.
*****same for all lines of table FIAPPTD_TRANS for current
* run (except for changed/reversed registers, where the
* origin_run_id will be overwritten)
    CALL METHOD INIT_FIAPPTD_TRANS_STR
      IMPORTING
        ES_FIAPPTD_TRANS = LS_FIAPPTD_TRANS.

*    Partial payment lines
    LT_BSEG_PARTIAL[] = IT_BSEG_REM_ALL[].
    DELETE LT_BSEG_PARTIAL WHERE REBZG IS INITIAL AND
    REBZJ IS INITIAL AND REBZZ IS INITIAL.

*    Other G/L lines in the payment document
    IF GT_OTH_ACC_RANGE[] IS NOT INITIAL.
      LT_BSEG_OTHER_GL[] = IT_BSEG_REM_ALL[].
      DELETE LT_BSEG_OTHER_GL WHERE KOART NE 'S'.
    ENDIF.

    SORT IT_BSEG_PAY_GL BY BUKRS BELNR GJAHR BUZEI HKONT.
    LT_BSEG_PAY_GL = IT_BSEG_PAY_GL.
* will ensure one bank line per belnr
    DELETE ADJACENT DUPLICATES FROM IT_BSEG_PAY_GL
         COMPARING BUKRS BELNR GJAHR.

* separate the payment lines in it_bseg_rem_all in a table
* that collects the payment lines with an invoice reference lt_bseg_rem_all_with_inv
* from the others without invoice reference lt_bseg_rem_all_wo_inv
    LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_REM_ALL WHERE KOART = 'D' OR KOART = 'K' .

      IF LS_BSEG_REM_ALL-AUGBL EQ LS_BSEG_REM_ALL-BELNR OR
        ( LS_BSEG_REM_ALL-REBZG IS NOT INITIAL AND LS_BSEG_REM_ALL-REBZJ IS NOT INITIAL
                                AND LS_BSEG_REM_ALL-REBZZ IS NOT INITIAL ) .
        APPEND LS_BSEG_REM_ALL TO LT_BSEG_REM_ALL_WITH_INV.
      ELSE.
        APPEND LS_BSEG_REM_ALL TO LT_BSEG_REM_ALL_WO_INV.
      ENDIF.

    ENDLOOP.

    LOOP AT IT_BSEG_PAY_GL INTO LS_BSEG_PAY_GL.

      CLEAR: LV_MULTIPLE_BANKS, LV_GL_PAYMENTS, LV_PAY_WO_INV, LV_PAY_WITH_INV.

      LOOP AT LT_BSEG_REM_ALL_WO_INV INTO LS_BSEG_REM_ALL_WO_INV WHERE
                    BUKRS = GV_BUKRS AND BELNR = LS_BSEG_PAY_GL-BELNR.
        LV_PAY_WO_INV = 'X'.
        EXIT.
      ENDLOOP.

      LOOP AT LT_BSEG_REM_ALL_WITH_INV INTO LS_BSEG_REM_ALL_WITH_INV WHERE
                    BUKRS = GV_BUKRS AND BELNR = LS_BSEG_PAY_GL-BELNR.
        LV_PAY_WITH_INV = 'X'.
        EXIT.
      ENDLOOP.

*    Internal table that are passed to the BAdI method
      CLEAR: LT_BSEG_PAY_BANK,
             LT_BSEG_PAY_OTHERS,
             LT_PAID_ITEMS_CURR_DOC,
             FLAG_DIFF_PERIOD.

      IF LS_BSEG_PAY_GL-XREVERSAL = '2'.
        SELECT * INTO TABLE LT_EX_FIAPPTD_TRANS
          FROM FIAPPTD_TRANS WHERE BUKRS = GV_BUKRS
           AND PAY_DOC = LS_BSEG_PAY_GL-STBLG
           AND DOC_YEAR = LS_BSEG_PAY_GL-STJAH
           AND PERIOD = GV_PERIOD.
*           AND run_id = gv_prev_runid
*           AND pay_doc = ''.
        IF LT_EX_FIAPPTD_TRANS IS INITIAL.
          SELECT * INTO TABLE LT_EX_FIAPPTD_TRANS
            FROM FIAPPTD_TRANS WHERE BUKRS = GV_BUKRS
             AND PAY_DOC = LS_BSEG_PAY_GL-STBLG
             AND PAY_DOC_YEAR = LS_BSEG_PAY_GL-STJAH.
*             AND period = gv_period.
*             AND run_id = gv_prev_runid.
          IF SY-SUBRC = 0.
            FLAG_DIFF_PERIOD = 'X'.
          ENDIF.
        ENDIF.

        IF LT_EX_FIAPPTD_TRANS IS NOT INITIAL.

          READ TABLE LT_EX_FIAPPTD_TRANS INTO LS_EX_TRANS
          WITH KEY REGISTER_ID = 'A'.
          IF SY-SUBRC NE  0.
            LOOP AT LT_EX_FIAPPTD_TRANS INTO LS_EX_FIAPPTD_TRANS WHERE
              REGISTER_ID = 'M'.

              IF LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'C' OR
                 LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'M'.
                LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'A'.
                LS_EX_FIAPPTD_TRANS-RUN_ID = GV_RUNID.

                IF FLAG_DIFF_PERIOD EQ 'X'.
*               Add the next sequence to the run id
                  GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
                  LV_RUNID_SEQ = GV_RUNID_SEQ.
                  SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
                  CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
                  LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'C'.
                  LS_EX_FIAPPTD_TRANS-REFERENCE_DATE = LS_BSEG_PAY_GL-BUDAT.
                ENDIF.
" Fixed SCB indicator for reversed docu
                LS_EX_FIAPPTD_TRANS-SCB_IND = 'A4010'.
                IF LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'S'.
                  LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'E'.
                ELSE.
                  LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'S'.
                ENDIF.

                MOVE-CORRESPONDING LS_EX_FIAPPTD_TRANS TO LS_XML_ITEMS.

                IF LS_XML_ITEMS-WAERS NE CV_CURR.
                  CALL METHOD CONVERT_AMOUNT_FOR_FILE
                    EXPORTING
                      IV_WAERS  = LS_XML_ITEMS-WAERS
                      IV_AMOUNT = LS_XML_ITEMS-WRBTR
                    IMPORTING
                      EV_AMOUNT = LS_XML_ITEMS-WRBTR.
                ENDIF.
                CLEAR LS_TCURC.
                READ TABLE GT_TCURC INTO LS_TCURC WITH KEY
                  WAERS = LS_XML_ITEMS-WAERS BINARY SEARCH.
                IF SY-SUBRC = 0 AND LS_TCURC-ISOCD IS NOT INITIAL.
                  LS_XML_ITEMS-ISOCD = LS_TCURC-ISOCD.
                  LS_EX_FIAPPTD_TRANS-ISOCD = LS_TCURC-ISOCD.
                ENDIF.

                APPEND LS_XML_ITEMS TO ET_XML_ITEMS.
                APPEND LS_EX_FIAPPTD_TRANS TO ET_FIAPPTD_TRANS.
                DELETE LT_EX_FIAPPTD_TRANS WHERE BELNR = LS_EX_FIAPPTD_TRANS-BELNR
                    AND BUZEI = LS_EX_FIAPPTD_TRANS-BUZEI.
              ELSE.
                "Do nothing.
              ENDIF.
            ENDLOOP.

            IF LT_EX_FIAPPTD_TRANS IS NOT INITIAL.
              LOOP AT LT_EX_FIAPPTD_TRANS INTO LS_EX_FIAPPTD_TRANS .

                IF LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'C' OR
                   LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'M'.

                  LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'A'.
                  LS_EX_FIAPPTD_TRANS-RUN_ID = GV_RUNID.

                  IF FLAG_DIFF_PERIOD EQ 'X'.
*               Add the next sequence to the run id
                    GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
                    LV_RUNID_SEQ = GV_RUNID_SEQ.
                    SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
                    CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
                    LS_EX_FIAPPTD_TRANS-REGISTER_ID = 'C'.
                    LS_EX_FIAPPTD_TRANS-REFERENCE_DATE = LS_BSEG_PAY_GL-BUDAT.
                  ENDIF.
" Fixed SCB indicator for docu reversal
                  LS_EX_FIAPPTD_TRANS-SCB_IND = 'A4010'.
                  IF LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'S'.
                    LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'E'.
                  ELSE.
                    LS_EX_FIAPPTD_TRANS-VALUE_TYPE = 'S'.
                  ENDIF.

                  MOVE-CORRESPONDING LS_EX_FIAPPTD_TRANS TO LS_XML_ITEMS.

                  IF LS_XML_ITEMS-WAERS NE CV_CURR.
                    CALL METHOD CONVERT_AMOUNT_FOR_FILE
                      EXPORTING
                        IV_WAERS  = LS_XML_ITEMS-WAERS
                        IV_AMOUNT = LS_XML_ITEMS-WRBTR
                      IMPORTING
                        EV_AMOUNT = LS_XML_ITEMS-WRBTR.
                  ENDIF.

                  CLEAR LS_TCURC.
                  READ TABLE GT_TCURC INTO LS_TCURC WITH KEY
                    WAERS = LS_XML_ITEMS-WAERS BINARY SEARCH.
                  IF SY-SUBRC = 0 AND LS_TCURC-ISOCD IS NOT INITIAL.
                    LS_XML_ITEMS-ISOCD = LS_TCURC-ISOCD.
                    LS_EX_FIAPPTD_TRANS-ISOCD = LS_TCURC-ISOCD.
                  ENDIF.

                  APPEND LS_XML_ITEMS TO ET_XML_ITEMS.
                  APPEND LS_EX_FIAPPTD_TRANS TO ET_FIAPPTD_TRANS.
                ELSE.
                  "Do nothing.
                ENDIF.
              ENDLOOP.
            ENDIF.
          ENDIF.
        ENDIF.
        CONTINUE.
      ENDIF.

* Check for multiple banks
      LOOP AT LT_BSEG_PAY_GL INTO LS_BSEG_PAY_GL_CURR_DOC
        WHERE BUKRS = LS_BSEG_PAY_GL-BUKRS
          AND BELNR = LS_BSEG_PAY_GL-BELNR
          AND GJAHR = LS_BSEG_PAY_GL-GJAHR
          AND HKONT NE LS_BSEG_PAY_GL-HKONT .

        LV_MULTIPLE_BANKS = 'X'.
        EXIT.
      ENDLOOP.

      IF LV_MULTIPLE_BANKS NE 'X'.
* Check if there are gl payments - if yes, pass the entire document to user exit
        READ TABLE IT_BSEG_PAY_VC INTO LS_BSEG_PAY_VC
          WITH KEY BUKRS = LS_BSEG_PAY_GL-BUKRS
                   BELNR = LS_BSEG_PAY_GL-BELNR
                   GJAHR = LS_BSEG_PAY_GL-GJAHR
                   KOART = 'S'.
        IF SY-SUBRC EQ 0.
          LV_GL_PAYMENTS = 'X'.               "g/l payments
        ELSE.
          READ TABLE IT_BSEG_PAY_VC INTO LS_BSEG_PAY_VC
          WITH KEY BUKRS = LS_BSEG_PAY_GL-BUKRS
                   BELNR = LS_BSEG_PAY_GL-BELNR
                   GJAHR = LS_BSEG_PAY_GL-GJAHR.
          IF SY-SUBRC NE 0.
*    Direct gl posting to bank gl account
            LV_GL_PAYMENTS = 'X'.
            LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_PAY_VC WHERE
              BUKRS = LS_BSEG_PAY_GL-BUKRS AND
                   BELNR = LS_BSEG_PAY_GL-BELNR AND
                   GJAHR = LS_BSEG_PAY_GL-GJAHR AND
                   ( KOART = 'K' OR KOART = 'D' ).
              CLEAR LV_GL_PAYMENTS.
              EXIT.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDIF.

      IF LV_MULTIPLE_BANKS EQ 'X' OR LV_GL_PAYMENTS = 'X'.
* Get all lines of the document and call user exit.

        LOOP AT LT_BSEG_PAY_GL INTO LS_BSEG_PAY_GL_CURR_DOC
                     WHERE BUKRS = LS_BSEG_PAY_GL-BUKRS
                       AND BELNR = LS_BSEG_PAY_GL-BELNR
                       AND GJAHR = LS_BSEG_PAY_GL-GJAHR.
          APPEND LS_BSEG_PAY_GL_CURR_DOC TO LT_BSEG_PAY_BANK.

        ENDLOOP.
        CLEAR: LS_BSEG_PAY_GL_CURR_DOC.
        LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_PAY_GL_CURR_DOC
          WHERE BUKRS = LS_BSEG_PAY_GL-BUKRS
            AND BELNR = LS_BSEG_PAY_GL-BELNR
            AND GJAHR = LS_BSEG_PAY_GL-GJAHR.
          APPEND LS_BSEG_PAY_GL_CURR_DOC TO LT_BSEG_PAY_OTHERS.

        ENDLOOP.
* Get paid documents for the current line
        LOOP AT IT_PAID_ITEMS INTO LS_PAID_ITEMS_CURR_DOC
            WHERE AUGBL EQ LS_BSEG_PAY_GL-BELNR.
          APPEND LS_PAID_ITEMS_CURR_DOC TO LT_PAID_ITEMS_CURR_DOC.
        ENDLOOP.
*******************G/L Payments****************
        CLEAR: LT_FIAPPTD_TRANS_EXIT.
        TRY.
            CALL BADI OBJ_CBR_HANDLE->PROCESS_OTHER_PAYMENTS
              EXPORTING
                IT_BSEG_BANK_ENTRIES  = LT_BSEG_PAY_BANK
                IT_BSEG_OTHER_ENTRIES = LT_BSEG_PAY_OTHERS
                IT_BSEG_PAID_ENTRIES  = LT_PAID_ITEMS_CURR_DOC
                IT_LFA1               = IT_LFA1
                IT_KNA1               = IT_KNA1
                IS_FIAPPTD_TRANS_EXIT = LS_FIAPPTD_TRANS
              CHANGING
                ET_FIAPPTD_TRANS_EXIT = LT_FIAPPTD_TRANS_EXIT.

          CATCH CX_SY_DYN_CALL_ILLEGAL_METHOD.
        ENDTRY.

        LOOP AT LT_FIAPPTD_TRANS_EXIT INTO LS_FIAPPTD_TRANS_EXIT.

          CALL METHOD CHECK_OLD_REGISTERS
            EXPORTING
              IS_FIAPPTD_TRANS   = LS_FIAPPTD_TRANS_EXIT
              IV_XREVERSAL       = LS_BSEG_PAY_GL-XREVERSAL
            IMPORTING
              EV_REGISTER_EXISTS = LV_REGISTER_EXISTS
              EV_REGISTER_ID     = LV_REGISTER_ID
              EV_ORIGIN_RUN_ID   = LV_ORIGIN_RUN_ID.

          IF LV_REGISTER_EXISTS = 'X' AND LV_REGISTER_ID IS INITIAL.
            "Register already reported and no change so far
            CONTINUE.
          ELSEIF LV_REGISTER_ID = 'C'.
*  Add the next sequence to the run id
            GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
            LV_RUNID_SEQ = GV_RUNID_SEQ.
            SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
            CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_TRANS_EXIT-ORIGIN_RUN_ID.
            LS_FIAPPTD_TRANS_EXIT-REGISTER_ID = LV_REGISTER_ID.
          ELSE.
            LS_FIAPPTD_TRANS_EXIT-REGISTER_ID = LV_REGISTER_ID.
            LS_FIAPPTD_TRANS_EXIT-ORIGIN_RUN_ID = LV_ORIGIN_RUN_ID.
          ENDIF.

          MOVE-CORRESPONDING LS_FIAPPTD_TRANS_EXIT TO LS_XML_ITEMS.

          IF LS_XML_ITEMS-WAERS NE CV_CURR.
            CALL METHOD CONVERT_AMOUNT_FOR_FILE
              EXPORTING
                IV_WAERS  = LS_XML_ITEMS-WAERS
                IV_AMOUNT = LS_XML_ITEMS-WRBTR
              IMPORTING
                EV_AMOUNT = LS_XML_ITEMS-WRBTR.
          ENDIF.

          CLEAR LS_TCURC.
          READ TABLE GT_TCURC INTO LS_TCURC WITH KEY
            WAERS = LS_XML_ITEMS-WAERS BINARY SEARCH.
          IF SY-SUBRC = 0 AND LS_TCURC-ISOCD IS NOT INITIAL.
            LS_XML_ITEMS-ISOCD = LS_TCURC-ISOCD.
            LS_FIAPPTD_TRANS_EXIT-ISOCD = LS_TCURC-ISOCD.
          ENDIF.

          APPEND LS_XML_ITEMS TO ET_XML_ITEMS.
          APPEND LS_FIAPPTD_TRANS_EXIT TO ET_FIAPPTD_TRANS.

        ENDLOOP.

      ELSE.
        IF LV_PAY_WO_INV = 'X'.

          CALL METHOD ME->PROCESS_PAY_WO_INV
            EXPORTING
              IS_BSEG_PAY_GL   = LS_BSEG_PAY_GL
              IS_BSEG_PAY_VC   = LS_BSEG_PAY_VC
              IT_PAID_ITEMS    = IT_PAID_ITEMS
              IT_LFA1          = IT_LFA1
              IT_KNA1          = IT_KNA1
              IS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS
              IT_BSEG_REM_ALL  = LT_BSEG_REM_ALL_WO_INV
            CHANGING
              CT_FIAPPTD_TRANS = ET_FIAPPTD_TRANS
              CT_XML_ITEMS     = ET_XML_ITEMS.

        ENDIF.

        IF LV_PAY_WITH_INV = 'X'.
* Vendor/customer payment with one bank per payment document
          IF LS_BSEG_PAY_VC-TCODE NE 'F110' OR IV_AUTO_TO_MAN EQ 'X'.
* Manual payments - get information from individual invoices.
* iv_auto_to_man EQ 'X' -> Process automatic payments as manual payments

          CALL METHOD PROCESS_INDV_INVOICES
            EXPORTING
              IT_PAID_ITEMS    = IT_PAID_ITEMS
              IS_BSEG_PAY_GL   = LS_BSEG_PAY_GL
              IS_BSEG_PAY_VC   = LS_BSEG_PAY_VC
              IS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS
              IT_LFA1          = IT_LFA1
              IT_KNA1          = IT_KNA1
                IT_BSEG_REM_ALL  = LT_BSEG_REM_ALL_WITH_INV
              IT_BSEG_PARTIAL  = LT_BSEG_PARTIAL
              IT_BSEG_OTHER_GL = LT_BSEG_OTHER_GL
              IV_AUTO_TO_MAN   = IV_AUTO_TO_MAN
            CHANGING
              CT_FIAPPTD_TRANS = ET_FIAPPTD_TRANS
              CT_XML_ITEMS     = ET_XML_ITEMS.

        ELSE.              "Automatic payments.
* For automatic payments - the lines in the payment document will be processed

            CALL METHOD PROCESS_AUTO_PAYMENTS
            EXPORTING
              IS_BSEG_PAY_GL   = LS_BSEG_PAY_GL
              IS_BSEG_PAY_VC   = LS_BSEG_PAY_VC
              IS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS
              IT_PAID_ITEMS    = IT_PAID_ITEMS
              IT_LFA1          = IT_LFA1
              IT_KNA1          = IT_KNA1
                IT_BSEG_REM_ALL  = LT_BSEG_REM_ALL_WITH_INV
                 IT_BSEG_OTHER_GL = LT_BSEG_OTHER_GL
            CHANGING
              CT_FIAPPTD_TRANS = ET_FIAPPTD_TRANS
              CT_XML_ITEMS     = ET_XML_ITEMS.
        ENDIF.

      ENDIF.

      ENDIF.

      DELETE IT_BSEG_PAY_GL WHERE BUKRS = LS_BSEG_PAY_GL-BUKRS AND
      BELNR = LS_BSEG_PAY_GL-BELNR AND GJAHR = LS_BSEG_PAY_GL-GJAHR
      AND BUZEI = LS_BSEG_PAY_GL-BUZEI AND BLDAT = LS_BSEG_PAY_GL-BLDAT
      AND AUGDT = LS_BSEG_PAY_GL-AUGDT AND AUGCP = LS_BSEG_PAY_GL-AUGCP
      AND AUGBL = LS_BSEG_PAY_GL-AUGBL.

      CLEAR LS_BSEG_PAY_GL.
    ENDLOOP.

  ENDMETHOD.