  METHOD PROCESS_PARTIAL_ITEMS.

    DATA: LV_LAND TYPE LAND1,
          LS_KNA1 TYPE KNA1,
          LS_LFA1 TYPE LFA1.

    DATA: LS_FIAPPTD_TRANS TYPE FIAPPTD_TRANS.
    DATA: LV_NPC2 TYPE CHAR20,
          LV_PAIS_ATIVO TYPE INTCA3,
          LV_ID_BANCO TYPE FIAPPT_BANKID,
          LV_REGISTER_EXISTS TYPE FLAG,
          LV_REGISTER_ID TYPE FIAPPT_REG_ID,
          LV_ORIGIN_RUN_ID TYPE FIAPPT_RUN_ID,
          LV_DUEDATE TYPE DATUM,
          LV_LAND_ISO TYPE INTCA3,
          LV_STCEG TYPE STCEG.


    DATA: OBJ_CBR_HANDLE TYPE REF TO BADI_RFFOPT_CBR.
    DATA: LS_XML_ITEMS TYPE FIAPPT_S_CBR_XML_ITEMS.

    DATA: LS_BSEG_ALL TYPE FIAPPT_S_TRANS_CBR,
          LV_FOREIGN_TRANS TYPE FLAG,
          LS_BSEG_PARTIAL TYPE FIAPPT_S_TRANS_CBR,
          LV_AMT_PAY_CURR TYPE WRBTR,
          LS_TCURC TYPE TCURC.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID,
          OBSERVATIONS TYPE STRING.

    CONSTANTS: CV_CURR TYPE WAERS VALUE 'EUR'.
**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.

    LS_FIAPPTD_TRANS = IS_FIAPPTD_TRANS.

*   This loop will check in the payment document for partially
*   paid items and for the G/L lines which are posted to the
*   G/L's maintained under the ID 'OTH' in the FSV and to
*   report the same.

    SORT IT_BSEG_REM_ALL BY BUKRS BELNR GJAHR REBZG REBZJ.

    LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_ALL WHERE
      BUKRS = GV_BUKRS AND BELNR = IS_BSEG_PAY_GL-BELNR
          AND GJAHR = GV_YEAR AND REBZG IS NOT INITIAL AND
           REBZJ IS NOT INITIAL.
      CLEAR LS_TCURC.

      IF LS_BSEG_ALL-KOART = 'K' OR LS_BSEG_ALL-KOART = 'D'.
*      Process Partial Payments
        SELECT SINGLE BUKRS BELNR GJAHR BUZEI AUGDT AUGBL KOART BSCHL
          SHKZG DMBTR WRBTR SGTXT HKONT LIFNR KUNNR ZFBDT
          ZTERM ZBD1T ZBD2T ZBD1P ZBD2P WSKTO HBKID LZBKZ LANDL
          STCEG XRAGL HKTID FROM BSEG INTO
          CORRESPONDING FIELDS OF LS_BSEG_PARTIAL
          WHERE BUKRS = GV_BUKRS
          AND BELNR = LS_BSEG_ALL-REBZG
          AND GJAHR =  LS_BSEG_ALL-REBZJ
          AND BUZEI = LS_BSEG_ALL-REBZZ.

        IF SY-SUBRC = 0.

          IF IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'. " AND
*                  is_bseg_pay_gl-waers EQ cv_curr.

            CALL METHOD ME->CHECK_FOREIGN_TRANS
              EXPORTING
                IT_LFA1          = IT_LFA1
                IT_KNA1          = IT_KNA1
                IS_BSEG_PAY_VC   = LS_BSEG_ALL
                IS_BSEG_PAY_GL   = IS_BSEG_PAY_GL
              IMPORTING
                EV_FOREIGN_TRANS = LV_FOREIGN_TRANS
                EV_LAND          = LV_LAND
                ES_LFA1          = LS_LFA1
                ES_KNA1          = LS_KNA1
                EV_STCEG         = LV_STCEG
              CHANGING
                CS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS.

*            IF lv_foreign_trans NE 'X'.
*              CONTINUE.
*            ENDIF.
          ELSE.
            LV_FOREIGN_TRANS = 'X'.
          ENDIF.

          CALL BADI OBJ_CBR_HANDLE->SELECT_PAYMENT_DOCUMENT
            EXPORTING
              IT_LFA1         = IT_LFA1
              IT_KNA1         = IT_KNA1
              IS_BSEG_PAY_VC  = LS_BSEG_ALL
              IS_BSEG_PAY_GL  = IS_BSEG_PAY_GL
            CHANGING
              CV_FOREIGN_TRAN = LV_FOREIGN_TRANS.

          IF LV_FOREIGN_TRANS NE 'X'.
             CONTINUE.
          ENDIF.

          IF LV_FOREIGN_TRANS EQ 'X'.

            IF LS_BSEG_ALL-KOART = 'K' AND LS_BSEG_ALL-LIFNR IS NOT INITIAL
              AND LS_LFA1 IS INITIAL.

              READ TABLE IT_LFA1 INTO LS_LFA1 WITH KEY LIFNR = LS_BSEG_ALL-LIFNR.
              IF SY-SUBRC NE 0.
                SELECT SINGLE * FROM LFA1 INTO LS_LFA1 WHERE
                  LIFNR = LS_BSEG_ALL-LIFNR.
              ENDIF.
              LV_LAND = LS_LFA1-LAND1.
              LS_FIAPPTD_TRANS-ACCOUNT = LS_LFA1-LIFNR.
              CALL METHOD ME->GET_VAT_REG_NO
                EXPORTING
                  IV_STCEG = LS_LFA1-STCEG
                IMPORTING
                  EV_STCEG = LV_STCEG.
            ELSEIF LS_BSEG_ALL-KOART = 'D' AND LS_BSEG_ALL-KUNNR IS NOT INITIAL
              AND LS_KNA1 IS INITIAL.

              READ TABLE IT_KNA1 INTO LS_KNA1 WITH KEY KUNNR = LS_BSEG_ALL-KUNNR.
              IF SY-SUBRC NE 0.
                SELECT SINGLE * FROM KNA1 INTO LS_KNA1 WHERE
                  KUNNR = LS_BSEG_ALL-KUNNR.
              ENDIF.
              LV_LAND = LS_KNA1-LAND1.
              LS_FIAPPTD_TRANS-ACCOUNT = LS_KNA1-KUNNR.
              CALL METHOD ME->GET_VAT_REG_NO
                EXPORTING
                  IV_STCEG = LS_KNA1-STCEG
                IMPORTING
                  EV_STCEG = LV_STCEG.

            ENDIF.
* One time accounts
            IF LS_BSEG_ALL-XCPDD IS NOT INITIAL.

              CALL METHOD ME->GET_VAT_REG_NO
                EXPORTING
                  IV_STCEG = LS_BSEG_ALL-STCEG
                IMPORTING
                  EV_STCEG = LV_STCEG.
              IF LS_BSEG_ALL-LANDL IS NOT INITIAL.
                LV_LAND = LS_BSEG_ALL-LANDL.
              ENDIF.
            ENDIF.

          ENDIF.

          CALL METHOD READ_T005
            EXPORTING
              IV_LAND   = LV_LAND
            IMPORTING
              EV_INTCA3 = LV_LAND_ISO.
          LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE = LV_LAND_ISO.


          IF LS_BSEG_PARTIAL-LZBKZ IS NOT INITIAL.
            CALL METHOD GET_SCB_IND
              EXPORTING
                IV_LZBKZ   = LS_BSEG_PARTIAL-LZBKZ
              IMPORTING
                EV_SCB_IND = LS_FIAPPTD_TRANS-SCB_IND.
          ELSE.
            "SCB indicator blank- need to show in the ALV
          ENDIF.

          LS_FIAPPTD_TRANS-BELNR = LS_BSEG_ALL-REBZG.
          LS_FIAPPTD_TRANS-DOC_YEAR = LS_BSEG_ALL-REBZJ.
          LS_FIAPPTD_TRANS-BUZEI = LS_BSEG_ALL-REBZZ.
        IF LS_BSEG_ALL-KOART = 'K'.
           LS_FIAPPTD_TRANS-WRBTR = LS_BSEG_ALL-WRBTR - LS_BSEG_ALL-WSKTO - LS_BSEG_ALL-QBSHB.
        ELSEIF LS_BSEG_ALL-KOART = 'D'.
          LS_FIAPPTD_TRANS-WRBTR = LS_BSEG_ALL-WRBTR - LS_BSEG_ALL-WSKTO .
        ENDIF.

          LS_FIAPPTD_TRANS-WAERS = LS_BSEG_ALL-WAERS.
          LS_FIAPPTD_TRANS-KOART = LS_BSEG_ALL-KOART.
          LS_FIAPPTD_TRANS-REFERENCE_DATE = IS_BSEG_PAY_GL-BUDAT.

          CALL METHOD CALC_DUEDATE           "duedate
            EXPORTING
              IV_ZFBDT   = LS_BSEG_ALL-ZFBDT
              IV_ZBD1T   = LS_BSEG_ALL-ZBD1T
              IV_ZBD2T   = LS_BSEG_ALL-ZBD2T
              IV_ZBD3T   = LS_BSEG_ALL-ZBD3T
              IV_ZBD1P   = LS_BSEG_ALL-ZBD1P
              IV_ZBD2P   = LS_BSEG_ALL-ZBD2P
              IV_WSKTO   = LS_BSEG_ALL-WSKTO
            IMPORTING
              EV_DUEDATE = LV_DUEDATE.
*          ls_fiapptd_trans-due_date = lv_duedate.


          IF LS_BSEG_ALL-SHKZG EQ 'S'.
            LS_FIAPPTD_TRANS-VALUE_TYPE = 'S'. "exit
          ELSE.
            LS_FIAPPTD_TRANS-VALUE_TYPE = 'E'. "entry
          ENDIF.


*one time vendor/customer get the vat registration number
*from the invoice document
          IF LS_LFA1-XCPDK = 'X' AND LS_LFA1-LIFNR IS NOT INITIAL.
            CALL METHOD ME->GET_VAT_REG_NO
              EXPORTING
                IV_STCEG = LS_BSEG_PARTIAL-STCEG
              IMPORTING
                EV_STCEG = LV_STCEG.

          ELSEIF LS_KNA1-XCPDK = 'X' AND LS_KNA1-KUNNR IS NOT INITIAL.
            CALL METHOD ME->GET_VAT_REG_NO
              EXPORTING
                IV_STCEG = LS_BSEG_PARTIAL-STCEG
              IMPORTING
                EV_STCEG = LV_STCEG.

          ENDIF.
*****User Exit here******************************************
          "ls_fiapptd_trans-npc2 =
          "ls_fiapptd_trans-pais_ativo =

          TRY.
              CALL BADI OBJ_CBR_HANDLE->PROCESS_INDV_INVOICES
                EXPORTING
                  LFA1          = LS_LFA1
                  KNA1          = LS_KNA1
                  BSEG_PAY_GL   = IS_BSEG_PAY_GL
                  BSEG_PAY_VC   = LS_BSEG_ALL
                  PAID_ITEMS    = LS_BSEG_PARTIAL
                  FIAPPTD_TRANS = LS_FIAPPTD_TRANS
                CHANGING
                  NPC2          = LV_NPC2
                  PAIS_ATIVO    = LV_PAIS_ATIVO
                  ID_BANCO      = LV_ID_BANCO
                  CV_STCEG      = LV_STCEG.

            CATCH CX_SY_DYN_CALL_ILLEGAL_METHOD.
          ENDTRY.

          IF LV_ID_BANCO IS NOT INITIAL AND
            IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'.
            LS_FIAPPTD_TRANS-BANK_ID = LV_ID_BANCO.
          ENDIF.
          LS_FIAPPTD_TRANS-NPC2 = LV_NPC2.
          LS_FIAPPTD_TRANS-PAIS_ATIVO = LV_PAIS_ATIVO.
          LS_FIAPPTD_TRANS-STCEG = LV_STCEG.

        ENDIF.

        IF GV_PERIOD_1ST_RUN NE 'X'.
*Not the first run for the reporting period - need to check
*existing registers in DB
          CALL METHOD CHECK_OLD_REGISTERS
            EXPORTING
              IS_FIAPPTD_TRANS   = LS_FIAPPTD_TRANS
              IV_XREVERSAL       = IS_BSEG_PAY_GL-XREVERSAL
            IMPORTING
              EV_REGISTER_EXISTS = LV_REGISTER_EXISTS
              EV_REGISTER_ID     = LV_REGISTER_ID
              EV_ORIGIN_RUN_ID   = LV_ORIGIN_RUN_ID.

          IF LV_REGISTER_EXISTS = 'X' AND LV_REGISTER_ID IS INITIAL.
            "Register already reported and no change so far
            CV_PAY_DOC_PROCESS = 'X'.
            CONTINUE.

          ELSEIF LV_REGISTER_ID = 'C'.
*  Add the next sequence to the run id
            GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
            LV_RUNID_SEQ = GV_RUNID_SEQ.
            SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
            CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_TRANS-ORIGIN_RUN_ID.
            LS_FIAPPTD_TRANS-REGISTER_ID = LV_REGISTER_ID.
          ELSE.
            LS_FIAPPTD_TRANS-REGISTER_ID = LV_REGISTER_ID.
            LS_FIAPPTD_TRANS-ORIGIN_RUN_ID = LV_ORIGIN_RUN_ID.
          ENDIF.

        ELSE.
          "First run for the specified reporting month
          LS_FIAPPTD_TRANS-REGISTER_ID = 'C'.
*   Add the next sequence to the run id
          GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
          LV_RUNID_SEQ = GV_RUNID_SEQ.
          SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
          CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_TRANS-ORIGIN_RUN_ID.
        ENDIF.

        CONCATENATE LS_FIAPPTD_TRANS-PAY_DOC '/' LS_FIAPPTD_TRANS-PAY_DOC_YEAR
            '/' GV_BUKRS INTO OBSERVATIONS.

        LS_FIAPPTD_TRANS-OBSERVATIONS = OBSERVATIONS.

        MOVE-CORRESPONDING LS_FIAPPTD_TRANS TO LS_XML_ITEMS.

        IF LS_XML_ITEMS-WAERS NE CV_CURR.
          CALL METHOD CONVERT_AMOUNT_FOR_FILE
            EXPORTING
              IV_WAERS  = LS_XML_ITEMS-WAERS
              IV_AMOUNT = LS_XML_ITEMS-WRBTR
            IMPORTING
              EV_AMOUNT = LS_XML_ITEMS-WRBTR.
        ENDIF.

        READ TABLE GT_TCURC INTO LS_TCURC WITH KEY
          WAERS = LS_XML_ITEMS-WAERS BINARY SEARCH.
        IF SY-SUBRC = 0 AND LS_TCURC-ISOCD IS NOT INITIAL.
          LS_XML_ITEMS-ISOCD = LS_TCURC-ISOCD.
          LS_FIAPPTD_TRANS-ISOCD = LS_TCURC-ISOCD.
        ENDIF.

        APPEND LS_XML_ITEMS TO CT_XML_ITEMS.
        APPEND LS_FIAPPTD_TRANS TO CT_FIAPPTD_TRANS.
        CV_PAY_DOC_PROCESS = 'X'.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.