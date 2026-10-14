  METHOD PROCESS_AUTO_PAYMENTS.

    DATA: LS_FIAPPTD_TRANS LIKE LINE OF CT_FIAPPTD_TRANS,
          LS_PAID_ITEMS LIKE LINE OF IT_PAID_ITEMS,
          LV_COUNTER_PAID TYPE I.
    DATA: LV_DUEDATE TYPE DATUM.
    DATA: LV_BANKL TYPE BANKL,
          LV_INTCA3 TYPE INTCA3.
    DATA: LV_REGISTER_EXISTS TYPE FLAG,
          LV_REGISTER_ID TYPE FIAPPT_REG_ID.
    DATA: LV_ORIGIN_RUN_ID TYPE FIAPPT_RUN_ID.
    DATA: LV_NPC2 TYPE CHAR20,
          LV_PAIS_ATIVO TYPE INTCA3,
          LV_ID_BANCO TYPE FIAPPT_BANKID,
          LV_PAIS_CONTA TYPE FIAPPT_PAIS_CONTA,
          LV_SCB_IND    TYPE FIAPPT_SCB_IND,
          LV_STCEG      TYPE STCEG,
          LV_PAY_DOC_PROCESS TYPE FLAG.
    DATA: OBJ_CBR_HANDLE TYPE REF TO BADI_RFFOPT_CBR.

    DATA: LS_XML_ITEMS TYPE FIAPPT_S_CBR_XML_ITEMS.
    DATA: LV_AMOUNT(16) TYPE C.

    DATA: LS_BSEG_ALL TYPE FIAPPT_S_TRANS_CBR,
          LS_BSEG_PAY_GL TYPE FIAPPT_S_TRANS_CBR,
          LS_BSEG_PAY_VC TYPE FIAPPT_S_TRANS_CBR.

    DATA: LV_LAND TYPE LAND1,
          LS_LFA1 TYPE LFA1,
          LS_KNA1 TYPE KNA1,
          LV_LAND_ISO TYPE INTCA3,
          LV_FOREIGN_TRANS TYPE FLAG,
          LS_TCURC TYPE TCURC.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID,
          OBSERVATIONS TYPE STRING.

    CONSTANTS: CV_CURR TYPE WAERS VALUE 'EUR'.

**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.

    LS_BSEG_PAY_GL = IS_BSEG_PAY_GL.
    LS_BSEG_PAY_VC = IS_BSEG_PAY_VC.
*************************
    LS_FIAPPTD_TRANS = IS_FIAPPTD_TRANS.

    IF LS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I' . "AND
*        is_bseg_pay_gl-waers EQ cv_curr.

      CLEAR: LV_FOREIGN_TRANS.

      CALL METHOD ME->CHECK_FOREIGN_TRANS
        EXPORTING
          IT_LFA1          = IT_LFA1
          IT_KNA1          = IT_KNA1
          IS_BSEG_PAY_VC   = LS_BSEG_PAY_VC
          IS_BSEG_PAY_GL   = IS_BSEG_PAY_GL
        IMPORTING
          EV_FOREIGN_TRANS = LV_FOREIGN_TRANS
          EV_LAND          = LV_LAND
          ES_LFA1          = LS_LFA1
          ES_KNA1          = LS_KNA1
          EV_STCEG         = LV_STCEG
        CHANGING
          CS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS.
    ELSE.
      LV_FOREIGN_TRANS = 'X'.  "Foreign transaction
    ENDIF.

    CALL BADI OBJ_CBR_HANDLE->SELECT_PAYMENT_DOCUMENT
      EXPORTING
        IT_LFA1         = IT_LFA1
        IT_KNA1         = IT_KNA1
        IS_BSEG_PAY_VC  = LS_BSEG_PAY_VC
        IS_BSEG_PAY_GL  = IS_BSEG_PAY_GL
      CHANGING
        CV_FOREIGN_TRAN = LV_FOREIGN_TRANS.

    IF LV_FOREIGN_TRANS = 'X'.
*   If an external account then, get the vendor/customer details
      IF LS_BSEG_PAY_VC-KOART = 'K'  AND LS_BSEG_PAY_VC-LIFNR IS NOT INITIAL
        AND LS_LFA1 IS INITIAL.

        READ TABLE IT_LFA1 INTO LS_LFA1 WITH KEY LIFNR = LS_BSEG_PAY_VC-LIFNR.
        IF SY-SUBRC NE 0.
          SELECT SINGLE * FROM LFA1 INTO LS_LFA1 WHERE
            LIFNR = LS_BSEG_PAY_VC-LIFNR.
        ENDIF.
        LV_LAND = LS_LFA1-LAND1.
        LS_FIAPPTD_TRANS-ACCOUNT = LS_LFA1-LIFNR.

        CALL METHOD ME->GET_VAT_REG_NO
          EXPORTING
            IV_STCEG = LS_LFA1-STCEG
          IMPORTING
            EV_STCEG = LV_STCEG.

      ELSEIF LS_BSEG_PAY_VC-KOART = 'D'  AND LS_BSEG_PAY_VC-KUNNR IS NOT INITIAL
        AND LS_KNA1 IS INITIAL.

        READ TABLE IT_KNA1 INTO LS_KNA1 WITH KEY KUNNR = LS_BSEG_PAY_VC-KUNNR.
        IF SY-SUBRC NE 0.
          SELECT SINGLE * FROM KNA1 INTO LS_KNA1 WHERE
           KUNNR = LS_BSEG_PAY_VC-KUNNR.
        ENDIF.
        LV_LAND = LS_KNA1-LAND1.
        LS_FIAPPTD_TRANS-ACCOUNT = LS_KNA1-KUNNR.
        CALL METHOD ME->GET_VAT_REG_NO
          EXPORTING
            IV_STCEG = LS_KNA1-STCEG
          IMPORTING
            EV_STCEG = LV_STCEG.

      ENDIF.
    ENDIF.

    IF LV_FOREIGN_TRANS IS NOT INITIAL.

      CLEAR LV_SCB_IND.
      IF IS_BSEG_PAY_VC-LZBKZ IS NOT INITIAL.
        CALL METHOD GET_SCB_IND
          EXPORTING
            IV_LZBKZ   = IS_BSEG_PAY_VC-LZBKZ
          IMPORTING
            EV_SCB_IND = LV_SCB_IND.
      ELSE.
        "SCB is blank -need to show in the ALV
      ENDIF.


      LS_FIAPPTD_TRANS-PAY_DOC = IS_BSEG_PAY_GL-BELNR.
      LS_FIAPPTD_TRANS-PAY_DOC_YEAR = IS_BSEG_PAY_GL-GJAHR.

      CALL METHOD READ_T005
        EXPORTING
          IV_LAND   = LV_LAND
        IMPORTING
          EV_INTCA3 = LV_LAND_ISO.
      LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE = LV_LAND_ISO.

      IF IS_BSEG_PAY_VC-HBKID IS NOT INITIAL.
        CALL METHOD GET_BANK_DATA
          EXPORTING
            IV_HBKID = IS_BSEG_PAY_VC-HBKID
          IMPORTING
            EV_BANKL = LV_BANKL
            EV_BANKS = LV_INTCA3.
      ENDIF.
      IF IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'.
        LS_FIAPPTD_TRANS-BANK_ID = LV_BANKL.
      ENDIF.

      LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_ALL WHERE BUKRS = GV_BUKRS
         AND BELNR = IS_BSEG_PAY_GL-BELNR  AND GJAHR = GV_YEAR.
*   Processing all the payment document lines

        CLEAR LS_TCURC.
        IF LS_BSEG_ALL-KOART = 'S' AND
          LS_BSEG_ALL-HKONT IN GT_OTH_ACC_RANGE[] AND
          GT_OTH_ACC_RANGE IS NOT INITIAL.
          CONTINUE.
        ENDIF.

          IF LS_BSEG_ALL-WRBTR IS INITIAL.
            CONTINUE.
          ENDIF.

*   One time vendor/customer get the VAT Registration Number
*   from the invoice document
          IF LS_LFA1-XCPDK = 'X' AND LS_LFA1-LIFNR IS NOT INITIAL.
            CALL METHOD ME->GET_VAT_REG_NO
              EXPORTING
                IV_STCEG = IS_BSEG_PAY_VC-STCEG
              IMPORTING
                EV_STCEG = LV_STCEG.

          ELSEIF LS_KNA1-XCPDK = 'X' AND LS_KNA1-KUNNR IS NOT INITIAL.

            CALL METHOD ME->GET_VAT_REG_NO
              EXPORTING
                IV_STCEG = IS_BSEG_PAY_VC-STCEG
              IMPORTING
                EV_STCEG = LV_STCEG.

          ENDIF.

*****User Exit here******************************************
*available details are:
*1. is_bseg_pay_gl - G/L line of the payment document.
*2. is_bseg_pay_vc - Vendor Customer line of the payment document
          TRY.
              CALL BADI OBJ_CBR_HANDLE->PROCESS_AUTO_PAYMENTS
                EXPORTING
                  LFA1          = LS_LFA1
                  KNA1          = LS_KNA1
                  BSEG_PAY_GL   = IS_BSEG_PAY_GL
                  BSEG_PAY_VC   = LS_BSEG_ALL
*               paid_items    = ls_paid_items
                  FIAPPTD_TRANS = LS_FIAPPTD_TRANS
                CHANGING
                  NPC2          = LV_NPC2
                  PAIS_ATIVO    = LV_PAIS_ATIVO
                  ID_BANCO      = LV_ID_BANCO
                CV_STCEG      = LV_STCEG
                PAIS_CONTA    = LV_PAIS_CONTA.

            CATCH CX_SY_DYN_CALL_ILLEGAL_METHOD.
          ENDTRY.

          LS_FIAPPTD_TRANS-PAIS_ATIVO = LV_PAIS_ATIVO.
          LS_FIAPPTD_TRANS-NPC2 = LV_NPC2.
          LS_FIAPPTD_TRANS-SCB_IND = LV_SCB_IND.
          LS_FIAPPTD_TRANS-STCEG = LV_STCEG.

          IF LV_ID_BANCO IS NOT INITIAL AND
            IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'.
            LS_FIAPPTD_TRANS-BANK_ID = LV_ID_BANCO.
            "Value from user exit for external accounts
          ENDIF.
*          IF ls_fiapptd_trans-bank_acc_type EQ 'E'.
            "to be filled only when account type is external.
        IF LV_PAIS_CONTA IS INITIAL.
          LS_FIAPPTD_TRANS-PAIS_CONTA = LV_INTCA3.
        ELSE.
          LS_FIAPPTD_TRANS-PAIS_CONTA = LV_PAIS_CONTA.
        ENDIF.
*          ENDIF.

          LS_FIAPPTD_TRANS-DOC_YEAR = LS_BSEG_ALL-GJAHR.
          LS_FIAPPTD_TRANS-BELNR = LS_BSEG_ALL-BELNR.
          LS_FIAPPTD_TRANS-BUZEI = LS_BSEG_ALL-BUZEI.
          LS_FIAPPTD_TRANS-KOART = LS_BSEG_ALL-KOART.
          IF LS_BSEG_ALL-KOART = 'K'.
            LS_FIAPPTD_TRANS-ACCOUNT = LS_BSEG_ALL-LIFNR.
          ELSEIF LS_BSEG_ALL-KOART = 'D'.
            LS_FIAPPTD_TRANS-ACCOUNT = LS_BSEG_ALL-KUNNR.
          ELSEIF  LS_BSEG_ALL-KOART = 'S'.
            LS_FIAPPTD_TRANS-ACCOUNT = LS_BSEG_ALL-HKONT.
            LS_FIAPPTD_TRANS-STCEG = ''.
*            ls_fiapptd_trans-pais_contraparte = ''.
          ENDIF.

          IF LS_BSEG_ALL-KOART EQ 'S'.
            LS_FIAPPTD_TRANS-ACCOUNT = LS_BSEG_ALL-HKONT.
            LS_FIAPPTD_TRANS-STCEG = ''.
*            ls_fiapptd_trans-pais_contraparte = ''.
           IF LS_BSEG_ALL-LZBKZ IS NOT INITIAL.
             CALL METHOD GET_SCB_IND
             EXPORTING
               IV_LZBKZ   = LS_BSEG_ALL-LZBKZ
             IMPORTING
               EV_SCB_IND = LS_FIAPPTD_TRANS-SCB_IND.
           ENDIF.
          ENDIF.

          IF GV_PERIOD_1ST_RUN NE 'X'.
*Not the first run for the reporting period - check
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

          ELSE.    "First run for the specified reporting month
            LS_FIAPPTD_TRANS-REGISTER_ID = 'C'.
*  Add the next sequence to the run id
            GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
            LV_RUNID_SEQ = GV_RUNID_SEQ.
            SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
            CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_TRANS-ORIGIN_RUN_ID.
          ENDIF.

          LS_FIAPPTD_TRANS-REFERENCE_DATE = LS_BSEG_ALL-BUDAT.
          LS_FIAPPTD_TRANS-WRBTR = LS_BSEG_ALL-WRBTR - LS_BSEG_ALL-WSKTO - LS_BSEG_ALL-QBSHB.
          "amount from gl-account line as there can be multiple vc lines
          LS_FIAPPTD_TRANS-WAERS = LS_BSEG_ALL-WAERS.

          IF LS_BSEG_ALL-SHKZG EQ 'S'.
            LS_FIAPPTD_TRANS-VALUE_TYPE = 'S'. "Exit
          ELSE.
            LS_FIAPPTD_TRANS-VALUE_TYPE = 'E'. "entry
          ENDIF.

          LS_FIAPPTD_TRANS-BANK_ACC_TYPE = IS_BSEG_PAY_GL-BANK_ACC_TYPE.

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

      ENDLOOP.

    LV_PAY_DOC_PROCESS = 'X' .
    CALL METHOD ME->PROCESS_OTHER_LINES
      EXPORTING
*       it_paid_items      =
        IS_BSEG_PAY_GL     = IS_BSEG_PAY_GL
        IS_BSEG_PAY_VC     = IS_BSEG_PAY_VC
        IT_LFA1            = IT_LFA1
        IT_KNA1            = IT_KNA1
        IS_FIAPPTD_TRANS   = LS_FIAPPTD_TRANS
        IT_BSEG_REM_ALL    = IT_BSEG_OTHER_GL
        IV_PAY_DOC_PROCESS = LV_PAY_DOC_PROCESS
        IV_LAND_ISO        = LV_LAND_ISO
      CHANGING
        CT_FIAPPTD_TRANS   = CT_FIAPPTD_TRANS
        CT_XML_ITEMS       = CT_XML_ITEMS.

    ENDIF.

  ENDMETHOD.