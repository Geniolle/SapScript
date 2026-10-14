  METHOD PROCESS_INDV_INVOICES.

    DATA: LS_PAID_ITEMS          LIKE LINE OF IT_PAID_ITEMS,
          LT_PAID_ITEMS          LIKE IT_PAID_ITEMS,
          LS_FIAPPTD_TRANS       LIKE LINE OF CT_FIAPPTD_TRANS,
          LS_EXIST_FIAPPTD_TRANS LIKE LINE OF CT_FIAPPTD_TRANS.
    DATA: LV_DUEDATE TYPE DATUM.
    DATA: LV_BANKL  TYPE BANKL,
          LV_INTCA3 TYPE INTCA3.
    DATA: LV_REGISTER_EXISTS TYPE FLAG,
          LV_REGISTER_ID     TYPE FIAPPT_REG_ID.
    DATA: LV_ORIGIN_RUN_ID TYPE FIAPPT_RUN_ID.
    DATA: LV_HBKID TYPE HBKID.
    DATA: LV_NPC2       TYPE CHAR20,
          LV_PAIS_ATIVO TYPE INTCA3,
          LV_ID_BANCO   TYPE FIAPPT_BANKID.
    DATA: OBJ_CBR_HANDLE TYPE REF TO BADI_RFFOPT_CBR,
          LS_TRANS_INIT  TYPE FIAPPTD_TRANS.

    DATA: LS_XML_ITEMS TYPE FIAPPT_S_CBR_XML_ITEMS.

    DATA: LV_AMOUNT(16) TYPE C.
    DATA: LV_LIFNR TYPE LIFNR, LV_KUNNR TYPE KUNNR.
    DATA: LV_LAND     TYPE LAND1, LV_LAND_ISO TYPE INTCA3.

    DATA: LS_LFA1 TYPE LFA1, LS_KNA1 TYPE KNA1.

    DATA: LS_BSEG_ALL        TYPE FIAPPT_S_TRANS_CBR,
          LV_FOREIGN_TRANS   TYPE FLAG,
          LS_BSEG_PARTIAL    TYPE FIAPPT_S_TRANS_CBR,
          LV_AMT_PAY_CURR    TYPE WRBTR,
          LV_AMT_LOC_CURR    TYPE WRBTR,
          LV_LOC_CURR        TYPE WAERS,
          LV_PAY_DOC_PROCESS TYPE FLAG,
          LS_TCURC           TYPE TCURC,
          LV_STCEG           TYPE STCEG.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID,
          OBSERVATIONS TYPE STRING.

    CONSTANTS: CV_CURR TYPE WAERS VALUE 'EUR'.

**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.

    LS_FIAPPTD_TRANS = IS_FIAPPTD_TRANS.
*    ls_fiapptd_trans-doc_type = 'I'.

    LS_FIAPPTD_TRANS-PAY_DOC = IS_BSEG_PAY_GL-BELNR.
    LS_FIAPPTD_TRANS-PAY_DOC_YEAR = IS_BSEG_PAY_GL-GJAHR.
    LS_FIAPPTD_TRANS-BANK_ACC_TYPE = IS_BSEG_PAY_GL-BANK_ACC_TYPE.

    SELECT SINGLE HBKID INTO LV_HBKID FROM SKB1
      WHERE BUKRS = GV_BUKRS
        AND SAKNR = IS_BSEG_PAY_GL-HKONT.
    IF LV_HBKID IS INITIAL.
*      SELECT SINGLE hbkid INTO lv_hbkid FROM t012k
*        WHERE bukrs = gv_bukrs
*        AND   hkont = is_bseg_pay_gl-hkont.

      DATA(LT_T012K) = CL_FCLM_HOUSE_BANK_ACCOUNT=>GET_T012K_T( EXPORTING
                                                                  IT_BUKRS = VALUE #( ( GV_BUKRS ) )
                                                                  IT_HKONT = VALUE #( ( IS_BSEG_PAY_GL-HKONT ) ) ).

      READ TABLE LT_T012K ASSIGNING FIELD-SYMBOL(<FS_T012K>) INDEX 1.
      IF SY-SUBRC EQ 0.
        LV_HBKID = <FS_T012K>-HBKID.
      ENDIF.

    ENDIF.

    IF LV_HBKID IS NOT INITIAL.
      CALL METHOD GET_BANK_DATA
        EXPORTING
          IV_HBKID = LV_HBKID
        IMPORTING
          EV_BANKL = LV_BANKL
          EV_BANKS = LV_INTCA3.
    ELSE.
      CLEAR: LV_BANKL, LV_INTCA3.
    ENDIF.

    IF IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'.
      LS_FIAPPTD_TRANS-BANK_ID = LV_BANKL.
    ENDIF.

*    IF ls_fiapptd_trans-bank_acc_type EQ 'E'.
*      to be filled only when account type is external.
    LS_FIAPPTD_TRANS-PAIS_CONTA = LV_INTCA3.
*    ENDIF.

    LS_TRANS_INIT = LS_FIAPPTD_TRANS.

*    Get the company code currency.
    CALL METHOD ME->READ_T001
      IMPORTING
        EV_COMP_CURR = LV_LOC_CURR
      EXCEPTIONS
        NO_ENTRY     = 1
        OTHERS       = 2.

    IF LV_LOC_CURR IS INITIAL.
      LV_LOC_CURR = CV_CURR.
    ENDIF.

    LT_PAID_ITEMS[] = IT_PAID_ITEMS[].
    SORT LT_PAID_ITEMS BY AUGBL.

    LOOP AT LT_PAID_ITEMS INTO LS_PAID_ITEMS
          WHERE AUGBL EQ IS_BSEG_PAY_GL-BELNR.
* First get the SCB and due date and pass to user exit
* for the fields NPC2 and PAIS_ATIVO

      IF IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'. "AND
*         is_bseg_pay_gl-waers EQ cv_curr.

        CALL METHOD ME->CHECK_FOREIGN_TRANS
          EXPORTING
            IT_LFA1          = IT_LFA1
            IT_KNA1          = IT_KNA1
            IS_BSEG_PAY_VC   = LS_PAID_ITEMS
            IS_BSEG_PAY_GL   = IS_BSEG_PAY_GL
          IMPORTING
            EV_FOREIGN_TRANS = LV_FOREIGN_TRANS
            EV_LAND          = LV_LAND
            ES_LFA1          = LS_LFA1
            ES_KNA1          = LS_KNA1
            EV_STCEG         = LV_STCEG
          CHANGING
            CS_FIAPPTD_TRANS = LS_FIAPPTD_TRANS.

*        IF lv_foreign_trans IS INITIAL.
*          CONTINUE.
*        ENDIF.
      ELSE.
        LV_FOREIGN_TRANS = 'X'.
      ENDIF.

      CALL BADI OBJ_CBR_HANDLE->SELECT_PAYMENT_DOCUMENT
        EXPORTING
          IT_LFA1         = IT_LFA1
          IT_KNA1         = IT_KNA1
          IS_BSEG_PAY_VC  = LS_PAID_ITEMS
          IS_BSEG_PAY_GL  = IS_BSEG_PAY_GL
        CHANGING
          CV_FOREIGN_TRAN = LV_FOREIGN_TRANS.

      IF LV_FOREIGN_TRANS IS INITIAL.
        CONTINUE.
      ENDIF.

      IF LV_FOREIGN_TRANS EQ 'X'.

        IF LS_PAID_ITEMS-LIFNR IS NOT INITIAL
          AND LS_LFA1 IS INITIAL.

          READ TABLE IT_LFA1 INTO LS_LFA1 WITH KEY
                LIFNR = LS_PAID_ITEMS-LIFNR.
          IF SY-SUBRC NE 0.
            SELECT SINGLE * FROM LFA1 INTO LS_LFA1 WHERE
              LIFNR = LS_PAID_ITEMS-LIFNR.
          ENDIF.
          LV_LAND = LS_LFA1-LAND1.
          LS_FIAPPTD_TRANS-ACCOUNT = LS_LFA1-LIFNR.
          CALL METHOD ME->GET_VAT_REG_NO
            EXPORTING
              IV_STCEG = LS_LFA1-STCEG
            IMPORTING
              EV_STCEG = LV_STCEG.

          LS_FIAPPTD_TRANS-KOART = 'K'.

        ELSEIF  LS_PAID_ITEMS-KUNNR IS NOT INITIAL
          AND LS_KNA1 IS INITIAL.

          READ TABLE IT_KNA1 INTO LS_KNA1 WITH KEY KUNNR = LS_PAID_ITEMS-KUNNR.
          IF SY-SUBRC NE 0.
            SELECT SINGLE * FROM KNA1 INTO LS_KNA1 WHERE
             KUNNR = LS_PAID_ITEMS-KUNNR.
          ENDIF.
          LV_LAND = LS_KNA1-LAND1.
          LS_FIAPPTD_TRANS-ACCOUNT = LS_KNA1-KUNNR.

          CALL METHOD ME->GET_VAT_REG_NO
            EXPORTING
              IV_STCEG = LS_KNA1-STCEG
            IMPORTING
              EV_STCEG = LV_STCEG.

          LS_FIAPPTD_TRANS-KOART = 'D'.
        ENDIF.
* One time accounts
        IF LS_PAID_ITEMS-XCPDD IS NOT INITIAL.

          CALL METHOD ME->GET_VAT_REG_NO
            EXPORTING
              IV_STCEG = LS_PAID_ITEMS-STCEG
            IMPORTING
              EV_STCEG = LV_STCEG.
          IF LS_PAID_ITEMS-LANDL IS NOT INITIAL.
            LV_LAND = LS_PAID_ITEMS-LANDL.
          ENDIF.
        ENDIF.

      ENDIF.

      CALL METHOD READ_T005
        EXPORTING
          IV_LAND   = LV_LAND
        IMPORTING
          EV_INTCA3 = LV_LAND_ISO.
      LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE = LV_LAND_ISO.

      IF LS_PAID_ITEMS-LZBKZ IS NOT INITIAL.
        CALL METHOD GET_SCB_IND
          EXPORTING
            IV_LZBKZ   = LS_PAID_ITEMS-LZBKZ
          IMPORTING
            EV_SCB_IND = LS_PAID_ITEMS-SCB_IND.
      ELSE.
        "SCB indicator blank- need to show in the ALV
      ENDIF.
      LS_FIAPPTD_TRANS-SCB_IND = LS_PAID_ITEMS-SCB_IND.

*   One time vendor/customer get the VAT Registration Number
*   from the invoice document
      IF LS_LFA1-XCPDK = 'X' AND LS_LFA1-LIFNR IS NOT INITIAL.
        CALL METHOD ME->GET_VAT_REG_NO
          EXPORTING
            IV_STCEG = LS_PAID_ITEMS-STCEG
          IMPORTING
            EV_STCEG = LV_STCEG.
      ELSEIF LS_KNA1-XCPDK = 'X' AND LS_KNA1-KUNNR IS NOT INITIAL.
        CALL METHOD ME->GET_VAT_REG_NO
          EXPORTING
            IV_STCEG = LS_PAID_ITEMS-STCEG
          IMPORTING
            EV_STCEG = LV_STCEG.
      ENDIF.

      TRY.
          CALL BADI OBJ_CBR_HANDLE->PROCESS_INDV_INVOICES
            EXPORTING
              LFA1          = LS_LFA1
              KNA1          = LS_KNA1
              BSEG_PAY_GL   = IS_BSEG_PAY_GL
              BSEG_PAY_VC   = IS_BSEG_PAY_VC
              PAID_ITEMS    = LS_PAID_ITEMS
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

      LS_FIAPPTD_TRANS-DOC_YEAR = LS_PAID_ITEMS-GJAHR.
      LS_FIAPPTD_TRANS-BUZEI = LS_PAID_ITEMS-BUZEI.
      LS_FIAPPTD_TRANS-BELNR = LS_PAID_ITEMS-BELNR.
      IF LS_FIAPPTD_TRANS-KOART = 'S'.
        LS_FIAPPTD_TRANS-ACCOUNT = LS_PAID_ITEMS-HKONT.
        LS_FIAPPTD_TRANS-STCEG = ''.
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
          LV_PAY_DOC_PROCESS = 'X'.
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

      " Should still be the document date of the payment document
      LS_FIAPPTD_TRANS-REFERENCE_DATE = IS_BSEG_PAY_GL-BUDAT.
      LS_FIAPPTD_TRANS-WRBTR = LS_PAID_ITEMS-WRBTR - LS_PAID_ITEMS-WSKTO - LS_PAID_ITEMS-QBSHB..
      LS_FIAPPTD_TRANS-WAERS = LS_PAID_ITEMS-WAERS.

************************************************************************************
* Convert invoice amt to pay curr here need to write code .....
      IF IS_BSEG_PAY_GL-WAERS NE LS_PAID_ITEMS-WAERS.
* Convert the amount to local currency 'EUR'
*        CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
*          EXPORTING
*            date             = is_bseg_pay_gl-wwert
*            foreign_amount   = ls_paid_items-wrbtr
*            foreign_currency = ls_paid_items-waers
*            local_currency   = lv_loc_curr
**           rate             = is_bseg_pay_gl-kursf
**           type_of_rate     = is_bseg_pay_gl-kurst
*            read_tcurr       = 'X'
*          IMPORTING
*            local_amount     = lv_amt_loc_curr
*          EXCEPTIONS
*            no_rate_found    = 1
*            overflow         = 2
*            no_factors_found = 3
*            no_spread_found  = 4
*            derived_2_times  = 5
*            OTHERS           = 6.

*        IF sy-subrc = 0.
        IF LS_PAID_ITEMS-PYAMT IS INITIAL.
*  Convert amount in local currency to payment currency
          CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
            EXPORTING
              DATE             = IS_BSEG_PAY_GL-WWERT
              FOREIGN_AMOUNT   = LS_PAID_ITEMS-DMBTR
              FOREIGN_CURRENCY = LV_LOC_CURR
              LOCAL_CURRENCY   = IS_BSEG_PAY_GL-WAERS
*             rate             = is_bseg_pay_gl-kursf
*             type_of_rate     = is_bseg_pay_gl-kurst
              READ_TCURR       = 'X'
            IMPORTING
              LOCAL_AMOUNT     = LV_AMT_PAY_CURR
            EXCEPTIONS
              NO_RATE_FOUND    = 1
              OVERFLOW         = 2
              NO_FACTORS_FOUND = 3
              NO_SPREAD_FOUND  = 4
              DERIVED_2_TIMES  = 5
              OTHERS           = 6.

*        ENDIF.

          IF SY-SUBRC <> 0.
* Implement suitable error handling here
          ENDIF.

          LS_FIAPPTD_TRANS-WRBTR = LV_AMT_PAY_CURR.
          LS_FIAPPTD_TRANS-WAERS = IS_BSEG_PAY_GL-WAERS.
        ELSE.

          IF IS_BSEG_PAY_GL-WAERS = LS_PAID_ITEMS-PYCUR.
            LS_FIAPPTD_TRANS-WRBTR = LS_PAID_ITEMS-PYAMT.
            LS_FIAPPTD_TRANS-WAERS = IS_BSEG_PAY_GL-WAERS.

          ELSEIF IS_BSEG_PAY_GL-WAERS EQ LV_LOC_CURR.
            LS_FIAPPTD_TRANS-WRBTR = LS_PAID_ITEMS-DMBTR.
            LS_FIAPPTD_TRANS-WAERS = IS_BSEG_PAY_GL-WAERS.

          ELSE.

            CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
              EXPORTING
                DATE             = IS_BSEG_PAY_GL-WWERT
                FOREIGN_AMOUNT   = LS_PAID_ITEMS-DMBTR
                FOREIGN_CURRENCY = LV_LOC_CURR
                LOCAL_CURRENCY   = IS_BSEG_PAY_GL-WAERS
*               rate             = is_bseg_pay_gl-kursf
*               type_of_rate     = is_bseg_pay_gl-kurst
                READ_TCURR       = 'X'
              IMPORTING
                LOCAL_AMOUNT     = LV_AMT_PAY_CURR
              EXCEPTIONS
                NO_RATE_FOUND    = 1
                OVERFLOW         = 2
                NO_FACTORS_FOUND = 3
                NO_SPREAD_FOUND  = 4
                DERIVED_2_TIMES  = 5
                OTHERS           = 6.

            LS_FIAPPTD_TRANS-WRBTR = LV_AMT_PAY_CURR.
            LS_FIAPPTD_TRANS-WAERS = IS_BSEG_PAY_GL-WAERS.
          ENDIF.
        ENDIF.
      ENDIF.

*******************************************************************
      IF LS_PAID_ITEMS-SHKZG EQ 'S'.
        LS_FIAPPTD_TRANS-VALUE_TYPE = 'E'.  "Entry
      ELSE.
        LS_FIAPPTD_TRANS-VALUE_TYPE = 'S'. "Exit
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

      CLEAR LS_TCURC.
      READ TABLE GT_TCURC INTO LS_TCURC WITH KEY
        WAERS = LS_XML_ITEMS-WAERS BINARY SEARCH.
      IF SY-SUBRC = 0 AND LS_TCURC-ISOCD IS NOT INITIAL.
        LS_XML_ITEMS-ISOCD = LS_TCURC-ISOCD.
        LS_FIAPPTD_TRANS-ISOCD = LS_TCURC-ISOCD.
      ENDIF.

      APPEND LS_XML_ITEMS TO CT_XML_ITEMS.
      APPEND LS_FIAPPTD_TRANS TO CT_FIAPPTD_TRANS.
      LV_PAY_DOC_PROCESS = 'X'.

    ENDLOOP.

    LS_FIAPPTD_TRANS = LS_TRANS_INIT.

    IF IS_BSEG_PAY_VC-TCODE NE 'F110'.
*  Only for manual payments
      CALL METHOD ME->PROCESS_PARTIAL_ITEMS
        EXPORTING
*         it_paid_items      =
          IS_BSEG_PAY_GL     = IS_BSEG_PAY_GL
          IS_BSEG_PAY_VC     = IS_BSEG_PAY_VC
          IT_LFA1            = IT_LFA1
          IT_KNA1            = IT_KNA1
          IS_FIAPPTD_TRANS   = LS_FIAPPTD_TRANS
          IT_BSEG_REM_ALL    = IT_BSEG_PARTIAL
        CHANGING
          CV_PAY_DOC_PROCESS = LV_PAY_DOC_PROCESS
          CT_FIAPPTD_TRANS   = CT_FIAPPTD_TRANS
          CT_XML_ITEMS       = CT_XML_ITEMS.
    ENDIF.

*    IF gt_oth_acc_range[] IS NOT INITIAL.

    LS_FIAPPTD_TRANS = LS_TRANS_INIT.

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

*    ENDIF.
  ENDMETHOD.