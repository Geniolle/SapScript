  METHOD PROCESS_OTHER_LINES.

    DATA: LV_LAND TYPE LAND1,
       LS_KNA1 TYPE KNA1,
       LS_LFA1 TYPE LFA1.

    DATA: LS_FIAPPTD_TRANS TYPE FIAPPTD_TRANS.
    DATA: LV_NPC2 TYPE CHAR20,
          LV_PAIS_ATIVO TYPE INTCA3,
          LV_ID_BANCO TYPE FIAPPT_BANKID,
          LV_REGISTER_EXISTS TYPE FLAG,
          LV_REGISTER_ID TYPE FIAPPT_REG_ID,
          LV_ORIGIN_RUN_ID TYPE FIAPPT_RUN_ID.

    DATA: OBJ_CBR_HANDLE TYPE REF TO BADI_RFFOPT_CBR.
    DATA: LS_XML_ITEMS TYPE FIAPPT_S_CBR_XML_ITEMS.

    DATA: LS_BSEG_ALL TYPE FIAPPT_S_TRANS_CBR,
          LV_FOREIGN_TRANS TYPE FLAG,
          LS_BSEG_PARTIAL TYPE FIAPPT_S_TRANS_CBR,
          LV_AMT_PAY_CURR TYPE WRBTR,
          PAY_DOC_PROCESS TYPE FLAG,
          LS_TCURC TYPE TCURC,
          LV_SCB_BANK TYPE FIAPPT_SCB_IND,
          LV_ISO_COUNTRY TYPE INTCA3.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID,
          OBSERVATIONS TYPE STRING.

    CONSTANTS: CV_CURR TYPE WAERS VALUE 'EUR'.
**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.

    LS_FIAPPTD_TRANS = IS_FIAPPTD_TRANS.

    CLEAR: LV_SCB_BANK, LV_ISO_COUNTRY.

    IF IV_PAY_DOC_PROCESS EQ 'X'.
*   For the other lines, get the SCB from the
*   G/L Bank line as the SCB codes cannot be maintined
*   in the other lines
      IF IS_BSEG_PAY_GL-LZBKZ IS NOT INITIAL.
        CALL METHOD ME->GET_SCB_IND
          EXPORTING
            IV_LZBKZ   = IS_BSEG_PAY_GL-LZBKZ
          IMPORTING
            EV_SCB_IND = LV_SCB_BANK.
      ENDIF.

      IF IS_BSEG_PAY_GL-LANDL IS NOT INITIAL.
        CALL METHOD READ_T005
          EXPORTING
            IV_LAND   = IS_BSEG_PAY_GL-LANDL
          IMPORTING
            EV_INTCA3 = LV_ISO_COUNTRY.
      ENDIF.

      LOOP AT IT_BSEG_REM_ALL INTO LS_BSEG_ALL WHERE
        BUKRS = GV_BUKRS AND BELNR = IS_BSEG_PAY_GL-BELNR
        AND KOART = 'S'. " AND hkont IN gt_oth_acc_range[].

        CLEAR LS_TCURC.
**Processing of the G/L lines(Cost lines, Tax lines etc.) of the
**payment document which has to be reported
* If the amount is zero then do not report the line
        IF LS_BSEG_ALL-WRBTR IS INITIAL.
          CONTINUE.
        ENDIF.

    IF LS_BSEG_ALL-HKONT IN GT_OTH_ACC_RANGE[] AND
        GT_OTH_ACC_RANGE IS NOT INITIAL.
     " OR ls_bseg_all-tcode EQ 'F110'.

        LS_FIAPPTD_TRANS-BELNR = LS_BSEG_ALL-BELNR.
        LS_FIAPPTD_TRANS-DOC_YEAR = LS_BSEG_ALL-GJAHR.
        LS_FIAPPTD_TRANS-BUZEI = LS_BSEG_ALL-BUZEI.
        LS_FIAPPTD_TRANS-REFERENCE_DATE = IS_BSEG_PAY_GL-BUDAT.
*        ls_fiapptd_trans-due_date = is_bseg_pay_gl-bldat.
        LS_FIAPPTD_TRANS-WRBTR = LS_BSEG_ALL-WRBTR - LS_BSEG_ALL-WSKTO - LS_BSEG_ALL-QBSHB .
        LS_FIAPPTD_TRANS-WAERS = LS_BSEG_ALL-WAERS.
*        ls_fiapptd_trans-due_date = is_bseg_pay_gl-bldat.
        LS_FIAPPTD_TRANS-ACCOUNT = LS_BSEG_ALL-HKONT.
        LS_FIAPPTD_TRANS-KOART = 'S'.
        LS_FIAPPTD_TRANS-STCEG = ''.

        IF LS_BSEG_ALL-SHKZG EQ 'S'.
          LS_FIAPPTD_TRANS-VALUE_TYPE = 'S'. "Exit
        ELSE.
          LS_FIAPPTD_TRANS-VALUE_TYPE = 'E'. "Entry
        ENDIF.

        IF LS_BSEG_ALL-LZBKZ IS NOT INITIAL.
          CALL METHOD ME->GET_SCB_IND
            EXPORTING
              IV_LZBKZ   = LS_BSEG_ALL-LZBKZ
            IMPORTING
              EV_SCB_IND = LS_FIAPPTD_TRANS-SCB_IND.
        ELSE.
          LS_FIAPPTD_TRANS-SCB_IND = LV_SCB_BANK.
        ENDIF.

        IF LS_BSEG_ALL-LANDL IS NOT INITIAL.
          CALL METHOD READ_T005
            EXPORTING
              IV_LAND   = LS_BSEG_ALL-LANDL
            IMPORTING
              EV_INTCA3 = LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE.
        ELSE.
          LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE = LV_ISO_COUNTRY.
        ENDIF.

*        if ls_bseg_all-tcode EQ 'F110' and
          IF LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE IS INITIAL.
            LS_FIAPPTD_TRANS-PAIS_CONTRAPARTE = IV_LAND_ISO.
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
*               paid_items    =
                FIAPPTD_TRANS = LS_FIAPPTD_TRANS
              CHANGING
                NPC2          = LV_NPC2
                PAIS_ATIVO    = LV_PAIS_ATIVO
                ID_BANCO      = LV_ID_BANCO.

          CATCH CX_SY_DYN_CALL_ILLEGAL_METHOD.
        ENDTRY.

        IF LV_ID_BANCO IS NOT INITIAL AND
          IS_BSEG_PAY_GL-BANK_ACC_TYPE EQ 'I'.
          LS_FIAPPTD_TRANS-BANK_ID = LV_ID_BANCO.
        ENDIF.
        LS_FIAPPTD_TRANS-NPC2 = LV_NPC2.
        LS_FIAPPTD_TRANS-PAIS_ATIVO = LV_PAIS_ATIVO.


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
            PAY_DOC_PROCESS = 'X'.
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
        PAY_DOC_PROCESS = 'X'.
       ENDIF.
      ENDLOOP.
    ENDIF.


  ENDMETHOD.