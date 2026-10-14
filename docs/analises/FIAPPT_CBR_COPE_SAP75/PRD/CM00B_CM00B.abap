  METHOD GET_BANK_BALANCES_2.

    DATA: LT_GLT0      TYPE FAGL_T_GLT0,
          LS_GLT0      LIKE LINE OF LT_GLT0,
          LT_ACC_RANGE TYPE FIAPPT_T_ACC_RANGE,
          LS_ACC_RANGE LIKE LINE OF LT_ACC_RANGE.
    DATA: LS_XML_ITEMS TYPE FIAPPT_S_CBR_XML_ITEMS.
    DATA: LS_FIAPPTD_BALANCE TYPE FIAPPTD_BALANCE,
          LS_FIAP_BAL        TYPE FIAPPTD_BALANCE,
          LS_SKB1            TYPE SKB1,
          LT_SKB1            TYPE TABLE OF SKB1,
          LT_SKB1_CHANGE     TYPE TABLE OF SKB1,
          LS_SKB1_CHANGE     TYPE SKB1.
    DATA: LV_GL_ACC TYPE HKONT,
          LV_CURR   TYPE WAERS.
    DATA: LV_BANKL TYPE BANKL,
          LV_BANKS TYPE BANKS.
    DATA: LV_NPC2     TYPE FIAPPT_NPC2,
          LV_ID_BANCO TYPE FIAPPT_BANKID.
    DATA: LV_HBKID TYPE HBKID.
    DATA: OBJ_CBR_HANDLE   TYPE REF TO BADI_RFFOPT_CBR,
          LV_GLFLEX_ACTIVE TYPE BOOLE_D,
          LT_FAGLFLEXT     TYPE TABLE OF FAGLFLEXT,
          LS_FAGLFLEXT     TYPE FAGLFLEXT,
          LT_GROUPLIST     TYPE TTFIELDNAME,
          LV_ACCOUNT       TYPE FIAPPT_ACCOUNT_NO,
          LV_AMT_LOC_CURR  TYPE WRBTR.

    TYPES : BEGIN OF GL_BALANCE,
              GL_ACCOUNT TYPE HKONT,
              HBKID      TYPE HBKID,
              WAERS      TYPE WAERS,
              WRBTR      TYPE FIAPPT_AMOUNT,
            END OF GL_BALANCE.

    DATA: LS_GL_BALANCE TYPE GL_BALANCE,
          LT_GL_BALANCE TYPE TABLE OF GL_BALANCE.

    TYPES: BEGIN OF BANK_DETAILS,
             ZBUKR TYPE BUKRS,
             HBKID TYPE HBKID,
             UKONT TYPE HKONT,
             WAERS TYPE WAERS,
           END OF BANK_DETAILS.

    DATA: LS_BANK_DETAILS   TYPE BANK_DETAILS,
          LT_BANK_DETAILS   TYPE TABLE OF BANK_DETAILS,
          LT_BANK_DETAILS_2 TYPE TABLE OF BANK_DETAILS.

    DATA: LV_REG_EXISTS   TYPE FLAG,
          LV_REG_CHANGED  TYPE FLAG,
          LV_ORIGIN_RUNID TYPE FIAPPT_RUN_ID.

    DATA: LT_BAL_SCB   TYPE TABLE OF FIAPPTD_BAL_SCB,
          LS_BAL_SCB   TYPE FIAPPTD_BAL_SCB,
          LS_BAL_SCB_H TYPE FIAPPTD_BAL_SCB,
          LS_BAL_SCB_S TYPE FIAPPTD_BAL_SCB,
          LS_TCURC     TYPE TCURC.

    DATA: LV_LEN  TYPE I,
          LV_LEN1 TYPE I.

    DATA: LV_RUNID_SEQ TYPE FIAPPT_RUN_ID.

    CONSTANTS CV_CURR TYPE WAERS VALUE 'EUR'.

**Instantiation of Badi..
    GET BADI OBJ_CBR_HANDLE.

    CALL METHOD INIT_FIAPPTD_BAL_STR
      EXPORTING
        IT_DATE_RANGE      = IT_DATE_RANGE
      CHANGING
        ES_FIAPPTD_BALANCE = LS_FIAPPTD_BALANCE.
    LS_FIAPPTD_BALANCE-KOART = 'S'.
*    ls_fiapptd_balance-bank_acc_type = 'E'.

    LT_ACC_RANGE[] = GT_EXT_ACC_RANGE[].
*    APPEND LINES OF gt_int_acc_range TO lt_acc_range.
*    G/l Balances
    APPEND LINES OF GT_BAL_ACC_RANGE TO LT_ACC_RANGE.
    APPEND LINES OF GT_OTH_ACC_RANGE TO LT_ACC_RANGE.

    IF LT_ACC_RANGE[] IS NOT INITIAL.

      CALL METHOD ME->GET_GL_ACC_FROM_FSV
        EXPORTING
          IT_ACC_RANGE = LT_ACC_RANGE
        IMPORTING
          ET_SKB1      = LT_SKB1.


      LT_SKB1_CHANGE[] = LT_SKB1[].

      CALL FUNCTION 'FAGL_CHECK_GLFLEX_ACTIVE'
        EXPORTING
          ID_BUKRS        = GV_BUKRS
        IMPORTING
          E_GLFLEX_ACTIVE = LV_GLFLEX_ACTIVE
        EXCEPTIONS
          ERROR_IN_SETUP  = 1
          OTHERS          = 2.

      IF SY-SUBRC <> 0.
* Implement suitable error handling here
      ENDIF.

      IF LV_GLFLEX_ACTIVE IS INITIAL.

        CALL FUNCTION 'FAGL_GET_GLT0'
          EXPORTING
            I_GLT0_RLDNR      = '00'
            I_RRCTY           = '0'
            I_RVERS           = '001'
            I_BUKRS           = GV_BUKRS
            I_RYEAR           = GV_YEAR
            I_RPMAX           = '016'
            I_RANGE_RACCT     = LT_ACC_RANGE
          IMPORTING
            ET_GLT0           = LT_GLT0
          EXCEPTIONS
            INVALID_SELECTION = 1
            OTHERS            = 2.
        IF SY-SUBRC <> 0.
* Implement suitable error handling here
*    ELSE.
        ENDIF.

      ELSE.

        CALL METHOD READ_FAGLFLEXT
          EXPORTING
            IV_LEDGER    = IV_LEDGER
            IT_ACC_RANGE = LT_ACC_RANGE
          IMPORTING
            ET_GLT0      = LT_GLT0.

      ENDIF.


      SORT LT_GLT0 BY RACCT RTCUR.

      CLEAR: LV_GL_ACC, LV_CURR, LV_HBKID.

      LOOP AT LT_GLT0 INTO LS_GLT0.

        IF LV_GL_ACC NE LS_GLT0-RACCT OR LV_CURR NE LS_GLT0-RTCUR.
          IF LV_GL_ACC IS INITIAL AND LV_CURR IS INITIAL.

            CLEAR LV_HBKID.
            SELECT SINGLE HBKID INTO LV_HBKID FROM T042I WHERE ZBUKR = LS_GLT0-BUKRS AND UKONT = LS_GLT0-RACCT.
            IF SY-SUBRC NE 0.
*            SELECT SINGLE hbkid INTO lv_hbkid FROM t012k WHERE bukrs = ls_glt0-bukrs AND hkont = ls_glt0-racct.
*          ENDIF.
*            IF sy-subrc EQ 0.

              DATA(LT_T012K) = CL_FCLM_HOUSE_BANK_ACCOUNT=>GET_T012K_T( EXPORTING
                                                                          IT_BUKRS = VALUE #( ( LS_GLT0-BUKRS ) )
                                                                          IT_HKONT = VALUE #( ( LS_GLT0-RACCT ) ) ).

              READ TABLE LT_T012K ASSIGNING FIELD-SYMBOL(<FS_T012K>) INDEX 1.
              IF SY-SUBRC EQ 0.
                LV_HBKID = <FS_T012K>-HBKID.
                LS_BANK_DETAILS-ZBUKR = LS_GLT0-BUKRS.
                LS_BANK_DETAILS-HBKID = LV_HBKID.
                LS_BANK_DETAILS-UKONT = LS_GLT0-RACCT.
                LS_BANK_DETAILS-WAERS = LS_GLT0-RTCUR.
                DELETE LT_SKB1_CHANGE WHERE BUKRS = LS_GLT0-BUKRS
                  AND SAKNR = LS_GLT0-RACCT.
                APPEND LS_BANK_DETAILS TO LT_BANK_DETAILS.
              ENDIF.
            ENDIF.

          ELSE.

            IF LV_GL_ACC NE LS_GLT0-RACCT.

              CLEAR LV_HBKID.
              SELECT SINGLE HBKID INTO LV_HBKID FROM T042I WHERE ZBUKR = LS_GLT0-BUKRS AND UKONT = LS_GLT0-RACCT.
              IF SY-SUBRC NE 0.
*                select single hbkid into lv_hbkid from t012k where bukrs = ls_glt0-bukrs and hkont = ls_glt0-racct.
*              endif.
*              if sy-subrc eq 0.

                LT_T012K = CL_FCLM_HOUSE_BANK_ACCOUNT=>GET_T012K_T( EXPORTING
                                                                      IT_BUKRS = VALUE #( ( LS_GLT0-BUKRS ) )
                                                                      IT_HKONT = VALUE #( ( LS_GLT0-RACCT ) ) ).

                READ TABLE LT_T012K ASSIGNING <FS_T012K> INDEX 1.
                IF SY-SUBRC EQ 0.
                  LV_HBKID = <FS_T012K>-HBKID.
                  LS_BANK_DETAILS-ZBUKR = LS_GLT0-BUKRS.
                  LS_BANK_DETAILS-HBKID = LV_HBKID.
                  LS_BANK_DETAILS-UKONT = LS_GLT0-RACCT.
                  LS_BANK_DETAILS-WAERS = LS_GLT0-RTCUR.
                  DELETE LT_SKB1_CHANGE WHERE BUKRS = LS_GLT0-BUKRS
                   AND SAKNR = LS_GLT0-RACCT.
                  APPEND LS_BANK_DETAILS TO LT_BANK_DETAILS.
                ENDIF.
              ENDIF.

            ELSEIF LV_CURR NE LS_GLT0-RTCUR.
              LS_BANK_DETAILS-WAERS = LS_GLT0-RTCUR.
              APPEND LS_BANK_DETAILS TO LT_BANK_DETAILS.
            ENDIF.

            APPEND LS_GL_BALANCE TO LT_GL_BALANCE.
            CLEAR LS_GL_BALANCE.
          ENDIF.

          LV_GL_ACC = LS_GLT0-RACCT.
          LV_CURR = LS_GLT0-RTCUR.
        ENDIF.

        IF LV_HBKID IS NOT INITIAL.
          LS_GL_BALANCE-GL_ACCOUNT = LS_GLT0-RACCT.
          LS_GL_BALANCE-HBKID = LV_HBKID.
          LS_GL_BALANCE-WAERS = LS_GLT0-RTCUR.

          CASE GV_PERIOD.
            WHEN '01'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01.
            WHEN '02'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02.
            WHEN '03'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03.
            WHEN '04'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04.
            WHEN '05'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05.
            WHEN '06'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06.
            WHEN '07'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07.
            WHEN '08'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07 + LS_GLT0-TSL08.
            WHEN '09'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07 + LS_GLT0-TSL08 +
                                    LS_GLT0-TSL09.
            WHEN '10'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07 + LS_GLT0-TSL08 +
                                    LS_GLT0-TSL09 + LS_GLT0-TSL10.
            WHEN '11'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07 + LS_GLT0-TSL08 +
                                    LS_GLT0-TSL09 + LS_GLT0-TSL10 +
                                    LS_GLT0-TSL11.
            WHEN '12'.
              LS_GL_BALANCE-WRBTR = LS_GL_BALANCE-WRBTR + LS_GLT0-TSLVT +
                                    LS_GLT0-TSL01 + LS_GLT0-TSL02 +
                                    LS_GLT0-TSL03 + LS_GLT0-TSL04 +
                                    LS_GLT0-TSL05 + LS_GLT0-TSL06 +
                                    LS_GLT0-TSL07 + LS_GLT0-TSL08 +
                                    LS_GLT0-TSL09 + LS_GLT0-TSL10 +
                                    LS_GLT0-TSL11 + LS_GLT0-TSL12 +
                                    LS_GLT0-TSL13 + LS_GLT0-TSL14 +
                                    LS_GLT0-TSL15 + LS_GLT0-TSL16.
          ENDCASE.

        ENDIF.
      ENDLOOP.

      IF LS_GL_BALANCE-GL_ACCOUNT IS NOT INITIAL.
        APPEND LS_GL_BALANCE TO LT_GL_BALANCE.
      ENDIF.


      IF LT_SKB1_CHANGE IS NOT INITIAL.

        SELECT * FROM T042I INTO CORRESPONDING FIELDS OF TABLE LT_BANK_DETAILS_2
        FOR ALL ENTRIES IN LT_SKB1_CHANGE WHERE ZBUKR EQ LT_SKB1_CHANGE-BUKRS AND UKONT EQ LT_SKB1_CHANGE-SAKNR
                                                 ORDER BY PRIMARY KEY.

        LOOP AT LT_SKB1_CHANGE INTO LS_SKB1_CHANGE.
          CLEAR LS_BANK_DETAILS.
          IF LS_SKB1_CHANGE-HBKID IS NOT INITIAL.
            LS_BANK_DETAILS-ZBUKR = LS_SKB1_CHANGE-BUKRS.
            LS_BANK_DETAILS-UKONT = LS_SKB1_CHANGE-SAKNR.
            LS_BANK_DETAILS-HBKID = LS_SKB1_CHANGE-HBKID.
            IF LS_SKB1_CHANGE-WAERS IS NOT INITIAL.
              LS_BANK_DETAILS-WAERS = LS_SKB1_CHANGE-WAERS.
            ELSE.
              LS_BANK_DETAILS-WAERS = CV_CURR.
            ENDIF.
            APPEND LS_BANK_DETAILS TO LT_BANK_DETAILS_2.
          ENDIF.
        ENDLOOP.

      ENDIF.
      APPEND LINES OF LT_BANK_DETAILS_2 TO LT_BANK_DETAILS.
      CLEAR: LV_HBKID, LV_CURR.

      SORT LT_BANK_DETAILS BY HBKID WAERS.
      DELETE ADJACENT DUPLICATES FROM LT_BANK_DETAILS COMPARING HBKID WAERS.
      LS_FIAP_BAL = LS_FIAPPTD_BALANCE.


* Fetch all the SCB codes for the House bank ID
* from the table fiapptd_bal_scb
      SELECT * FROM FIAPPTD_BAL_SCB INTO TABLE LT_BAL_SCB
        WHERE BUKRS = GV_BUKRS AND KOART = 'S'.
*Reading the default SCB values
      READ TABLE LT_BAL_SCB INTO LS_BAL_SCB_H WITH KEY
                  BUKRS = GV_BUKRS
                  ACCOUNT = '*'
                  KOART = 'S'
                  SHKZG = 'H'.

      READ TABLE LT_BAL_SCB INTO LS_BAL_SCB_S WITH KEY
                  BUKRS = GV_BUKRS
                  ACCOUNT = '*'
                  KOART = 'S'
                  SHKZG = 'S'.

      DELETE LT_GL_BALANCE WHERE WRBTR IS INITIAL.

      LOOP AT LT_GL_BALANCE INTO LS_GL_BALANCE WHERE
        ( WAERS = 'PTE' OR WAERS = 'SDR'
        OR WAERS EQ 'ESP' OR WAERS EQ 'ESP.' OR WAERS EQ 'FRF' ).

        CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
          EXPORTING
            DATE             = LS_FIAPPTD_BALANCE-REFERENCE_DATE
            FOREIGN_AMOUNT   = LS_GL_BALANCE-WRBTR
            FOREIGN_CURRENCY = LS_GL_BALANCE-WAERS
            LOCAL_CURRENCY   = CV_CURR
          IMPORTING
            LOCAL_AMOUNT     = LV_AMT_LOC_CURR.

        DELETE LT_GL_BALANCE WHERE GL_ACCOUNT = LS_GL_BALANCE-GL_ACCOUNT AND
        HBKID = LS_GL_BALANCE-HBKID AND WAERS = LS_GL_BALANCE-WAERS AND
        WRBTR = LS_GL_BALANCE-WRBTR.

        IF LV_AMT_LOC_CURR IS NOT INITIAL.
          LS_GL_BALANCE-WAERS = CV_CURR.
          LS_GL_BALANCE-WRBTR = LV_AMT_LOC_CURR.
          APPEND LS_GL_BALANCE TO LT_GL_BALANCE.
        ENDIF.
      ENDLOOP.

      LOOP AT LT_BANK_DETAILS INTO LS_BANK_DETAILS.

        CLEAR: LV_REG_EXISTS,  LV_REG_CHANGED , LV_ORIGIN_RUNID,
               LS_TCURC.
        LS_FIAPPTD_BALANCE = LS_FIAP_BAL.
        LS_FIAPPTD_BALANCE-ACCOUNT = LS_BANK_DETAILS-HBKID.
        LS_FIAPPTD_BALANCE-WAERS = LS_BANK_DETAILS-WAERS.
        CLEAR LS_GL_BALANCE.

* determination of bank_acc_type: EXT -> E, BAL -> C and OTH -> O
        CLEAR LS_FIAPPTD_BALANCE-BANK_ACC_TYPE.
        LOOP AT GT_EXT_ACC_RANGE INTO LS_ACC_RANGE WHERE LOW <= LS_BANK_DETAILS-UKONT
                                                   AND HIGH >= LS_BANK_DETAILS-UKONT.
        ENDLOOP.
        IF SY-SUBRC = 0.
          LS_FIAPPTD_BALANCE-BANK_ACC_TYPE = 'E'.
        ELSE.
          LOOP AT GT_BAL_ACC_RANGE INTO LS_ACC_RANGE WHERE LOW <= LS_BANK_DETAILS-UKONT
                                                     AND HIGH >= LS_BANK_DETAILS-UKONT.
          ENDLOOP.
          IF SY-SUBRC = 0.
            LS_FIAPPTD_BALANCE-BANK_ACC_TYPE = 'C'.
          ELSE.
            LOOP AT GT_OTH_ACC_RANGE INTO LS_ACC_RANGE WHERE LOW <= LS_BANK_DETAILS-UKONT
                                                       AND HIGH >= LS_BANK_DETAILS-UKONT.
            ENDLOOP.
            IF SY-SUBRC = 0.
              LS_FIAPPTD_BALANCE-BANK_ACC_TYPE = 'O'.
            ENDIF.
          ENDIF.
        ENDIF.

        LOOP AT LT_GL_BALANCE INTO LS_GL_BALANCE WHERE
          HBKID = LS_BANK_DETAILS-HBKID AND
          WAERS = LS_BANK_DETAILS-WAERS.
          LS_FIAPPTD_BALANCE-WRBTR = LS_FIAPPTD_BALANCE-WRBTR +
                                     LS_GL_BALANCE-WRBTR.
        ENDLOOP.
* Zero balance? Continue the processing.
        IF LS_FIAPPTD_BALANCE-WRBTR IS INITIAL.
          CONTINUE.
        ENDIF.

* Fill the SCB Codes for the House Bank ID based on
* the sign of the amount
        IF LS_FIAPPTD_BALANCE-ACCOUNT+0(1) NA '0123456789'.
          LV_ACCOUNT = LS_FIAPPTD_BALANCE-ACCOUNT.
        ELSE.
          CLEAR: LV_LEN, LV_LEN1.
          LV_ACCOUNT = '0000000000'.
          LV_LEN = STRLEN( LS_FIAPPTD_BALANCE-ACCOUNT ).
          LV_LEN1 = 10 - LV_LEN.
          LV_ACCOUNT+LV_LEN1(LV_LEN) = LS_FIAPPTD_BALANCE-ACCOUNT.
        ENDIF.

        IF LS_FIAPPTD_BALANCE-WRBTR < 0.
          LS_FIAPPTD_BALANCE-WRBTR = LS_FIAPPTD_BALANCE-WRBTR * -1.
          READ TABLE LT_BAL_SCB INTO LS_BAL_SCB WITH KEY
            BUKRS = GV_BUKRS
            ACCOUNT =  LV_ACCOUNT
            KOART = 'S'
            SHKZG = 'H'.
          IF SY-SUBRC = 0.
            LS_FIAPPTD_BALANCE-SCB_IND = LS_BAL_SCB-SCB_IND.
          ELSE.
            LS_FIAPPTD_BALANCE-SCB_IND = LS_BAL_SCB_H-SCB_IND.
          ENDIF.
        ELSE.
          READ TABLE LT_BAL_SCB INTO LS_BAL_SCB WITH KEY
           BUKRS = GV_BUKRS
           ACCOUNT = LV_ACCOUNT
           KOART = 'S'
           SHKZG = 'S'.
          IF SY-SUBRC = 0.
            LS_FIAPPTD_BALANCE-SCB_IND = LS_BAL_SCB-SCB_IND.
          ELSE.
            LS_FIAPPTD_BALANCE-SCB_IND = LS_BAL_SCB_S-SCB_IND.
          ENDIF.
        ENDIF.

        CALL METHOD GET_BANK_DATA
          EXPORTING
            IV_HBKID = LS_BANK_DETAILS-HBKID
          IMPORTING
            EV_BANKL = LV_BANKL
            EV_BANKS = LV_BANKS.

        IF LS_BANK_DETAILS-UKONT IN GT_INT_ACC_RANGE[].
          LS_FIAPPTD_BALANCE-BANK_ID = LV_BANKL.
        ENDIF.
        LS_FIAPPTD_BALANCE-PAIS_CONTA = LV_BANKS.
        LS_FIAPPTD_BALANCE-PAIS_CONTRAPARTE = LV_BANKS.

        TRY.
            CALL BADI OBJ_CBR_HANDLE->GET_BALANCES
              EXPORTING
                FIAPPTD_BALANCE = LS_FIAPPTD_BALANCE
              CHANGING
                NPC2            = LV_NPC2
                ID_BANCO        = LV_ID_BANCO.
          CATCH CX_SY_DYN_CALL_ILLEGAL_METHOD.
        ENDTRY.
        IF LV_ID_BANCO IS NOT INITIAL.
          LS_FIAPPTD_BALANCE-BANK_ID = LV_ID_BANCO.
        ENDIF.
        LS_FIAPPTD_BALANCE-NPC2 = LV_NPC2.

        IF GV_PERIOD_1ST_RUN EQ 'X'.
*   Add the next sequence to the run id
          GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
          LV_RUNID_SEQ = GV_RUNID_SEQ.
          SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
          CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_BALANCE-ORIGIN_RUN_ID.

        ELSE.
          CALL METHOD ME->CHECK_OLD_BAL_REGISTERS
            EXPORTING
              IS_FIAPPTD_BALANCE = LS_FIAPPTD_BALANCE
            IMPORTING
              EV_REG_EXISTS      = LV_REG_EXISTS
              EV_REG_CHANGED     = LV_REG_CHANGED
              EV_ORIGIN_RUNID    = LV_ORIGIN_RUNID.

          IF LV_REG_EXISTS EQ 'X' AND LV_REG_CHANGED = 'X'.
            LS_FIAPPTD_BALANCE-REGISTER_ID = 'M'.
            LS_FIAPPTD_BALANCE-ORIGIN_RUN_ID = LV_ORIGIN_RUNID.
          ELSEIF LV_REG_EXISTS EQ ''.
            LS_FIAPPTD_BALANCE-REGISTER_ID = 'C'.
*   Add the next sequence to the run id
            GV_RUNID_SEQ = GV_RUNID_SEQ + 1.
            LV_RUNID_SEQ = GV_RUNID_SEQ.
            SHIFT LV_RUNID_SEQ LEFT DELETING LEADING SPACE.
            CONCATENATE GV_RUNID LV_RUNID_SEQ INTO LS_FIAPPTD_BALANCE-ORIGIN_RUN_ID.

          ELSEIF LV_REG_EXISTS EQ 'X'.
            CONTINUE.
          ENDIF.
        ENDIF.

        MOVE-CORRESPONDING LS_FIAPPTD_BALANCE TO LS_XML_ITEMS.

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
          LS_FIAPPTD_BALANCE-ISOCD = LS_TCURC-ISOCD.
        ENDIF.

        APPEND LS_XML_ITEMS TO ET_XML_ITEMS.
        APPEND LS_FIAPPTD_BALANCE TO ET_FIAPPTD_BALANCE.
      ENDLOOP.

    ENDIF.

    ET_GLT0 = LT_GLT0.

  ENDMETHOD.