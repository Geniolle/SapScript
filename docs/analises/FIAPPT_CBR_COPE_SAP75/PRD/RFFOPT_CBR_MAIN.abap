*&---------------------------------------------------------------------*
*&  Include           RFFOPT_CBR_MAIN
*&---------------------------------------------------------------------*
  INITIALIZATION.
    B0SG-XLDDS = 'X'.
    B0SG-XLDKS = 'X'.

* selection screen and behavior differs in cloud
*  - file cannot be saved on application server
*  - Provide a button to access the view FIAPPTV_BAL_SCB
    GV_IS_S4H_CLOUD  =  CL_COS_UTILITIES=>IS_S4H_CLOUD( ).

    MNT_SCB = 'Assign SCB Indicator'(019).

  START-OF-SELECTION.
*
**Creating the instance for the user exit.
    PERFORM GET_BADI_INSTANCE.

    CREATE OBJECT OBJ_CBR
      EXPORTING
        IV_BUKRS  = P_BUKRS
        IV_YEAR   = P_YEAR
        IV_PERIOD = P_PERIOD
        IV_FSV    = P_FSV
        IV_RUNID  = P_RUN_ID.

    CALL METHOD OBJ_CBR->READ_ACC_FROM_FSV
      EXCEPTIONS
        NO_FSV_DATA    = 1
        NO_INT_EXT_ACC = 2
        OTHERS         = 3.
    IF SY-SUBRC = 1.
      MESSAGE E009(FIAPPT_CBR).
    ELSEIF SY-SUBRC = 2.
      MESSAGE E013(FIAPPT_CBR) WITH P_FSV.
    ENDIF.


  END-OF-SELECTION.

*********************FM LDB***************************

    DATA: LT_BUDAT_RANGE TYPE FIAPPT_T_DATE_RANGE.

    FREE GT_FIAPPTD_TRANS.                          "note 2215769

    CALL METHOD OBJ_CBR->INITIALIZE_XML_HEADER
      EXPORTING
        IV_TEST       = P_TEST
      IMPORTING
        ES_XML_HEADER = GS_XML_HEADER.

    IF GS_XML_HEADER-NPC IS INITIAL.
      IF P_TEST = 'X'.
        MESSAGE W014(FIAPPT_CBR) WITH P_BUKRS.
      ELSE.
        MESSAGE E014(FIAPPT_CBR) WITH P_BUKRS.
      ENDIF.
    ENDIF.

    CALL METHOD OBJ_CBR->SET_POSTING_DATE
      IMPORTING
        ET_POSTING_DATE = LT_BUDAT_RANGE[].

    PERFORM CALL_LDB_KDF.
    PERFORM CALL_LDB_DDF.

    PERFORM CHECK_VAT_REG_NUMBER_COUNTRIES.

    CALL METHOD OBJ_CBR->GET_ISO_CURR_CODES.
    CALL METHOD OBJ_CBR->GET_ISO_COUNTRY_CODES.

    PERFORM GET_PAYMENT_DOCS.

    CALL METHOD OBJ_CBR->PROCESS_ADDITIONAL_PAYMENTS
      EXPORTING
        IT_LFA1          = GT_LFA1
        IT_KNA1          = GT_KNA1
      CHANGING
        CT_FIAPPTD_TRANS = GT_FIAPPTD_TRANS
        CT_XML_ITEMS     = GT_XML_ITEMS_TRANS.


    PERFORM CLEAR_GLOBAL_VARIABLES.

    CALL METHOD OBJ_CBR->GET_VEND_CUS_BALANCES
      EXPORTING
        IT_LFA1            = GT_LFA1
        IT_KNA1            = GT_KNA1
        IT_DATE_RANGE      = LT_BUDAT_RANGE
        IT_VEND_SPLGL      = P_GL_V[]
        IT_CUST_SPLGL      = P_GL_C[]
        IT_VEND            = P_VEND[]
        IT_CUST            = P_CUST[]
      CHANGING
        ET_FIAPPTD_BALANCE = GT_FIAPPTD_BALANCE
        ET_XML_ITEMS       = GT_XML_ITEMS_BAL.

    IF P_GL IS INITIAL.
      P_GL = '0L'.
    ENDIF.
* House bank balances
    CALL METHOD OBJ_CBR->GET_BANK_BALANCES_2
      EXPORTING
        IT_DATE_RANGE      = LT_BUDAT_RANGE
        IV_LEDGER          = P_GL
      IMPORTING
        ET_GLT0            = GT_GLT0
      CHANGING
        ET_FIAPPTD_BALANCE = GT_FIAPPTD_BALANCE
        ET_XML_ITEMS       = GT_XML_ITEMS_BAL.

* G/L Balances
     CALL METHOD OBJ_CBR->GET_GL_BAL
      EXPORTING
        IT_DATE_RANGE      = LT_BUDAT_RANGE
        IT_GLT0            = GT_GLT0
      CHANGING
        ET_FIAPPTD_BALANCE = GT_FIAPPTD_BALANCE
        ET_XML_ITEMS       = GT_XML_ITEMS_BAL.


    GT_XML_ITEMS_ALL[] = GT_XML_ITEMS_TRANS[].
    APPEND LINES OF GT_XML_ITEMS_BAL TO GT_XML_ITEMS_ALL.

    CALL METHOD OBJ_CBR->CREATE_XML_FILE
      EXPORTING
        IS_HEADER      = GS_XML_HEADER
        IT_ITEMS       = GT_XML_ITEMS_ALL
      IMPORTING
        EV_XML_XSTRING = GV_XML_XSTRING
      EXCEPTIONS
        NO_XML         = 1
        OTHERS         = 2.
    IF SY-SUBRC <> 0.
      MESSAGE E015.
    ENDIF.


*******CALL METHOD TO DWNLOAD*******************
    IF ( P_PC EQ 'X' AND P_FILE1 IS NOT INITIAL )
      AND SY-BATCH IS INITIAL.
      GV_FILENAME = P_FILE1.
      CALL METHOD OBJ_CBR->DOWNLOAD_FILE
        EXPORTING
          IV_DWNLD_TO_PC = 'X'
          IV_FILEPATH    = GV_FILENAME
          IV_XML_XSTRING = GV_XML_XSTRING
        EXCEPTIONS
          FILEPATH_ERROR = 1
          OTHERS         = 2.
      IF SY-SUBRC <> 0.
        MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.

    ELSEIF P_APPL EQ 'X' AND P_FILE2 IS NOT INITIAL.
      GV_FILENAME = P_FILE2.
      CALL METHOD OBJ_CBR->DOWNLOAD_FILE
        EXPORTING
          IV_DWNLD_TO_PC = ' '
          IV_FILEPATH    = GV_FILENAME
          IV_XML_XSTRING = GV_XML_XSTRING
        EXCEPTIONS
          FILEPATH_ERROR = 1
          OTHERS         = 2.
      IF SY-SUBRC <> 0.
        MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.

    ENDIF.

    IF P_TEST NE 'X'.

      CALL METHOD OBJ_CBR->UPDATE_DB_TABLES
        EXPORTING
          IT_FIAPPTD_TRANS   = GT_FIAPPTD_TRANS
          IT_FIAPPTD_BALANCE = GT_FIAPPTD_BALANCE.

    ENDIF.
*********ALV PROCESSING********************************************************

    PERFORM BUILD_FIELD_CATALOG.

    PERFORM BUILD_LAYOUT.

    PERFORM EVENTTAB_BUILD.

    PERFORM ALV_DISPLAY.