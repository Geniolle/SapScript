*----------------------------------------------------------------------*
***INCLUDE RFFOPT_CBR_F01.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Form  BUILD_FIELD_CATALOG
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM BUILD_FIELD_CATALOG .
* Build Fieldcatalog for displaying the Foreign Transactions
  CALL FUNCTION 'REUSE_ALV_FIELDCATALOG_MERGE'
    EXPORTING
*     I_PROGRAM_NAME         =
*     I_INTERNAL_TABNAME     =
      I_STRUCTURE_NAME       = 'FIAPPTD_TRANS'
*     I_CLIENT_NEVER_DISPLAY = 'X'
*     I_INCLNAME             =
*     I_BYPASSING_BUFFER     =
*     I_BUFFER_ACTIVE        =
    CHANGING
      CT_FIELDCAT            = GT_FIELDCATALOG_TRANS
    EXCEPTIONS
      INCONSISTENT_INTERFACE = 1
      PROGRAM_ERROR          = 2
      OTHERS                 = 3.

  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
             WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

* Build Fieldcatalog for displaying the Vendor/Customer and Bank Balances
  CALL FUNCTION 'REUSE_ALV_FIELDCATALOG_MERGE'
    EXPORTING
*     I_PROGRAM_NAME         =
*     I_INTERNAL_TABNAME     =
      I_STRUCTURE_NAME       = 'FIAPPTD_BALANCE'
*     I_CLIENT_NEVER_DISPLAY = 'X'
*     I_INCLNAME             =
*     I_BYPASSING_BUFFER     =
*     I_BUFFER_ACTIVE        =
    CHANGING
      CT_FIELDCAT            = GT_FIELDCATALOG_BALANCES
    EXCEPTIONS
      INCONSISTENT_INTERFACE = 1
      PROGRAM_ERROR          = 2
      OTHERS                 = 3.
  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
                 WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

  PERFORM MODIFY_FIELDCAT.

ENDFORM.                    " BUILD_FIELD_CATALOG
*&---------------------------------------------------------------------*
*&      Form  SET_PF_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->EX_TAB     text
*----------------------------------------------------------------------*
FORM SET_PF_STATUS USING EX_TAB TYPE SLIS_T_EXTAB.

  SET PF-STATUS 'CBR_TRANS_STATUS'.

ENDFORM.                    "SET_PF_STATUS
*&---------------------------------------------------------------------*
*&      Form  SET_BAL_PF_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->EX_TAB     text
*----------------------------------------------------------------------*
FORM SET_BAL_PF_STATUS USING EX_TAB TYPE SLIS_T_EXTAB.

  SET PF-STATUS 'CBR_BAL_STATUS'.

ENDFORM.                    "SET_PF_STATUS
*&---------------------------------------------------------------------*
*&      Form  BUILD_LAYOUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM BUILD_LAYOUT .
  DATA: LV_X TYPE C VALUE 'X'.
* To expand the header table for item details
  GS_LAYOUT-EXPAND_FIELDNAME = 'EXPAND'.
  GS_LAYOUT-COLWIDTH_OPTIMIZE = LV_X.
  GS_LAYOUT-LIST_APPEND = 'X'.
  GS_LAYOUT-ZEBRA = LV_X.

ENDFORM.                    " BUILD_LAYOUT
*&---------------------------------------------------------------------*
*&      Form  EVENTTAB_BUILD
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM EVENTTAB_BUILD .

* Registration of events to happen during list display
  DATA: LS_EVENT TYPE SLIS_ALV_EVENT.

  CALL FUNCTION 'REUSE_ALV_EVENTS_GET'
    EXPORTING
      I_LIST_TYPE     = 0
    IMPORTING
      ET_EVENTS       = GT_EVENTS
    EXCEPTIONS
      LIST_TYPE_WRONG = 1
      OTHERS          = 2.

  IF SY-SUBRC <> 0.
    MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.

* Assgining Form Name as USER_COMMAND for Transactions
  READ TABLE GT_EVENTS INTO LS_EVENT
                      WITH KEY NAME = SLIS_EV_USER_COMMAND.
  IF SY-SUBRC = 0.

    LS_EVENT-FORM =  'USER_COMMAND'.
    MODIFY GT_EVENTS FROM LS_EVENT INDEX SY-TABIX.
* Assgining Form Name as USER_COMMAND_BAL for Balances
    LS_EVENT-FORM = 'USER_COMMAND_BAL'.
    GT_EVENTS_BAL[] = GT_EVENTS[].
    MODIFY GT_EVENTS_BAL FROM LS_EVENT INDEX SY-TABIX.

  ENDIF.

ENDFORM.                    " EVENTTAB_BUILD
*&---------------------------------------------------------------------*
*&      Form  ALV_DISPLAY
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM ALV_DISPLAY.

  DATA: LV_BAL_TITLE      TYPE LVC_TITLE,
        LV_TRANS_TITLE    TYPE LVC_TITLE,
        LV_BAL_COUNT(8)   TYPE C,
        LV_TRANS_COUNT(8) TYPE C.

  DESCRIBE TABLE GT_FIAPPTD_TRANS LINES LV_TRANS_COUNT.
  DESCRIBE TABLE GT_FIAPPTD_BALANCE LINES LV_BAL_COUNT.
  SHIFT LV_TRANS_COUNT LEFT DELETING LEADING SPACE.
  SHIFT LV_BAL_COUNT LEFT DELETING LEADING SPACE.

  LV_TRANS_TITLE = TEXT-003.
  LV_BAL_TITLE = TEXT-004.
  REPLACE '&1' WITH LV_TRANS_COUNT INTO LV_TRANS_TITLE.
  REPLACE '&1' WITH LV_BAL_COUNT INTO LV_BAL_TITLE.

  IF GV_FLAG_BALANCE IS INITIAL.

    MESSAGE S024(MSITEM) WITH LV_TRANS_COUNT.
    CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
      EXPORTING
        I_CALLBACK_PROGRAM       = SY-CPROG
        I_CALLBACK_PF_STATUS_SET = 'SET_PF_STATUS'
        I_GRID_TITLE             = LV_TRANS_TITLE
        IS_LAYOUT                = GS_LAYOUT
        IT_FIELDCAT              = GT_FIELDCATALOG_TRANS
        IT_EVENTS                = GT_EVENTS
      TABLES
        T_OUTTAB                 = GT_FIAPPTD_TRANS
      EXCEPTIONS
        PROGRAM_ERROR            = 1
        OTHERS                   = 2.

    IF SY-SUBRC <> 0.
* Implement suitable error handling here
    ENDIF.

  ELSEIF GV_FLAG_BALANCE = 'X'.

    MESSAGE S024(MSITEM) WITH LV_BAL_COUNT.
    CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
      EXPORTING
        I_CALLBACK_PROGRAM       = SY-CPROG
        I_CALLBACK_PF_STATUS_SET = 'SET_BAL_PF_STATUS'
        I_STRUCTURE_NAME         = 'FIAPPTD_BALANCE'
        I_GRID_TITLE             = LV_BAL_TITLE
        IS_LAYOUT                = GS_LAYOUT
        IT_FIELDCAT              = GT_FIELDCATALOG_BALANCES
        IT_EVENTS                = GT_EVENTS_BAL
      TABLES
        T_OUTTAB                 = GT_FIAPPTD_BALANCE
      EXCEPTIONS
        PROGRAM_ERROR            = 1
        OTHERS                   = 2.

    IF SY-SUBRC <> 0.
* Implement suitable error handling here
    ENDIF.

  ENDIF.

  IF SY-BATCH = 'X'.

    MESSAGE S024(MSITEM) WITH LV_BAL_COUNT.
    CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
      EXPORTING
        I_CALLBACK_PROGRAM       = SY-CPROG
        I_CALLBACK_PF_STATUS_SET = 'SET_BAL_PF_STATUS'
        I_STRUCTURE_NAME         = 'FIAPPTD_BALANCE'
        I_GRID_TITLE             = LV_BAL_TITLE
        IS_LAYOUT                = GS_LAYOUT
        IT_FIELDCAT              = GT_FIELDCATALOG_BALANCES
        IT_EVENTS                = GT_EVENTS_BAL
      TABLES
        T_OUTTAB                 = GT_FIAPPTD_BALANCE
      EXCEPTIONS
        PROGRAM_ERROR            = 1
        OTHERS                   = 2.
  ENDIF.

ENDFORM.                    " ALV_DISPLAY
*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM USER_COMMAND USING P_UCOMM TYPE SY-UCOMM
       P_SELFIELD TYPE SLIS_SELFIELD.

  DATA: LS_TRANS LIKE LINE OF GT_FIAPPTD_TRANS.

  IF P_UCOMM = 'BALANCE'.
    GV_FLAG_BALANCE = 'X'.
    PERFORM ALV_DISPLAY.
  ELSEIF P_UCOMM = 'BCK'.
*    LEAVE TO TRANSACTION gv_trans.
    LEAVE TO SCREEN 0.
  ENDIF.

  IF P_SELFIELD IS NOT INITIAL.

    READ TABLE GT_FIAPPTD_TRANS INTO LS_TRANS
       INDEX P_SELFIELD-TABINDEX.
    IF P_SELFIELD-FIELDNAME = 'BELNR'.
      SET PARAMETER ID 'BLN' FIELD LS_TRANS-BELNR.
      SET PARAMETER ID 'BUK' FIELD LS_TRANS-BUKRS.
      SET PARAMETER ID 'GJR' FIELD LS_TRANS-DOC_YEAR.
      IF LS_TRANS-BELNR IS NOT INITIAL.
        CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.
      ENDIF.
    ELSEIF P_SELFIELD-FIELDNAME = 'PAY_DOC'.
      SET PARAMETER ID 'BLN' FIELD LS_TRANS-PAY_DOC.
      SET PARAMETER ID 'BUK' FIELD LS_TRANS-BUKRS.
      SET PARAMETER ID 'GJR' FIELD LS_TRANS-PAY_DOC_YEAR.
      IF LS_TRANS-PAY_DOC IS NOT INITIAL.
        CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.
      ENDIF.
    ENDIF.

  ENDIF.

ENDFORM.                    "USER_COMMAND
*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND_BAL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM USER_COMMAND_BAL USING P_UCOMM TYPE SY-UCOMM
       P_SELFIELD TYPE SLIS_SELFIELD.

  IF P_UCOMM = 'TRANS'.
    CLEAR : GV_FLAG_BALANCE.
    PERFORM ALV_DISPLAY.
  ENDIF.

ENDFORM.                    "USER_COMMAND_BAL
*&---------------------------------------------------------------------*
*&      Form  MODIFY_FIELDCAT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM MODIFY_FIELDCAT .

  DATA: LS_FIELDCAT_TRANS LIKE LINE OF GT_FIELDCATALOG_TRANS,
        LS_FIELDCAT_BAL   LIKE LINE OF GT_FIELDCATALOG_BALANCES.

*Modify field catalog to display Foreign Transactions
  LOOP AT GT_FIELDCATALOG_TRANS INTO LS_FIELDCAT_TRANS.
    IF LS_FIELDCAT_TRANS-FIELDNAME = 'MANDT'.
      LS_FIELDCAT_TRANS-NO_OUT = 'X'.
      MODIFY GT_FIELDCATALOG_TRANS FROM LS_FIELDCAT_TRANS.
    ELSEIF LS_FIELDCAT_TRANS-FIELDNAME = 'BELNR'.
      LS_FIELDCAT_TRANS-HOTSPOT = 'X'.
      MODIFY GT_FIELDCATALOG_TRANS FROM LS_FIELDCAT_TRANS.
    ELSEIF LS_FIELDCAT_TRANS-FIELDNAME = 'PAY_DOC'.
      LS_FIELDCAT_TRANS-SELTEXT_M = TEXT-013.
      LS_FIELDCAT_TRANS-DDICTXT = 'M'.
      LS_FIELDCAT_TRANS-HOTSPOT = 'X'.
      MODIFY GT_FIELDCATALOG_TRANS FROM LS_FIELDCAT_TRANS.
    ELSEIF LS_FIELDCAT_TRANS-FIELDNAME = 'REPORTING_YEAR'.
      LS_FIELDCAT_TRANS-SELTEXT_M = TEXT-010.
      LS_FIELDCAT_TRANS-DDICTXT = 'M'.
      MODIFY GT_FIELDCATALOG_TRANS FROM LS_FIELDCAT_TRANS.
    ELSEIF LS_FIELDCAT_TRANS-FIELDNAME = 'ORIGIN_RUN_ID'.
      LS_FIELDCAT_TRANS-SELTEXT_M = TEXT-011.
      LS_FIELDCAT_TRANS-DDICTXT = 'M'.
      MODIFY GT_FIELDCATALOG_TRANS FROM LS_FIELDCAT_TRANS.
    ENDIF.
  ENDLOOP.

*Modify field catalog for Balances
  LOOP AT GT_FIELDCATALOG_BALANCES INTO LS_FIELDCAT_BAL.
    IF LS_FIELDCAT_BAL-FIELDNAME = 'MANDT'.
      LS_FIELDCAT_BAL-NO_OUT = 'X'.
      MODIFY GT_FIELDCATALOG_BALANCES FROM LS_FIELDCAT_BAL.
    ELSEIF LS_FIELDCAT_BAL-FIELDNAME = 'ORIGIN_RUN_ID'.
      LS_FIELDCAT_BAL-SELTEXT_M = TEXT-011.
      LS_FIELDCAT_BAL-DDICTXT = 'M'.
      MODIFY GT_FIELDCATALOG_BALANCES FROM LS_FIELDCAT_BAL.
    ELSEIF LS_FIELDCAT_BAL-FIELDNAME = 'REFERENCE_DATE'.
      LS_FIELDCAT_BAL-SELTEXT_M = TEXT-012.
      LS_FIELDCAT_BAL-DDICTXT = 'M'.
      MODIFY GT_FIELDCATALOG_BALANCES FROM LS_FIELDCAT_BAL.
    ELSEIF LS_FIELDCAT_BAL-FIELDNAME = 'DUE_DATE'.
      LS_FIELDCAT_BAL-NO_OUT = 'X'.
      MODIFY GT_FIELDCATALOG_BALANCES FROM LS_FIELDCAT_BAL.
    ELSEIF LS_FIELDCAT_BAL-FIELDNAME = 'PAIS_ATIVO'.
      LS_FIELDCAT_BAL-NO_OUT = 'X'.
      MODIFY GT_FIELDCATALOG_BALANCES FROM LS_FIELDCAT_BAL.
    ENDIF.
  ENDLOOP.

ENDFORM.                    " MODIFY_FIELDCAT
*&---------------------------------------------------------------------*
*&      Form  GET_BAD_INSTANCE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM GET_BADI_INSTANCE .
  GET BADI OBJ_CBR_HANDLE.
ENDFORM.                    " GET_BAD_INSTANCE
*&---------------------------------------------------------------------*
*&      Form  CLEAR_GLOBAL_VARIABLES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CLEAR_GLOBAL_VARIABLES .

  CLEAR: GT_PAID_ITEMS,
         GT_BSEG_PAY_GL,
         GT_BSEG_PAY_VC,
         GT_BSEG_REM_ALL,
         GT_BSEG_PAY_VC_VEN,
         GT_BSEG_PAY_VC_CUS,
         GT_BSEG_PAY_VC_GL.

ENDFORM.                    " CLEAR_GLOBAL_VARIABLES
*&---------------------------------------------------------------------*
*&      Form  GET_PAYMENT_DOCS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM GET_PAYMENT_DOCS .

  DATA: LS_BUDAT_RANGE LIKE LINE OF LT_BUDAT_RANGE,
        LV_YEAR        TYPE BAPI0002_4-FISCAL_YEAR,
        LV_PERIOD      TYPE BAPI0002_4-FISCAL_PERIOD.

  DATA: LV_CURSOR       TYPE CURSOR,
        LV_PACKAGE_SIZE TYPE I VALUE 30000,
        LT_BKPF         TYPE TABLE OF BKPF,
        LT_BSEG         TYPE TABLE OF BSEG,
        LS_BKPF         TYPE BKPF,
        LS_BSEG         TYPE BSEG.

  DATA: LV_INDEX TYPE SY-TABIX VALUE 1.

  READ TABLE LT_BUDAT_RANGE INTO LS_BUDAT_RANGE INDEX 1.

  CALL FUNCTION 'BAPI_COMPANYCODE_GET_PERIOD'
    EXPORTING
      COMPANYCODEID = P_BUKRS
      POSTING_DATE  = LS_BUDAT_RANGE-LOW
    IMPORTING
      FISCAL_YEAR   = LV_YEAR
      FISCAL_PERIOD = LV_PERIOD.
****Filter and optimization

** Account Logic*********

 DATA :  FILT_VEND     TYPE TABLE OF LFA1-LIFNR,
         FILT_CUST     TYPE TABLE OF KNA1-KUNNR,
         LS_CUST       TYPE KNA1,
         LS_VEND       TYPE LFA1,
         FILT_HKONT    TYPE RANGE OF HKONT,
         LS_FILT_HKONT LIKE LINE OF FILT_HKONT,
         LT_SKB1       TYPE TABLE OF SKB1,
         LS_SKB1       TYPE SKB1.

  REFRESH : FILT_VEND, FILT_CUST,FILT_HKONT,LT_SKB1.

  CALL METHOD OBJ_CBR->GET_ACCOUNTS
    IMPORTING
      ET_SKB1 = LT_SKB1.

  LOOP AT LT_SKB1 INTO LS_SKB1.
    LS_FILT_HKONT-LOW = LS_SKB1-SAKNR .
    LS_FILT_HKONT-OPTION = 'EQ' .
    LS_FILT_HKONT-SIGN = 'I' .
    APPEND LS_FILT_HKONT TO FILT_HKONT.
    CLEAR :LS_SKB1,LS_FILT_HKONT.
  ENDLOOP.

  OPEN CURSOR LV_CURSOR FOR  SELECT * FROM BKPF AS BKPF
                         INNER  JOIN   BSEG AS BSEG
                         ON  BKPF~BUKRS = BSEG~BUKRS
                         AND BKPF~BELNR = BSEG~BELNR
                         AND BKPF~GJAHR = BSEG~GJAHR

  WHERE
   BKPF~BUKRS EQ P_BUKRS AND
   BKPF~BELNR IN P_BELNR[] AND
   BKPF~GJAHR EQ LV_YEAR AND
   BKPF~BLART IN P_BLART[] AND
   BKPF~BUDAT IN LT_BUDAT_RANGE[] AND
   BSEG~HKONT IN FILT_HKONT[]
   ORDER BY BKPF~BUKRS BKPF~BELNR BKPF~GJAHR.

  DO.
    CLEAR: LT_BKPF, LT_BSEG.
    FETCH NEXT CURSOR LV_CURSOR INTO CORRESPONDING FIELDS OF TABLE
       LT_BKPF PACKAGE SIZE LV_PACKAGE_SIZE.

    IF SY-SUBRC NE 0.
      EXIT.
    ENDIF.

    IF LT_BKPF IS NOT INITIAL.
      CLEAR: LV_INDEX.
      LV_INDEX = 1.

      DELETE ADJACENT DUPLICATES FROM LT_BKPF COMPARING BUKRS BELNR GJAHR.

      SELECT * FROM BSEG INTO TABLE LT_BSEG FOR ALL ENTRIES IN  LT_BKPF[]
        WHERE BUKRS EQ  P_BUKRS AND
              BELNR EQ LT_BKPF-BELNR AND
              GJAHR EQ LV_YEAR
*              ktosl ne 'WIT'
              ORDER BY PRIMARY KEY .

      LOOP AT LT_BKPF INTO LS_BKPF.

        LOOP AT LT_BSEG INTO LS_BSEG FROM LV_INDEX.

          IF LS_BKPF-BELNR NE LS_BSEG-BELNR.
            LV_INDEX = SY-TABIX.
            EXIT.
          ENDIF.

          CALL METHOD OBJ_CBR->GET_BSEG
            EXPORTING
              I_BKPF         = LS_BKPF
              I_BSEG         = LS_BSEG
            IMPORTING
              E_BSEG_VC      = GT_BSEG_PAY_VC
              E_BSEG_GL      = GT_BSEG_PAY_GL
              E_BSEG_VC_VEN  = GT_BSEG_PAY_VC_VEN
              E_BSEG_VC_CUS  = GT_BSEG_PAY_VC_CUS
              E_BSEG_VC_GL   = GT_BSEG_PAY_VC_GL
              E_BSEG_REM_ALL = GT_BSEG_REM_ALL.

        ENDLOOP.
      ENDLOOP.

    ENDIF.

**** Only In case of only filter by customer and vendor on selection screen
    IF P_VEND IS NOT INITIAL OR P_CUST IS NOT INITIAL .
      CALL METHOD OBJ_CBR->FILTER_CVP
        EXPORTING
          RT_LFA1         = P_VEND[]
          RT_KNA1         = P_CUST[]
        CHANGING
          CT_BSEG_VC      = GT_BSEG_PAY_VC
          CT_BSEG_GL      = GT_BSEG_PAY_GL
          CT_BSEG_VC_VEN  = GT_BSEG_PAY_VC_VEN
          CT_BSEG_VC_CUS  = GT_BSEG_PAY_VC_CUS
          CT_BSEG_VC_GL   = GT_BSEG_PAY_VC_GL
          CT_BSEG_REM_ALL = GT_BSEG_REM_ALL.
    ENDIF .

    CALL METHOD OBJ_CBR->GET_PAID_DOC_FROM_BSEG
      EXPORTING
        IT_PAYMENT_DOCS     = GT_BSEG_PAY_GL
        IT_PAYMENT_DOCS_VEN = GT_BSEG_PAY_VC_VEN
        IT_PAYMENT_DOCS_CUS = GT_BSEG_PAY_VC_CUS
        IT_PAYMENT_DOCS_GL  = GT_BSEG_PAY_VC_GL
        IT_POSTING_DATE     = LT_BUDAT_RANGE
        IV_SKIP_GL_ITEMS    = P_GL_PAY
      IMPORTING
        ET_PAID_ITEMS       = GT_PAID_ITEMS
      EXCEPTIONS
        NO_DATA             = 1
        OTHERS              = 2.

    CALL METHOD OBJ_CBR->PROCESS_PMNT_TRANS
      EXPORTING
        IT_LFA1          = GT_LFA1
        IT_KNA1          = GT_KNA1
        IT_PAID_ITEMS    = GT_PAID_ITEMS
        IV_AUTO_TO_MAN   = P_AUTO
      CHANGING
        ET_FIAPPTD_TRANS = GT_FIAPPTD_TRANS
        ET_XML_ITEMS     = GT_XML_ITEMS_TRANS
        IT_BSEG_PAY_GL   = GT_BSEG_PAY_GL
        IT_BSEG_PAY_VC   = GT_BSEG_PAY_VC
        IT_BSEG_REM_ALL  = GT_BSEG_REM_ALL.

    PERFORM CLEAR_GLOBAL_VARIABLES.

  ENDDO.
  CLOSE CURSOR LV_CURSOR.


ENDFORM.                    " GET_PAYMENT_DOCS
*&---------------------------------------------------------------------*
*&      Form  LOC_BROWSER_F4
*&---------------------------------------------------------------------*
*       F4 help for file location
*----------------------------------------------------------------------*
FORM LOC_BROWSER_F4 CHANGING PV_PATH TYPE RLGRAP-FILENAME.

  DATA: LV_PATH  TYPE STRING,
        LV_TITLE TYPE STRING.

  LV_TITLE = TEXT-031.

  CALL METHOD CL_GUI_FRONTEND_SERVICES=>DIRECTORY_BROWSE
    EXPORTING
      WINDOW_TITLE    = LV_TITLE
    CHANGING
      SELECTED_FOLDER = LV_PATH
    EXCEPTIONS
      OTHERS          = 0.
  PV_PATH = LV_PATH.
ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  CHECK_LOCAL_PATH
*&---------------------------------------------------------------------*
FORM CHECK_LOCAL_PATH CHANGING PV_PATH TYPE RLGRAP-FILENAME.
  DATA:
    LV_EXIST     TYPE ABAP_BOOL,
    LV_DIRECTORY TYPE STRING.

  IF NOT PV_PATH IS INITIAL.
    LV_DIRECTORY = PV_PATH.
    LV_EXIST = CL_GUI_FRONTEND_SERVICES=>DIRECTORY_EXIST( LV_DIRECTORY ).
    IF LV_EXIST = ABAP_FALSE.
      MESSAGE E024 WITH PV_PATH.
    ENDIF.
  ENDIF.

ENDFORM.