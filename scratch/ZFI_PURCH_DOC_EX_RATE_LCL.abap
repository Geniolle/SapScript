*----------------------------------------------------------------------*
***INCLUDE ZFI_PURCH_DOC_EX_RATE_LCL.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Class ZFI_PURCH_DOC_EX_RATE_LCL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
CLASS LCL_X_RATE DEFINITION FINAL.

  PUBLIC SECTION.
    METHODS:
      CONSTRUCTOR
        RAISING
          ZCX_BC_EXCEPTIONS,

      "Auto assign method
      AUTO_INST_ASSIGN
        IMPORTING
          IV_BUKRS TYPE BUKRS
          IV_WAERS TYPE WAERS
          ID_IDATE TYPE DATUM
          ID_EDATE TYPE DATUM
          IR_EBELN TYPE ANY
          IR_BEDAT TYPE ANY
          IR_AEDAT TYPE ANY
          IV_KONT  TYPE BP_PARTNR_NEW
        RAISING
          ZCX_BC_EXCEPTIONS,
      "View Log method
      SHOW_LOG
        IMPORTING
          IR_ADATE TYPE ANY
          IR_COMP  TYPE ANY
          IR_WER   TYPE ANY
          IR_RF    TYPE ANY
          IR_EBL   TYPE ANY
          IR_KON   TYPE ANY
        RAISING
          ZCX_BC_EXCEPTIONS,

      "Manual assign method
      MANUAL_INST_ASSIGN
        IMPORTING
          IV_BUKR TYPE BUKRS
          IV_WAER TYPE WAERS
          IV_RFHA TYPE TB_RFHA
          IR_EBEL TYPE ANY
        RAISING
          ZCX_BC_EXCEPTIONS.

  PRIVATE SECTION.
    "Private Section

    DATA: GV_BUKRS   TYPE BUKRS,
          GV_KONTRH  TYPE BP_PARTNR_NEW,
          GV_RFHA    TYPE VTBFHAZU-RFHA,
          GV_SGSART  TYPE VVSART,
          GV_SFHAART TYPE TB_SFHAART,
          GV_RANTYP  TYPE RANTYP,
          GT_PDATE   TYPE TABLE OF TY_PAY_DATE,
          GT_AUTO    TYPE TABLE OF TY_AUTO,
          GT_MANUAL  TYPE TABLE OF TY_MANUAL,
          GT_LOG     TYPE TABLE OF ZFI_DOC_EX_LOG_T,
          AUTO_ALV   TYPE REF TO CL_SALV_TABLE,
          MANUAL_ALV TYPE REF TO CL_SALV_TABLE.

    CONSTANTS: GC_AUTHOR TYPE TB_AUTHOR  VALUE 'X',
               GC_CHECK  TYPE SAP_BOOL VALUE 'X',
               GC_NCHECK TYPE SAP_BOOL VALUE ' ',
               GC_MSGID  TYPE BAL_S_MSG-MSGID VALUE 'ZFI',
               GC_OBJCT  TYPE BAL_S_LOG-OBJECT VALUE 'ZFI',
               GC_SBOBJ  TYPE BAL_S_LOG-SUBOBJECT VALUE 'ZFI_PUR_DOC_X_RATE'.

    TYPES: BEGIN OF TY_LOGEBEL,
             EBELN TYPE ZFI_DOC_EX_LOG_T-EBELN,
           END OF TY_LOGEBEL,
           TY_TLEBELN TYPE STANDARD TABLE OF TY_LOGEBEL,
           TY_REBELN  TYPE RANGE OF EBELN,
           TY_BDCDATA TYPE STANDARD TABLE OF BDCDATA.

    METHODS
      "display auto assign alv
      DISPLAY_AUTO_ALV
        RAISING
          ZCX_BC_EXCEPTIONS.

    METHODS
      "display manual assign alv
      DISPLAY_MANUAL_ALV
        RAISING
          ZCX_BC_EXCEPTIONS.


    METHODS
      "display log alv
      DISPLAY_LOG_ALV
        RAISING
          ZCX_BC_EXCEPTIONS.

    METHODS
      "event for button click
      ON_USER_COMMAND
        FOR EVENT ADDED_FUNCTION OF CL_SALV_EVENTS
        IMPORTING
          E_SALV_FUNCTION.

    METHODS
      "event for checkbox click
      ON_CLICK
        FOR EVENT LINK_CLICK OF CL_SALV_EVENTS_TABLE
        IMPORTING
          ROW.

    METHODS
      "update document exchange rate
      CHANGE_EX_RATE
        IMPORTING
          IV_EBELN TYPE EBELN
          IV_BUKRS TYPE BUKRS
          IV_KKURS TYPE TB_KKURS
        RAISING
          ZCX_BC_EXCEPTIONS.

    METHODS
      "Update log method
      UPD_LOG
        IMPORTING
          IS_AUTO   TYPE TY_AUTO OPTIONAL
          IS_MANUAL TYPE TY_MANUAL OPTIONAL
        RAISING
          ZCX_BC_EXCEPTIONS.

    METHODS
      SET_LOG_MESSAGE
        IMPORTING
          IV_LEVEL TYPE BAL_S_MSG-DETLEVEL
          IV_MSGTY TYPE BAL_S_MSG-MSGTY
          IV_MSGID TYPE BAL_S_MSG-MSGID
          IV_MSGNO TYPE BAL_S_MSG-MSGNO
          IV_MSGV1 TYPE BAL_S_MSG-MSGV1 OPTIONAL
          IV_MSGV2 TYPE BAL_S_MSG-MSGV2 OPTIONAL
          IV_MSGV3 TYPE BAL_S_MSG-MSGV3 OPTIONAL
          IV_MSGV4 TYPE BAL_S_MSG-MSGV4 OPTIONAL
          IV_SAVES TYPE XFELD OPTIONAL
        CHANGING
          CO_EXLOG TYPE REF TO ZCLCA_BAL_LOG.

    METHODS
      BDC_DYNPRO
        IMPORTING
          IV_PROG    TYPE BDCDATA-PROGRAM
          IV_SRC     TYPE BDCDATA-DYNPRO
        CHANGING
          CT_BDCDATA TYPE TY_BDCDATA.

    METHODS
      BDC_FIELD
        IMPORTING
          IV_FNAM    TYPE BDCDATA-FNAM
          IV_FVAL    TYPE ANY
        CHANGING
          CT_BDCDATA TYPE TY_BDCDATA.

    METHODS
      CREATE_RANGE
        IMPORTING
          IT_EXLOG TYPE TY_TLEBELN
        EXPORTING
          ER_EBELN TYPE TY_REBELN.

ENDCLASS.

CLASS LCL_X_RATE IMPLEMENTATION.

  METHOD CONSTRUCTOR.

    CLEAR: GV_BUKRS, GV_KONTRH, GV_RFHA, GV_RANTYP, GV_SGSART, GV_SFHAART, GT_PDATE, GT_AUTO, GT_MANUAL, GT_LOG, AUTO_ALV,
           MANUAL_ALV.


    "Get range of user langs accepted
    CALL METHOD ZCLCA_FIXEDVALS=>GET_CONS_VAL
      EXPORTING
        IV_BUKRS = ''
        IV_MODUL = ZTGCA_C_MOD_FIN
        IV_PROCE = 'PO_EXCHANGE_RATE'
        IV_FNAME = 'SGSART'
        IV_SEQUE = 1
      IMPORTING
        EV_CONST = GV_SGSART
      EXCEPTIONS
        NO_DATA  = 1
        OTHERS   = 2.

    IF SY-SUBRC <> 0.

      IF 1 = 2. MESSAGE E086(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '086'.

    ENDIF.

    "Get range of user langs accepted
    CALL METHOD ZCLCA_FIXEDVALS=>GET_CONS_VAL
      EXPORTING
        IV_BUKRS = ''
        IV_MODUL = ZTGCA_C_MOD_FIN
        IV_PROCE = 'PO_EXCHANGE_RATE'
        IV_FNAME = 'SFHAART'
        IV_SEQUE = 1
      IMPORTING
        EV_CONST = GV_SFHAART
      EXCEPTIONS
        NO_DATA  = 1
        OTHERS   = 2.
    IF SY-SUBRC <> 0.

      IF 1 = 2. MESSAGE E086(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '086'.

    ENDIF.

    "Get range of user langs accepted
    CALL METHOD ZCLCA_FIXEDVALS=>GET_CONS_VAL
      EXPORTING
        IV_BUKRS = ''
        IV_MODUL = ZTGCA_C_MOD_FIN
        IV_PROCE = 'PO_EXCHANGE_RATE'
        IV_FNAME = 'RANTYP'
        IV_SEQUE = 1
      IMPORTING
        EV_CONST = GV_RANTYP
      EXCEPTIONS
        NO_DATA  = 1
        OTHERS   = 2.
    IF SY-SUBRC <> 0.

      IF 1 = 2. MESSAGE E086(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '086'.

    ENDIF.

  ENDMETHOD.

  METHOD AUTO_INST_ASSIGN.

    DATA: LR_EBELN  TYPE RANGE OF EBELN,
          LR_BEDAT  TYPE RANGE OF BEDAT,
          LR_AEDAT  TYPE RANGE OF AEDAT,
          LR_LEBELN TYPE RANGE OF EBELN,
          PAY_DATE  TYPE DATUM,
          LV_EBELN  TYPE EKKO-EBELN,
          LV_TABIX  TYPE SY-TABIX,
          LV_NETWR  TYPE EKPO-NETWR,
          LS_AUTOF  TYPE TY_AUTO,
          LS_PDATE  TYPE TY_PAY_DATE.

    LR_EBELN = IR_EBELN.
    LR_BEDAT = IR_BEDAT.
    LR_AEDAT = IR_AEDAT.

    "Populate global variables
    GV_BUKRS = IV_BUKRS.
    GV_KONTRH = IV_KONT.

    "Validate selected partner
    SELECT A~RANTYP,
           A~AUTHOR
    FROM VTBSTA3 AS A
    INNER JOIN TZPA AS B ON B~GSART = A~SGSART
    INTO TABLE @DATA(LT_VTBSTA3) ##NEEDED
    UP TO 1 ROWS
    WHERE A~BUKRS EQ @IV_BUKRS
    AND A~PARTNR EQ @IV_KONT
    AND A~RANTYP EQ @GV_RANTYP
    AND A~SANLF EQ B~SANLF
    AND A~SGSART EQ @GV_SGSART
    AND A~SFHAART EQ @GV_SFHAART
    AND A~AUTHOR EQ @GC_AUTHOR.

    IF SY-SUBRC IS NOT INITIAL.
      IF 1 = 2. MESSAGE E071(ZFI) WITH GV_KONTRH. ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '071'.

    ELSEIF SY-SUBRC EQ 0.

      "Get the data to display in the alv/Validation of the selected data
      SELECT A~BUKRS,
             A~WAERS,
             A~EBELN,
             B~EINDT,
             A~LIFNR,
             E~NDAYS,
             F~ZTAG1,
             G~NETWR,
             CASE WHEN G~ZZDAT03 IS NOT NULL THEN G~ZZDAT03 ELSE G~ZZDAT02 END AS ZZDAT02 "MPereira 20260108 SAP77
      FROM EKKO AS A
      INNER JOIN EKET AS B ON B~EBELN EQ A~EBELN
      INNER JOIN LFB1 AS D ON D~LIFNR EQ A~LIFNR AND D~BUKRS EQ A~BUKRS
      INNER JOIN ZFI_PAY_DATE_T AS E ON E~ZZCOORI EQ A~ZZCOORI AND E~ZZEXPVZ EQ A~ZZEXPVZ
      INNER JOIN T052 AS F ON F~ZTERM EQ D~ZTERM
      INNER JOIN EKPO AS G ON G~EBELN = A~EBELN AND G~EBELP = B~EBELP
      INTO TABLE @DATA(LT_AUTO)
      WHERE A~BUKRS EQ @IV_BUKRS
      AND A~WAERS EQ @IV_WAERS
      AND A~EBELN IN @LR_EBELN
      AND A~BEDAT IN @LR_BEDAT
      AND A~AEDAT IN @LR_AEDAT
      ORDER BY A~EBELN.

      IF SY-SUBRC IS NOT INITIAL.

        IF 1 = 2. MESSAGE E053(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '053'.

      ELSEIF SY-SUBRC EQ 0.

        SORT LT_AUTO BY EBELN.

        SELECT DISTINCT EBELN
        FROM ZFI_DOC_EX_LOG_T
        INTO TABLE @DATA(LT_EXLOG)
        FOR ALL ENTRIES IN @LT_AUTO
        WHERE BUKRS EQ @LT_AUTO-BUKRS
        AND EBELN EQ @LT_AUTO-EBELN.

        IF SY-SUBRC EQ 0.

          CREATE_RANGE(
            EXPORTING
              IT_EXLOG = LT_EXLOG
            IMPORTING
              ER_EBELN = LR_LEBELN ).

          DELETE LT_AUTO WHERE EBELN IN LR_LEBELN.

        ENDIF.

        DATA(LV_ROWS) = LINES( LT_AUTO ).

        "Loop at auto table to calculate the payment date field
        LOOP AT LT_AUTO ASSIGNING FIELD-SYMBOL(<FS_AUTO>).

          LV_TABIX = SY-TABIX.

          CLEAR: LS_AUTOF, LS_PDATE.

          IF LV_EBELN IS INITIAL.
            LV_EBELN = <FS_AUTO>-EBELN.
            PAY_DATE = <FS_AUTO>-ZZDAT02 - <FS_AUTO>-NDAYS + <FS_AUTO>-ZTAG1. "MPEREIRA 20260108 SAP 77 - <fs_auto>-eindt - <fs_auto>-ndays + <fs_auto>-ztag1.
          ENDIF.

          IF LV_EBELN EQ <FS_AUTO>-EBELN.
            LV_NETWR = LV_NETWR + <FS_AUTO>-NETWR.
          ELSE.
            "Validate if payment date is in the range defined in the select options
            IF PAY_DATE GE ID_IDATE AND PAY_DATE LE ID_EDATE.

              MOVE-CORRESPONDING LT_AUTO[ SY-TABIX - 1 ] TO LS_AUTOF. " MPereira 12.01.2026 SAP77 Retificação preechimento ALV

              LS_AUTOF-PDATE = PAY_DATE.
              LS_AUTOF-CHECK = GC_NCHECK.
              LS_AUTOF-NETWR = LV_NETWR.
*              ls_autof-ebeln = lv_ebeln.
*              TRY.
*                  ls_autof-lifnr = lt_auto[ ebeln = lv_ebeln ]-lifnr.
*                CATCH cx_sy_itab_line_not_found.
*              ENDTRY.


              APPEND LS_AUTOF TO GT_AUTO.

              LS_PDATE-EBELN = LV_EBELN. "FIX <fs_auto>-ebeln.
              LS_PDATE-PDATE = PAY_DATE.

              APPEND LS_PDATE TO GT_PDATE.

            ENDIF.

            CLEAR: PAY_DATE, LV_NETWR, LV_EBELN.

            "Calculate payment date
            PAY_DATE = <FS_AUTO>-ZZDAT02 - <FS_AUTO>-NDAYS + <FS_AUTO>-ZTAG1. "MPEREIRA 20260108 SAP 77 - <fs_auto>-eindt - <fs_auto>-ndays + <fs_auto>-ztag1.
            LV_EBELN = <FS_AUTO>-EBELN.
            LV_NETWR = LV_NETWR + <FS_AUTO>-NETWR.

          ENDIF.

          IF LV_TABIX EQ LV_ROWS.

            "Validate if payment date is in the range defined in the select options
            IF PAY_DATE GE ID_IDATE AND PAY_DATE LE ID_EDATE.

              MOVE-CORRESPONDING <FS_AUTO> TO LS_AUTOF.

              LS_AUTOF-PDATE = PAY_DATE.
              LS_AUTOF-CHECK = GC_NCHECK.
              LS_AUTOF-NETWR = LV_NETWR.

              APPEND LS_AUTOF TO GT_AUTO.

              LS_PDATE-EBELN = <FS_AUTO>-EBELN.
              LS_PDATE-PDATE = PAY_DATE.

              APPEND LS_PDATE TO GT_PDATE.

            ENDIF.
          ENDIF.

        ENDLOOP.
      ENDIF.

      DELETE ADJACENT DUPLICATES FROM GT_AUTO COMPARING EBELN.

      TRY.
          DISPLAY_AUTO_ALV( ).

        CATCH ZCX_BC_EXCEPTIONS INTO DATA(LO_EXCPT).
          "espoletar exceção
          RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
            EXPORTING
              MSGTY = LO_EXCPT->MSGTY
              MSGID = LO_EXCPT->MSGID
              MSGNO = LO_EXCPT->MSGNO
              MSGV1 = LO_EXCPT->MSGV1
              MSGV2 = LO_EXCPT->MSGV2
              MSGV3 = LO_EXCPT->MSGV3
              MSGV4 = LO_EXCPT->MSGV4.
      ENDTRY.

    ENDIF.

  ENDMETHOD.

  METHOD SHOW_LOG.

    DATA: LR_ADATA TYPE RANGE OF DATUM,
          LR_COMP  TYPE RANGE OF BUKRS,
          LR_WER   TYPE RANGE OF WAERS,
          LR_RF    TYPE RANGE OF TB_RFHA,
          LR_EBL   TYPE RANGE OF EBELN,
          LR_KON   TYPE RANGE OF TB_KUNNR_NEW.

    LR_ADATA = IR_ADATE.
    LR_COMP  = IR_COMP.
    LR_WER   = IR_WER.
    LR_RF    = IR_RF.
    LR_EBL   = IR_EBL.
    LR_KON   = IR_KON.

    "Filter log data based on selection screen
    SELECT *
    FROM ZFI_DOC_EX_LOG_T
    INTO TABLE @GT_LOG
    WHERE ASSIGN_DATE IN @LR_ADATA
    AND BUKRS  IN @LR_COMP
    AND RFHA   IN @LR_RF
    AND EBELN  IN @LR_EBL
    AND KONTRH IN @LR_KON
    AND WAERS  IN @LR_WER
    ORDER BY ASSIGN_DATE, BUKRS, RFHA.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E053(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '053'.

    ENDIF.

    TRY.
        DISPLAY_LOG_ALV( ).

      CATCH ZCX_BC_EXCEPTIONS INTO DATA(LO_EXCPT).
        "espoletar exceção
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = LO_EXCPT->MSGTY
            MSGID = LO_EXCPT->MSGID
            MSGNO = LO_EXCPT->MSGNO
            MSGV1 = LO_EXCPT->MSGV1
            MSGV2 = LO_EXCPT->MSGV2
            MSGV3 = LO_EXCPT->MSGV3
            MSGV4 = LO_EXCPT->MSGV4.
    ENDTRY.

  ENDMETHOD.

  METHOD MANUAL_INST_ASSIGN.

    DATA: LR_EBELN  TYPE RANGE OF EBELN,
          LR_LEBELN TYPE RANGE OF EBELN,
          LS_MANUAL TYPE TY_MANUAL.

    GV_BUKRS = IV_BUKR.
    GV_RFHA = IV_RFHA.
    LR_EBELN = IR_EBEL.

    "Validate sgsart and sfhaart corresponding to the bukrs and rfha chosen
    SELECT SGSART,
           SFHAART
    FROM VTBFHA
    INTO TABLE @DATA(LT_VTBFHA) ##NEEDED
    UP TO 1 ROWS
    WHERE BUKRS EQ @IV_BUKR
    AND RFHA EQ @IV_RFHA
    AND SGSART EQ @GV_SGSART
    AND SFHAART EQ @GV_SFHAART.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E072(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '072'.

    ELSEIF SY-SUBRC EQ 0.

      "Get data for manual alv
      SELECT BUKRS,
             WAERS,
             EBELN
      FROM EKKO
      INTO TABLE @DATA(LT_EKKO)
      WHERE EBELN IN @LR_EBELN
      AND BUKRS EQ @IV_BUKR
      AND WAERS EQ @IV_WAER.

      IF SY-SUBRC IS NOT INITIAL.

        IF 1 = 2. MESSAGE E053(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '053'.

      ELSEIF SY-SUBRC EQ 0.

        SELECT EBELN
        FROM ZFI_DOC_EX_LOG_T
        INTO TABLE @DATA(LT_EXLOG)
        FOR ALL ENTRIES IN @LT_EKKO
        WHERE BUKRS EQ @LT_EKKO-BUKRS
        AND EBELN EQ @LT_EKKO-EBELN.


        IF SY-SUBRC EQ 0.

          CREATE_RANGE(
            EXPORTING
              IT_EXLOG = LT_EXLOG
            IMPORTING
              ER_EBELN = LR_LEBELN ).

          DELETE LT_EKKO WHERE EBELN IN LR_LEBELN.

        ENDIF.

        "Fill remaining data
        LOOP AT LT_EKKO ASSIGNING FIELD-SYMBOL(<FS_EKKO>).

          CLEAR LS_MANUAL.

          MOVE-CORRESPONDING <FS_EKKO> TO LS_MANUAL.

          LS_MANUAL-CHECK = GC_NCHECK.
          LS_MANUAL-RFHA = IV_RFHA.

          APPEND LS_MANUAL TO GT_MANUAL.

        ENDLOOP.

        TRY.

            DISPLAY_MANUAL_ALV( ).

          CATCH ZCX_BC_EXCEPTIONS INTO DATA(LO_EXCPT).
            "espoletar exceção
            RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
              EXPORTING
                MSGTY = LO_EXCPT->MSGTY
                MSGID = LO_EXCPT->MSGID
                MSGNO = LO_EXCPT->MSGNO
                MSGV1 = LO_EXCPT->MSGV1
                MSGV2 = LO_EXCPT->MSGV2
                MSGV3 = LO_EXCPT->MSGV3
                MSGV4 = LO_EXCPT->MSGV4.
        ENDTRY.

      ENDIF.
    ENDIF.
  ENDMETHOD.

  METHOD DISPLAY_AUTO_ALV.

    DATA: LO_FUNCTIONS  TYPE REF TO CL_SALV_FUNCTIONS_LIST,
          GR_SELECTIONS TYPE REF TO CL_SALV_SELECTIONS,
          GR_EVENTS     TYPE REF TO CL_SALV_EVENTS_TABLE,
          LR_COLUMNS    TYPE REF TO CL_SALV_COLUMNS_TABLE,
          LR_COLUMN     TYPE REF TO CL_SALV_COLUMN_TABLE,
          LV_COLTXT     TYPE SCRTEXT_L.

    TRY.
        CL_SALV_TABLE=>FACTORY(
      IMPORTING
        R_SALV_TABLE = AUTO_ALV
      CHANGING
        T_TABLE      = GT_AUTO ).
      CATCH CX_SALV_MSG.

        "ocorreram erros ao gerar a ALV
        IF 1 = 2. MESSAGE E054(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '054'.

    ENDTRY.

    "Change pfstatus to custom gui status
    AUTO_ALV->SET_SCREEN_STATUS(
      PFSTATUS = 'Z_AUTO_ALV_STATUS'
      REPORT   = SY-REPID
      SET_FUNCTIONS = AUTO_ALV->C_FUNCTIONS_ALL ).

    TRY.

        LR_COLUMNS = AUTO_ALV->GET_COLUMNS( ).
        LR_COLUMNS->SET_OPTIMIZE( 'X' ).

        "Set field Check as checkbox
        LR_COLUMN ?= LR_COLUMNS->GET_COLUMN( 'CHECK' ).
        LR_COLUMN->SET_CELL_TYPE( IF_SALV_C_CELL_TYPE=>CHECKBOX_HOTSPOT ).

        CLEAR LR_COLUMN.
        LV_COLTXT = TEXT-006.
        LR_COLUMN ?= LR_COLUMNS->GET_COLUMN( 'PDATE' ).
        LR_COLUMN->SET_LONG_TEXT( LV_COLTXT ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LR_COLUMN.
        LV_COLTXT = TEXT-007.
        LR_COLUMN ?= LR_COLUMNS->GET_COLUMN( 'NETWR' ).
        LR_COLUMN->SET_LONG_TEXT( LV_COLTXT ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

      CATCH CX_SALV_NOT_FOUND .
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E055(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '055'.
    ENDTRY.

    GR_EVENTS = AUTO_ALV->GET_EVENT( ).

    "Set event for custom button
    SET HANDLER ON_USER_COMMAND FOR GR_EVENTS.

    SET HANDLER ME->ON_CLICK FOR GR_EVENTS.

    GR_SELECTIONS = AUTO_ALV->GET_SELECTIONS( ).
*    gr_selections->set_selection_mode( if_salv_c_selection_mode=>row_column ).
    GR_SELECTIONS->SET_SELECTION_MODE( 1 ).

    LO_FUNCTIONS = AUTO_ALV->GET_FUNCTIONS( ).
    LO_FUNCTIONS->SET_ALL( ABAP_TRUE ).

    AUTO_ALV->DISPLAY( ).

  ENDMETHOD.

  METHOD DISPLAY_MANUAL_ALV.

    DATA: LO_FUNCTIONS  TYPE REF TO CL_SALV_FUNCTIONS_LIST,
          GR_SELECTIONS TYPE REF TO CL_SALV_SELECTIONS,
          GR_EVENTS     TYPE REF TO CL_SALV_EVENTS_TABLE,
          LR_COLUMNS    TYPE REF TO CL_SALV_COLUMNS_TABLE,
          LR_COLUMN     TYPE REF TO CL_SALV_COLUMN_TABLE.

    TRY.
        CL_SALV_TABLE=>FACTORY(
      IMPORTING
        R_SALV_TABLE = MANUAL_ALV
      CHANGING
        T_TABLE      = GT_MANUAL ).
      CATCH CX_SALV_MSG.

        "ocorreram erros ao gerar a ALV
        IF 1 = 2. MESSAGE E054(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '054'.

    ENDTRY.

    "Change pfstatus to custom gui status
    MANUAL_ALV->SET_SCREEN_STATUS(
      PFSTATUS = 'Z_AUTO_ALV_STATUS'
      REPORT   = SY-REPID
      SET_FUNCTIONS = MANUAL_ALV->C_FUNCTIONS_ALL ).

    GR_SELECTIONS = MANUAL_ALV->GET_SELECTIONS( ).
    GR_SELECTIONS->SET_SELECTION_MODE( 1 ).

    TRY.

        LR_COLUMNS = MANUAL_ALV->GET_COLUMNS( ).
        LR_COLUMNS->SET_OPTIMIZE( 'X' ).

        "Set field Check as checkbox
        LR_COLUMN ?= LR_COLUMNS->GET_COLUMN( 'CHECK' ).
        LR_COLUMN->SET_CELL_TYPE( IF_SALV_C_CELL_TYPE=>CHECKBOX_HOTSPOT ).

      CATCH CX_SALV_NOT_FOUND .
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E055(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '055'.
    ENDTRY.

    GR_EVENTS = MANUAL_ALV->GET_EVENT( ).

    SET HANDLER ME->ON_CLICK FOR GR_EVENTS.

    "Set event for custom button
    SET HANDLER ON_USER_COMMAND FOR GR_EVENTS.

    LO_FUNCTIONS = MANUAL_ALV->GET_FUNCTIONS( ).
    LO_FUNCTIONS->SET_ALL( ABAP_TRUE ).

    MANUAL_ALV->DISPLAY( ).


  ENDMETHOD.

  METHOD DISPLAY_LOG_ALV.

    DATA: LOG_ALV      TYPE REF TO CL_SALV_TABLE,
          LO_FUNCTIONS TYPE REF TO CL_SALV_FUNCTIONS_LIST,
          LR_COLUMNS   TYPE REF TO CL_SALV_COLUMNS_TABLE,
          LR_COLUMN    TYPE REF TO CL_SALV_COLUMN,
          LR_SORTS     TYPE REF TO CL_SALV_SORTS,
          LV_LTX_TEX   TYPE SCRTEXT_L.

    TRY.
        CL_SALV_TABLE=>FACTORY(
      IMPORTING
        R_SALV_TABLE = LOG_ALV
      CHANGING
        T_TABLE      = GT_LOG ).
      CATCH CX_SALV_MSG.

        "ocorreram erros ao gerar a ALV
        IF 1 = 2. MESSAGE E054(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '054'.
    ENDTRY.

    TRY.
        LR_COLUMNS = LOG_ALV->GET_COLUMNS( ).
        LR_COLUMNS->SET_OPTIMIZE( 'X' ).

        LR_COLUMN = LR_COLUMNS->GET_COLUMN( 'LTX' ).
        LV_LTX_TEX = TEXT-002.
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LV_LTX_TEX.
        CLEAR LR_COLUMN.
        LR_COLUMN = LR_COLUMNS->GET_COLUMN( 'ASSIGN_DATE' ).
        LV_LTX_TEX = TEXT-003.
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LV_LTX_TEX.
        CLEAR LR_COLUMN.
        LR_COLUMN = LR_COLUMNS->GET_COLUMN( 'FORWARD_RATE' ).
        LV_LTX_TEX = TEXT-004.
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LV_LTX_TEX.
        CLEAR LR_COLUMN.
        LR_COLUMN = LR_COLUMNS->GET_COLUMN( 'AMOUNT1' ).
        LV_LTX_TEX = TEXT-005.
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LV_LTX_TEX.
        CLEAR LR_COLUMN.
        LR_COLUMN = LR_COLUMNS->GET_COLUMN( 'PAY_DATE' ).
        LV_LTX_TEX = TEXT-006.
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

        CLEAR LV_LTX_TEX.
        CLEAR LR_COLUMN.
        LV_LTX_TEX = TEXT-007.
        LR_COLUMN ?= LR_COLUMNS->GET_COLUMN( 'NETWR' ).
        LR_COLUMN->SET_LONG_TEXT( LV_LTX_TEX ).
        LR_COLUMN->SET_SHORT_TEXT( '' ).
        LR_COLUMN->SET_MEDIUM_TEXT( '' ).

      CATCH CX_SALV_NOT_FOUND .
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E055(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '055'.
    ENDTRY.

    LO_FUNCTIONS = LOG_ALV->GET_FUNCTIONS( ).
    LO_FUNCTIONS->SET_ALL( ABAP_TRUE ).

    TRY.

        LR_SORTS = LOG_ALV->GET_SORTS( ).

        LR_SORTS->ADD_SORT( 'EXTERNAL_REFERENCE' ).

        LR_SORTS->ADD_SORT( 'FORWARD_RATE' ) .

        LR_SORTS->ADD_SORT( 'AMOUNT1' ).

      CATCH CX_SALV_NOT_FOUND.
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E076(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '076'.

      CATCH CX_SALV_DATA_ERROR.
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E076(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '076'.

      CATCH CX_SALV_EXISTING.
        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E076(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '076'.

    ENDTRY.

    LOG_ALV->DISPLAY( ).

  ENDMETHOD.

  "Method to handle custom button action
  METHOD ON_USER_COMMAND.

    DATA: IS_FTR_CRE TYPE TY_FTR_CRE,
          LV_RFHA    TYPE VTBFHAZU-RFHA,
          LS_AUTO    TYPE TY_AUTO,
          LT_BDCDATA TYPE TABLE OF BDCDATA.

    CLEAR IS_FTR_CRE.

    TRY.

        CASE E_SALV_FUNCTION.

          WHEN '&USELALL'. "Unselect all
            LS_AUTO-CHECK = GC_NCHECK.
            MODIFY GT_AUTO FROM LS_AUTO TRANSPORTING CHECK WHERE CHECK EQ GC_CHECK.

            AUTO_ALV->REFRESH( ).

          WHEN '&SELALL'. "Select all
            LS_AUTO-CHECK = GC_CHECK.
            MODIFY GT_AUTO FROM LS_AUTO TRANSPORTING CHECK WHERE CHECK EQ GC_NCHECK.

            AUTO_ALV->REFRESH( ).

          WHEN 'MYFUNC'.

            CLEAR LT_BDCDATA.

            IF GT_AUTO IS NOT INITIAL.

              IF LINE_EXISTS( GT_AUTO[ CHECK = GC_CHECK ] ).

                "Fill structure with select options for call transaction
                IS_FTR_CRE-BUKRS = GV_BUKRS.
                IS_FTR_CRE-KONTRH = GV_KONTRH.
                IS_FTR_CRE-SGSART = GV_SGSART.
                IS_FTR_CRE-SFHAART = GV_SFHAART.

                "Sum the amount of all selected fields
                LOOP AT GT_AUTO ASSIGNING FIELD-SYMBOL(<FS_FTR_CRE>) WHERE CHECK EQ GC_CHECK.

                  IS_FTR_CRE-AMOUNT = IS_FTR_CRE-AMOUNT + <FS_FTR_CRE>-NETWR.

                ENDLOOP.

                IS_FTR_CRE-SAMOUNT = IS_FTR_CRE-AMOUNT.

                CONDENSE IS_FTR_CRE-SAMOUNT NO-GAPS.

                REPLACE ALL OCCURRENCES OF '.' IN IS_FTR_CRE-SAMOUNT WITH ','.

                "Batch input fields
*                PERFORM bdc_dynpro      USING 'FTR_ENTRY' '2000'.

                BDC_DYNPRO(
                  EXPORTING
                    IV_PROG = 'FTR_ENTRY'
                    IV_SRC = '2000'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).

*                PERFORM bdc_field       USING 'BDC_CURSOR'
*                                              'FTR_ENTRY-BUKRS'.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'BDC_CURSOR'
                    IV_FVAL = 'FTR_ENTRY-BUKRS'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).

*                PERFORM bdc_field       USING 'BDC_OKCODE'
*                                              '=RETURN'.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'BDC_OKCODE'
                    IV_FVAL = '=RETURN'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'FTR_ENTRY-BUKRS'
*                                              is_ftr_cre-bukrs.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'FTR_ENTRY-BUKRS'
                    IV_FVAL = IS_FTR_CRE-BUKRS
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'FTR_ENTRY-SGSART'
*                                              is_ftr_cre-sgsart.


                BDC_FIELD(
                   EXPORTING
                     IV_FNAM = 'FTR_ENTRY-SGSART'
                     IV_FVAL = IS_FTR_CRE-SGSART
                   CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'FTR_ENTRY-SFHAART'
*                                              is_ftr_cre-sfhaart.


                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'FTR_ENTRY-SFHAART'
                    IV_FVAL = IS_FTR_CRE-SFHAART
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'FTR_ENTRY-KONTRH'
*                                              is_ftr_cre-kontrh.


                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'FTR_ENTRY-KONTRH'
                    IV_FVAL = IS_FTR_CRE-KONTRH
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_dynpro      USING 'SAPLTTM_UI_FRAMEWORK' '1100'.

                BDC_DYNPRO(
                  EXPORTING
                    IV_PROG = 'SAPLTTM_UI_FRAMEWORK'
                    IV_SRC = '1100'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'BDC_OKCODE'
*                                              '/00'.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'BDC_OKCODE'
                    IV_FVAL = '/00'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).
*
*                PERFORM bdc_field       USING 'BDC_CURSOR'
*                                              'TTMS_FX_STRUCTURE_DATA-XTRADED_AMOUNT'.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'BDC_CURSOR'
                    IV_FVAL = 'TTMS_FX_STRUCTURE_DATA-XTRADED_AMOUNT'
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).


*                PERFORM bdc_field       USING 'TTMS_FX_STRUCTURE_DATA-XTRADED_AMOUNT'
*                                              is_ftr_cre-amount.

                BDC_FIELD(
                  EXPORTING
                    IV_FNAM = 'TTMS_FX_STRUCTURE_DATA-XTRADED_AMOUNT'
                    IV_FVAL = IS_FTR_CRE-SAMOUNT
                  CHANGING
                    CT_BDCDATA = LT_BDCDATA ).

                CALL TRANSACTION 'FTR_CREATE' USING LT_BDCDATA
                      MODE 'E'
                      UPDATE 'S'.

                "get rfha created in the transaction
                GET PARAMETER ID 'FAN' FIELD LV_RFHA.

                "Check if transaction was saved
                IF LV_RFHA IS NOT INITIAL.

                  GV_RFHA = LV_RFHA.

                  "Get exchange rate
                  SELECT KKURS
                  FROM VTBFHAZU
                  INTO TABLE @DATA(LT_VTBFHAZU)
                  UP TO 1 ROWS
                  WHERE BUKRS EQ @IS_FTR_CRE-BUKRS
                  AND RFHA EQ @LV_RFHA.

                  IF LINE_EXISTS( LT_VTBFHAZU[ 1 ] ).

                    DATA(LV_KKRUS) = LT_VTBFHAZU[ 1 ]-KKURS.

                    LOOP AT GT_AUTO ASSIGNING FIELD-SYMBOL(<FS_AUTO>) WHERE CHECK EQ GC_CHECK.

                      "update document exchange rate
                      CHANGE_EX_RATE(
                        EXPORTING
                          IV_EBELN = <FS_AUTO>-EBELN
                          IV_BUKRS = <FS_AUTO>-BUKRS
                          IV_KKURS =  LV_KKRUS ).

                      UPD_LOG(
                        EXPORTING
                          IS_AUTO = <FS_AUTO> ).

                    ENDLOOP.

                    MESSAGE S078(ZFI).

                    LEAVE TO SCREEN 0.

                  ENDIF.
                ELSE.

                  "If the transaction is not saved remove all checkbox checks
                  LS_AUTO-CHECK = GC_NCHECK.

                  MODIFY GT_AUTO FROM LS_AUTO TRANSPORTING CHECK WHERE CHECK EQ GC_CHECK.

                  AUTO_ALV->REFRESH( ).

                ENDIF.

              ELSE.

                "ocorreram erros ao obter colunas da ALV
                IF 1 = 2. MESSAGE E077(ZFI). ENDIF.
                RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
                  EXPORTING
                    MSGTY = SY-ABCDE+4(1)
                    MSGID = GC_MSGID
                    MSGNO = '077'.

              ENDIF.
            ELSEIF GT_MANUAL IS NOT INITIAL.

              IF LINE_EXISTS( GT_MANUAL[ CHECK = GC_CHECK ] ).

                "Get exchange rate
                SELECT KKURS
                FROM VTBFHAZU
                INTO TABLE @DATA(LT_VTBFHA)
                UP TO 1 ROWS
                WHERE BUKRS EQ @GV_BUKRS
                AND RFHA EQ @GV_RFHA.

                IF SY-SUBRC EQ 0.

                  DATA(LV_KKURS) = LT_VTBFHA[ 1 ]-KKURS.

                  LOOP AT GT_MANUAL ASSIGNING FIELD-SYMBOL(<FS_MANUAL>) WHERE CHECK EQ GC_CHECK.

                    "update document exchange rate
                    CHANGE_EX_RATE(
                      EXPORTING
                        IV_EBELN = <FS_MANUAL>-EBELN
                        IV_BUKRS = <FS_MANUAL>-BUKRS
                        IV_KKURS = LV_KKURS ).

                    UPD_LOG(
                   EXPORTING
                     IS_MANUAL = <FS_MANUAL> ).

                  ENDLOOP.

                  MESSAGE S078(ZFI).

                  LEAVE TO SCREEN 0.

                ENDIF.

                BACK.

              ELSE.

                "ocorreram erros ao obter colunas da ALV
                IF 1 = 2. MESSAGE E077(ZFI). ENDIF.
                RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
                  EXPORTING
                    MSGTY = SY-ABCDE+4(1)
                    MSGID = GC_MSGID
                    MSGNO = '077'.

              ENDIF.
            ENDIF.
        ENDCASE.

      CATCH ZCX_BC_EXCEPTIONS INTO DATA(LO_EXCPT).
        "mostrar mensagem de erro
        MESSAGE ID LO_EXCPT->MSGID
              TYPE SY-ABCDE+18(1)
            NUMBER LO_EXCPT->MSGNO
              WITH LO_EXCPT->MSGV1
                   LO_EXCPT->MSGV2
                   LO_EXCPT->MSGV3
                   LO_EXCPT->MSGV4
      DISPLAY LIKE LO_EXCPT->MSGTY.
    ENDTRY.

  ENDMETHOD.

  METHOD ON_CLICK.

    TRY.

        IF GT_AUTO IS NOT INITIAL.

          "Get clicked row
          DATA(LS_AUTO_CLICK) = VALUE #( GT_AUTO[ ROW ] ).

          "If checkbox is empty then fill checkbox
          IF LS_AUTO_CLICK-CHECK EQ GC_NCHECK.

            LS_AUTO_CLICK-CHECK = GC_CHECK.

          ELSE.

            "If checkbox is filled, empty checkbox
            LS_AUTO_CLICK-CHECK = GC_NCHECK.

          ENDIF.

          MODIFY GT_AUTO INDEX ROW FROM LS_AUTO_CLICK TRANSPORTING CHECK.

          AUTO_ALV->REFRESH( ).

        ELSEIF GT_MANUAL IS NOT INITIAL.

          "Get clicked row
          DATA(LS_MAN_CLICK) = VALUE #( GT_MANUAL[ ROW ] ).

          "If checkbox is empty then fill checkbox
          IF LS_MAN_CLICK-CHECK EQ GC_NCHECK.

            LS_MAN_CLICK-CHECK = GC_CHECK.

          ELSE.

            "If checkbox is filled, empty checkbox
            LS_MAN_CLICK-CHECK = GC_NCHECK.

          ENDIF.

          MODIFY GT_MANUAL INDEX ROW FROM LS_MAN_CLICK TRANSPORTING CHECK.

          MANUAL_ALV->REFRESH( ).

        ENDIF.

      CATCH ZCX_BC_EXCEPTIONS INTO DATA(LO_EXCPT).
        "mostrar mensagem de erro
        MESSAGE ID LO_EXCPT->MSGID
              TYPE SY-ABCDE+18(1)
            NUMBER LO_EXCPT->MSGNO
              WITH LO_EXCPT->MSGV1
                   LO_EXCPT->MSGV2
                   LO_EXCPT->MSGV3
                   LO_EXCPT->MSGV4
      DISPLAY LIKE LO_EXCPT->MSGTY.
    ENDTRY.
  ENDMETHOD.

  METHOD CHANGE_EX_RATE.

    DATA: LS_POHEADER TYPE BAPIMEPOHEADER,
          LS_POHEADX  TYPE BAPIMEPOHEADERX,
          LT_RET      TYPE TABLE OF BAPIRET2,
          LS_RETURN   TYPE BAPIRET2,
          LV_ERR      TYPE SAP_BOOL VALUE ABAP_FALSE.

    LS_POHEADER-COMP_CODE = IV_BUKRS.
    LS_POHEADER-EXCH_RATE = IV_KKURS * -1.
    LS_POHEADX-EXCH_RATE = 'X'.

    TRY.

        DATA(LO_OBJ) = NEW ZCLCA_BAL_LOG( IV_EXTNUMB = |{ IV_BUKRS }{ IV_EBELN }|
                                          IV_OBJETDT = GC_OBJCT
                                          IV_SUBOBJT = GC_SBOBJ ).


      CATCH ZCX_BC_EXCEPTIONS.
        "limpar objecto de log
        FREE LO_OBJ.

    ENDTRY.

    CALL FUNCTION 'BAPI_PO_CHANGE'
      EXPORTING
        PURCHASEORDER = IV_EBELN
        POHEADER      = LS_POHEADER
        POHEADERX     = LS_POHEADX
      TABLES
        RETURN        = LT_RET.

    IF LT_RET IS NOT INITIAL.

      LOOP AT LT_RET ASSIGNING FIELD-SYMBOL(<FS_RET>).

        IF <FS_RET>-TYPE EQ 'E'.

          LV_ERR = ABAP_TRUE.

        ENDIF.

        SET_LOG_MESSAGE( EXPORTING
                     IV_LEVEL = '1'
                     IV_MSGTY = <FS_RET>-TYPE
                     IV_MSGID = <FS_RET>-ID
                     IV_MSGNO = <FS_RET>-NUMBER
                     IV_MSGV1 = <FS_RET>-MESSAGE_V1
                     IV_MSGV2 = <FS_RET>-MESSAGE_V2
                     IV_MSGV3 = <FS_RET>-MESSAGE_V3
                     IV_MSGV4 = <FS_RET>-MESSAGE_V4
                     IV_SAVES = ABAP_TRUE
                   CHANGING
                     CO_EXLOG = LO_OBJ ).

      ENDLOOP.
    ENDIF.

    IF LV_ERR EQ ABAP_FALSE.

      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          WAIT   = GC_CHECK
        IMPORTING
          RETURN = LS_RETURN.

      IF LS_RETURN IS NOT INITIAL.

        SET_LOG_MESSAGE( EXPORTING
             IV_LEVEL = '1'
             IV_MSGTY = LS_RETURN-TYPE
             IV_MSGID = LS_RETURN-ID
             IV_MSGNO = LS_RETURN-NUMBER
             IV_MSGV1 = LS_RETURN-MESSAGE_V1
             IV_MSGV2 = LS_RETURN-MESSAGE_V2
             IV_MSGV3 = LS_RETURN-MESSAGE_V3
             IV_MSGV4 = LS_RETURN-MESSAGE_V4
             IV_SAVES = ABAP_TRUE
           CHANGING
             CO_EXLOG = LO_OBJ ).

        IF LS_RETURN-TYPE EQ 'E'.

          "ocorreram erros ao obter colunas da ALV
          IF 1 = 2. MESSAGE E087(ZFI). ENDIF.
          RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
            EXPORTING
              MSGTY = SY-ABCDE+4(1)
              MSGID = GC_MSGID
              MSGNO = '087'.

        ENDIF.

      ENDIF.

    ELSE.

      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

      "ocorreram erros ao obter colunas da ALV
      IF 1 = 2. MESSAGE E088(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '088'.

    ENDIF.

  ENDMETHOD.

  METHOD UPD_LOG.

    DATA: LS_LOG    TYPE ZFI_DOC_EX_LOG_T,
          LS_FOR    TYPE BAPI2042_MAINTAIN_FX,
          LV_EBELN  TYPE EBELN,
          LV_LIFNR  TYPE LIFNR,
          LV_KONTRH TYPE TB_KUNNR_NEW,
          LV_PDATE  TYPE DATUM,
          LV_NETWR  TYPE NETWR,
          LV_WAERS  TYPE WAERS,
          LS_RETURN TYPE BAPIRET2.

    TRY.

        DATA(LO_OBJ) = NEW ZCLCA_BAL_LOG( IV_EXTNUMB = |{ GV_BUKRS }{ GV_RFHA }|
                                          IV_OBJETDT = GC_OBJCT
                                          IV_SUBOBJT = GC_SBOBJ ).

      CATCH ZCX_BC_EXCEPTIONS.
        "limpar objecto de log
        FREE LO_OBJ.

    ENDTRY.

    CALL FUNCTION 'BAPI_FTR_GETDETAIL'
      EXPORTING
        COMPANYCODE     = GV_BUKRS
        TRANSACTION     = GV_RFHA
      IMPORTING
        FOREIGNEXCHANGE = LS_FOR
        RETURN          = LS_RETURN.

    IF LS_RETURN IS NOT INITIAL.

      SET_LOG_MESSAGE( EXPORTING
           IV_LEVEL = '1'
           IV_MSGTY = LS_RETURN-TYPE
           IV_MSGID = LS_RETURN-ID
           IV_MSGNO = LS_RETURN-NUMBER
           IV_MSGV1 = LS_RETURN-MESSAGE_V1
           IV_MSGV2 = LS_RETURN-MESSAGE_V2
           IV_MSGV3 = LS_RETURN-MESSAGE_V3
           IV_MSGV4 = LS_RETURN-MESSAGE_V4
           IV_SAVES = ABAP_TRUE
         CHANGING
           CO_EXLOG = LO_OBJ ).

      IF LS_RETURN-TYPE EQ 'E'.

        "ocorreram erros ao obter colunas da ALV
        IF 1 = 2. MESSAGE E079(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '079'.
      ENDIF.
    ENDIF.

    LS_LOG-EXTERNAL_REFERENCE = LS_FOR-EXTERNAL_REFERENCE.
    LS_LOG-FORWARD_RATE = LS_FOR-FORWARD_RATE.
    LS_LOG-AMOUNT1 = LS_FOR-AMOUNT1.

    IF IS_AUTO IS NOT INITIAL.

      LV_EBELN = IS_AUTO-EBELN.
      LV_LIFNR = IS_AUTO-LIFNR.
      LV_KONTRH = LS_FOR-PARTNER.
      LV_NETWR = IS_AUTO-NETWR.
      LV_WAERS = IS_AUTO-WAERS.

      IF LINE_EXISTS( GT_PDATE[ EBELN = LV_EBELN ] ).

        DATA(LS_PDATE) = GT_PDATE[ EBELN = LV_EBELN ].

        LV_PDATE = LS_PDATE-PDATE.

      ELSE.

        IF 1 = 2. MESSAGE E075(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '075'.

      ENDIF.

    ELSEIF IS_MANUAL IS NOT INITIAL.

      LV_EBELN = IS_MANUAL-EBELN.
      LV_WAERS = IS_MANUAL-WAERS.

      SELECT A~LIFNR,
             B~EINDT,
             D~NDAYS,
             F~ZTAG1
      FROM EKKO AS A
      INNER JOIN EKET AS B ON B~EBELN EQ A~EBELN
      INNER JOIN ZFI_PAY_DATE_T AS D ON D~ZZCOORI EQ A~ZZCOORI AND D~ZZEXPVZ EQ A~ZZEXPVZ
      INNER JOIN LFB1 AS E ON E~LIFNR EQ A~LIFNR AND E~BUKRS EQ A~BUKRS
      INNER JOIN T052 AS F ON F~ZTERM EQ E~ZTERM
      INTO TABLE @DATA(LT_EKKO)
      UP TO 1 ROWS
      WHERE A~EBELN EQ @LV_EBELN.

      IF SY-SUBRC IS NOT INITIAL.

        IF 1 = 2. MESSAGE E083(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '083'.

      ELSEIF SY-SUBRC EQ 0.

        LV_LIFNR = LT_EKKO[ 1 ]-LIFNR.
        LV_PDATE = LT_EKKO[ 1 ]-EINDT - LT_EKKO[ 1 ]-NDAYS - LT_EKKO[ 1 ]-ZTAG1.

      ENDIF.

      SELECT SUM( NETWR ) AS NETWR
      FROM EKPO
      INTO TABLE @DATA(LT_EKPO)
      WHERE EBELN EQ @LV_EBELN.

      IF SY-SUBRC IS NOT INITIAL.

        IF 1 = 2. MESSAGE E089(ZFI). ENDIF.
        RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
          EXPORTING
            MSGTY = SY-ABCDE+4(1)
            MSGID = GC_MSGID
            MSGNO = '089'.

      ELSEIF SY-SUBRC EQ 0.

        LV_NETWR = LT_EKPO[ 1 ]-NETWR.

      ENDIF.


      LV_KONTRH = LS_FOR-PARTNER.

    ENDIF.

    "Get data to insert into log table
    SELECT LTX
    FROM TZPAT
    INTO TABLE @DATA(LT_TZPAT)
    UP TO 1 ROWS
    WHERE SPRAS EQ @SY-LANGU
    AND GSART EQ @GV_SGSART.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E081(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '081'.

    ELSEIF SY-SUBRC EQ 0.

      LS_LOG-LTX = LT_TZPAT[ 1 ]-LTX.

    ENDIF.

    SELECT NAME_ORG1
    FROM BUT000
    INTO TABLE @DATA(LT_BUT000)
    UP TO 1 ROWS
    WHERE PARTNER EQ @LV_KONTRH.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E082(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '082'.

    ELSEIF SY-SUBRC EQ 0.

      LS_LOG-NAME_ORG1 = LT_BUT000[ 1 ]-NAME_ORG1.

    ENDIF.

    SELECT BEDAT,
           WKURS
    FROM EKKO
    INTO TABLE @DATA(LT_EKK)
    UP TO 1 ROWS
    WHERE EBELN EQ @LV_EBELN.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E084(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '084'.

    ELSEIF SY-SUBRC EQ 0.

      LS_LOG-BEDAT = LT_EKK[ 1 ]-BEDAT.
      LS_LOG-WKURS = LT_EKK[ 1 ]-WKURS.

    ENDIF.

    SELECT NAME1
    FROM LFA1
    INTO TABLE @DATA(LT_LFA1)
    UP TO 1 ROWS
    WHERE LIFNR EQ @LV_LIFNR.

    IF SY-SUBRC IS NOT INITIAL.

      IF 1 = 2. MESSAGE E085(ZFI). ENDIF.
      RAISE EXCEPTION TYPE ZCX_BC_EXCEPTIONS
        EXPORTING
          MSGTY = SY-ABCDE+4(1)
          MSGID = GC_MSGID
          MSGNO = '085'.

    ELSEIF SY-SUBRC EQ 0.

      LS_LOG-NAME1 = LT_LFA1[ 1 ]-NAME1.

    ENDIF.

    LS_LOG-PAY_DATE = LV_PDATE.
    LS_LOG-ASSIGN_DATE = SY-DATUM.
    LS_LOG-BUKRS = GV_BUKRS.
    LS_LOG-RFHA = GV_RFHA.
    LS_LOG-EBELN = LV_EBELN.
    LS_LOG-SGSART = GV_SGSART.
    LS_LOG-KONTRH = LV_KONTRH.
    LS_LOG-NETWR = LV_NETWR.
    LS_LOG-LIFNR = LV_LIFNR.
    LS_LOG-WAERS = LV_WAERS.

    INSERT ZFI_DOC_EX_LOG_T FROM LS_LOG.

  ENDMETHOD.

  METHOD SET_LOG_MESSAGE.

    "verificar se o objecto de log está instânciado
    IF CO_EXLOG IS INITIAL.
      "sair do processamento
      RETURN.
    ENDIF.

    TRY.
        "preencher estrutura de log
        DATA(LS_BMSGS) = VALUE BAL_S_MSG( MSGTY    = IV_MSGTY
                                          MSGID    = IV_MSGID
                                          MSGNO    = IV_MSGNO
                                          MSGV1    = IV_MSGV1
                                          MSGV2    = IV_MSGV2
                                          MSGV3    = IV_MSGV3
                                          MSGV4    = IV_MSGV4
                                          DETLEVEL = IV_LEVEL ).

        "adicionar mensagem ao log
        CO_EXLOG->ADD_MESSAGE_LOG( EXPORTING
                                     IS_BAL_MSG = LS_BMSGS ).

        "verificar se o log deve ser guardado na base de dados
        IF IV_SAVES IS NOT INITIAL.
          "guardar log na base de dados
          CO_EXLOG->SAVE_LOG_MESSAG( ).
        ENDIF.

      CATCH ZCX_BC_EXCEPTIONS.
        "sair do processamento
        RETURN.
    ENDTRY.

  ENDMETHOD.

  METHOD BDC_DYNPRO.

    DATA IS_BDCDATA TYPE BDCDATA.

    IS_BDCDATA-PROGRAM = IV_PROG.
    IS_BDCDATA-DYNPRO  = IV_SRC.
    IS_BDCDATA-DYNBEGIN = 'X'.

    APPEND IS_BDCDATA TO CT_BDCDATA.

  ENDMETHOD.

  METHOD BDC_FIELD.

    DATA IS_BDCDATA TYPE BDCDATA.

    IS_BDCDATA-FNAM = IV_FNAM.
    IS_BDCDATA-FVAL = IV_FVAL.

    APPEND IS_BDCDATA TO CT_BDCDATA.

  ENDMETHOD.

  METHOD CREATE_RANGE.

    CLEAR ER_EBELN.

    DATA: LS_LEBELN TYPE LINE OF TY_REBELN.

    LOOP AT IT_EXLOG ASSIGNING FIELD-SYMBOL(<FS_EXLOG>).
      CLEAR LS_LEBELN.

      LS_LEBELN-SIGN = 'I'.
      LS_LEBELN-OPTION = 'EQ'.
      LS_LEBELN-LOW = <FS_EXLOG>-EBELN.

      APPEND LS_LEBELN TO ER_EBELN.

    ENDLOOP.

    SORT ER_EBELN ASCENDING.

    DELETE ADJACENT DUPLICATES FROM ER_EBELN COMPARING ALL FIELDS.

  ENDMETHOD.

ENDCLASS.