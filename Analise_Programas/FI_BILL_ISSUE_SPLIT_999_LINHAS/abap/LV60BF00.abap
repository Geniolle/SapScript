*---------------------------------------------------------------------
*
*  Perform          ACCOUNTING_HEAD_LINE.
*
*  Perform          ACCOUNTING_ITEM_LINE.
*
*  Perform          ACCOUNTING_TAX_LINE.
*
*  Perform          READ_FI_PERIODE_INVOICEDATE.
*
*  Perform          TXJCD_AUFBEREITEN.
*
*  Perform          FREE_OF_CHARGE_LINE.
*
*  Perform          XACCIT_XNEGP_SET.
*
*  Perform          XACCIT_XVALGS_ZFBDT_SET.
*
*  Perform          CURRENCY_CONVERSION
*
*  Perform          FILL_ACCIT_DEB
*
* Perform           ENHANCE_CONTR_REFERENCE
*
* Perform           map_extensibility_flow_bill
*---------------------------------------------------------------------

*---------------------------------------------------------------------
*
*       FORM ACCOUNTING_HEAD_LINE
*
*---------------------------------------------------------------------
*
*       fill accounting document customer line item
*
*---------------------------------------------------------------------
*
*  -->  VBRK           workarea invoice header
*
*  <--  XACCHD         table document header
*
*  <--  XACCIT,XACCCR  table customer line item
*
*---------------------------------------------------------------------
*
FORM ACCOUNTING_HEAD_LINE USING UV_ASSIGN_BASEDOC         TYPE ABAP_BOOL
                                UV_DETERMINE_BASELINEDATE TYPE ABAP_BOOL.

*
* fill customer line item
*
  POSNR = POSNR + 1.

  CLEAR XACCIT.
  CLEAR XACCCR.
  CLEAR LT_ACCFI.

  MOVE-CORRESPONDING VBRK TO XACCIT.
  CLEAR: XACCIT-LAND1, XACCIT-KDGRP.
  MOVE-CORRESPONDING XACCIT_DEB TO XACCIT.
ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_01 SPOTS ES_SAPLV60B.
  CLEAR XACCIT-GJAHR.
  IF VBRK-VKONT IS INITIAL.
    CLEAR XACCIT-GSBER.
*   Restore field SPART in case FI/CA is not used
    XACCIT-SPART = VBRK-SPART.
  ENDIF.
  XACCIT-AWTYP = CON_AWTYP_VBRK.
  XACCIT-AWREF = VBRK-VBELN.
  XACCIT-BELNR = VBRK-VBELN.
  XACCIT-ZUONR = VBRK-ZUONR.
  IF VBRK-VKONT IS NOT INITIAL.
    XACCIT-TXJCD = TXJCD.
  ENDIF.
  XACCIT-MWSKZ = XACCIT_DEB-MWSK1.
* Time dependent taxes: Determine the tax rate validity start date

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.

    IF VBRK-LANDTX IS NOT INITIAL. "See Developer Memo of CM 51251 (2020)

      IF XACCIT-MWSKZ IS NOT INITIAL AND XACCIT-TXDAT IS NOT INITIAL AND CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
        IV_BUKRS = XACCIT-BUKRS
        IV_LAND1 = COND #( WHEN CL_FOT_TXA_UTILITIES=>AGENT->IS_TAX_ABROAD_ACTIVE( T001-BUKRS ) EQ ABAP_TRUE
                           THEN VBRK-LANDTX ELSE T001-LAND1 )
      ).
        XACCIT-TXDAT_FROM = CL_WLF_TDT_SERVICE=>GET_TAX_CAL_DATE_FROM( I_BUKRS = XACCIT-BUKRS
                                                                       I_MWSKZ = XACCIT-MWSKZ
                                                                       I_FBUDA = XACCIT-TXDAT
                                                                       I_TAX_COUNTRY = COND #( WHEN CL_FOT_TXA_UTILITIES=>AGENT->IS_TAX_ABROAD_ACTIVE( T001-BUKRS ) EQ ABAP_TRUE
                                                                                               THEN VBRK-LANDTX ELSE T001-LAND1 )
                                                                     ).
      ENDIF.

    ENDIF.

  ELSE.
    "For invoice lists vbrk-landtx is not filled therefore we use t001-land1 to check tdt

    IF XACCIT-MWSKZ IS NOT INITIAL AND XACCIT-TXDAT IS NOT INITIAL AND CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
      IV_BUKRS = XACCIT-BUKRS
      IV_LAND1 = T001-LAND1
    ).
      XACCIT-TXDAT_FROM = CL_WLF_TDT_SERVICE=>GET_TAX_CAL_DATE_FROM( I_BUKRS = XACCIT-BUKRS
                                                                     I_MWSKZ = XACCIT-MWSKZ
                                                                     I_FBUDA = XACCIT-TXDAT
                                                                     I_TAX_COUNTRY = T001-LAND1
                                                                   ).
    ENDIF.
  ENDIF.



  IF VBRK-FKTYP = CON_FKTYP_P.
    XACCIT-ZUMSK = CON_ZUMSK_A.
    XACCIT-UMSKZ = CON_UMSKZ_F.
    XACCIT-BSTAT = CON_BSTAT_S.
  ENDIF.
  MOVE-CORRESPONDING XKOMK1 TO XACCIT.

  XACCCR-MANDT = VBRK-MANDT.
  XACCCR-AWTYP = CON_AWTYP_VBRK.
  XACCCR-AWREF = VBRK-VBELN.
  XACCCR-AWORG = SPACE.

* no payment terms, when payment service provider is used
  IF CL_OPS_SWITCH_CHECK=>SD_SFWS_SC4( ) EQ ABAP_TRUE.
    IF VBRK-SPPAYM EQ CV_SPPAYM01.
      CLEAR XACCIT-ZTERM.
      CLEAR XACCIT-ZBD1T.
      CLEAR XACCIT-ZBD2T.
      CLEAR XACCIT-ZBD3T.
      CLEAR XACCIT-ZBD1P.
      CLEAR XACCIT-ZBD2P.
    ENDIF.
  ENDIF.

* cancellation
  XACCIT-AWREF_REV = VBRK-SFAKN.
  IF VBRK-VBTYP EQ IF_SD_DOC_CATEGORY=>INVOICE_CANCEL.
    XACCIT-REBZG(1) = 'V'.
  ENDIF.

  IF XACCIT_DEB-CASH = SPACE.
* customer line item
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      XACCCR-WRBTR = POS_BRUTTO.
    ELSE.
      XACCCR-WRBTR = WARENWERT + TAX.
    ENDIF.
    XACCCR-SKFBT = CASHDISCOUNT.
    XACCIT-ABSBT = SECUREVALUE.
ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_02 SPOTS ES_SAPLV60B.
* convert secure value to credit currency
    IF VBRK-WAERK NE VBRK-CMWAE AND NOT VBRK-CMWAE IS INITIAL.
      PERFORM CURRENCY_CONVERSION.
    ENDIF.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCCR-WRBTR = XACCCR-WRBTR * -1.
      XACCCR-SKFBT = XACCCR-SKFBT * -1.
      POS_NETTO    = POS_NETTO * -1.
      TAX          = TAX * -1.
    ENDIF.
    CLEAR XACCIT-SHKZG.                " determined by FI

* Settlement2Invoice
    IF CL_ERP_EHP_SWITCH_CHECK=>ERP_CF_SFWS_1( ) EQ ABAP_TRUE.
      IF NOT VBRK-FK_SOURCE_SYS IS INITIAL.
        XACCIT-FKTYP = CON_F.
      ENDIF.
    ENDIF.

    IF XACCCR-WRBTR LT 0.
      IF VBRK-FKTYP = CON_FKTYP_P.
        XACCIT-BSCHL = CON_BSCHL_19.
      ELSE.
        XACCIT-BSCHL = CON_BSCHL_11.
      ENDIF.
      IF XACCCR-SKFBT GT 0.
        XACCCR-SKFBT = 0.
      ENDIF.
    ELSE.
      IF VBRK-FKTYP = CON_FKTYP_P.
        XACCIT-BSCHL = CON_BSCHL_09.
      ELSE.
        XACCIT-BSCHL = CON_BSCHL_01.
      ENDIF.
      IF XACCCR-SKFBT LT 0.
        XACCCR-SKFBT = 0.
      ENDIF.
    ENDIF.
    XACCIT-BELNR = VBRK-VBELN.         "ext. numbering note 8583
    CASE TVFK-XFILKD.
      WHEN ' '.
        IF VBRK-KUNRG = SPACE OR VBRK-KUNRG = VBRK-KUNAG.
          XACCIT-KUNNR = VBRK-KUNAG.
          XACCIT-FILKD = SPACE.
        ELSE.
          XACCIT-KUNNR = VBRK-KUNRG.
          XACCIT-FILKD = VBRK-KUNAG.
        ENDIF.
      WHEN 'A'.
        IF NOT VBRK-KNKLI IS INITIAL AND
           VBRK-KNKLI NE VBRK-KUNAG.
          XACCIT-KUNNR = VBRK-KUNRG.
          XACCIT-FILKD = VBRK-KUNAG.
        ELSE.
          XACCIT-KUNNR = VBRK-KUNAG.
          XACCIT-FILKD = SPACE.
          XACCIT-XFILKD = 'X'.
        ENDIF.
      WHEN 'B'.
        XACCIT-KUNNR = VBRK-KUNRG.
        XACCIT-FILKD = SPACE.
        XACCIT-XFILKD = 'X'.
    ENDCASE.
ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_07 SPOTS ES_SAPLV60B.
*
    XACCIT-ZFBDT = VBRK-FKDAT + VBRK-VALTG.

    IF VBRK-VALDT NE 0.
      XACCIT-ZFBDT = VBRK-VALDT.
    ENDIF.
    IF NOT VBRK-FKART_RL IS INITIAL AND
       NOT VBRK-FKDAT_RL IS INITIAL AND
       CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_RETOUR( VBRK-VBTYP ).
      XACCIT-ZLSPR = 'A'.
    ENDIF.
    SAVE_ZFBDT = XACCIT-ZFBDT.

    CALL FUNCTION 'FI_TERMS_OF_PAYMENT_PROPOSE'
      EXPORTING
        I_BLDAT       = XACCIT-ZFBDT
        I_BUDAT       = XACCIT-ZFBDT
        I_CPUDT       = XACCIT-ZFBDT
        I_ZFBDT       = XACCIT-ZFBDT
        I_ZTERM       = VBRK-ZTERM
      IMPORTING
        E_ZBD1T       = XACCIT-ZBD1T
        E_ZBD1P       = XACCIT-ZBD1P
        E_ZBD2T       = XACCIT-ZBD2T
        E_ZBD2P       = XACCIT-ZBD2P
        E_ZBD3T       = XACCIT-ZBD3T
        E_ZFBDT       = XACCIT-ZFBDT
        E_ZLSCH       = XACCIT-ZLSCH
        E_T052        = X_T052
      EXCEPTIONS
        ERROR_MESSAGE = 4
        OTHERS        = 4.

    IF NOT VBRK-ZLSCH IS INITIAL.
      XACCIT-ZLSCH = VBRK-ZLSCH.
    ENDIF.

ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_03 SPOTS ES_SAPLV60B.
*   Downpayment request only with due date
    IF VBRK-FKTYP EQ CON_FKTYP_P.
*     Clearing the FI-terms of payments in the RW-Interface
      PERFORM CLEAR_TERMS_OF_PAYMENT.
    ENDIF.

* credit memo with value date: use baseline date for payment of the
* respective invoice
    IF LOC_BSID-ZFBDT IS NOT INITIAL AND
       LOC_BSID-ZFBDT GE XACCIT-ZFBDT AND
       UV_DETERMINE_BASELINEDATE EQ ABAP_TRUE.
      XACCIT-ZFBDT = LOC_BSID-ZFBDT.
    ENDIF.
* credit memo with value date: transfer reference to the accounting
* document of the respective invocie
    IF UV_ASSIGN_BASEDOC EQ ABAP_TRUE AND
       LOC_BKPF-BELNR IS NOT INITIAL.
      XACCIT-REBZG = LOC_BKPF-BELNR.
      XACCIT-REBZJ = LOC_BSID-GJAHR.
      XACCIT-REBZZ = LOC_BSID-BUZEI.
    ENDIF.

    XACCIT-ADRNR = CPD_ADRESS.

* CPD TAX NUMBERS.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      SORT XVBPA BY MANDT VBELN POSNR PARVW.
      READ TABLE XVBPA WITH KEY VBELN = XVBRP_KEY-VBELN
                                POSNR = 0
                                PARVW = 'RG'
                                BINARY SEARCH.           "#EC CI_SORTED

      IF SY-SUBRC = 0 AND XVBPA-XCPDK = 'X'.
        LT_ACCFI-MANDT = VBRK-MANDT.
        LT_ACCFI-FITYP = VBRK-J_1AFITP.
        LT_ACCFI-AWTYP = CON_AWTYP_VBRK.
        LT_ACCFI-AWREF = VBRK-VBELN.
        LT_ACCFI-AWORG = SPACE.
        MOVE-CORRESPONDING XVBPA TO LT_ACCFI.

        DATA: LS_VBADR LIKE VBADR.
        DATA: LS_VBPA LIKE VBPA.
        DATA: LS_SDADR LIKE SDPARTNER_ADDRESS.
        DATA: LS_ADNUMBER TYPE ADDR1_SEL-ADDRNUMBER.
        DATA: LS_ADHANDLE TYPE SZAD_FIELD-HANDLE.
        DATA: LS_ADTYP TYPE AD_ADRTYPE VALUE '1'.

        IF XVBPA-ADRNR CA '$'.
          LS_ADHANDLE = XVBPA-ADRNR.
        ELSE.
          LS_ADNUMBER = XVBPA-ADRNR.
        ENDIF.

        CALL FUNCTION 'SD_ADDRESS_GET'
          EXPORTING
            FIF_ADDRESS_NUMBER      = LS_ADNUMBER
            FIF_ADDRESS_HANDLE      = LS_ADHANDLE
            FIF_ADDRESS_TYPE        = LS_ADTYP
            FIF_LANGU               = XVBPA-SPRAS
          IMPORTING
            FES_ADDRESS             = LS_VBADR
            FES_SDPARTNER_ADDRESS   = LS_SDADR
          EXCEPTIONS
            ADDRESS_NOT_FOUND       = 1
            ADDRESS_TYPE_NOT_EXISTS = 2
            NO_PERSON_NUMBER        = 3
            ERROR_MESSAGE           = 4
            OTHERS                  = 5.

        IF SY-SUBRC = 0.
          MOVE-CORRESPONDING LS_VBADR TO LT_ACCFI.
          IF NOT LS_SDADR-PO_BOX_NUM IS INITIAL.
            LT_ACCFI-PO_BOX_NUM = LS_SDADR-PO_BOX_NUM.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.

* Argentina/ Brazil: invoice reference in invoice related credit memo
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.       " Not for invoice list
      DATA: J1B_ACTIVE_COMPONENT(2),
            J1B_ACTIVE.                " Brazil/Argentina ind.
      CALL FUNCTION 'J_1BSA_COMPONENT_ACTIVE'
        EXPORTING
          BUKRS                = VBRK-BUKRS
          COMPONENT            = '**'
        IMPORTING
          ACTIVE_COMPONENT     = J1B_ACTIVE_COMPONENT
        EXCEPTIONS
          COMPONENT_NOT_ACTIVE = 1.
      IF    J1B_ACTIVE_COMPONENT = 'AR'
         OR J1B_ACTIVE_COMPONENT = 'BR'.
        J1B_ACTIVE = 'X'.
      ELSE.
        J1B_ACTIVE = ' '.
      ENDIF.
      IF J1B_ACTIVE = 'X' AND VBRK-SFAKN IS INITIAL."<<ins. note 0522657
        CALL FUNCTION 'J_1B_SD_FI_INTERFACE'
          EXPORTING
            I_VBRK       = VBRK
            I_XACCIT     = XACCIT
            I_DOC_OLD    = DOCUMENT_OLD
            I_XACCIT_DEB = XACCIT_DEB
          IMPORTING
            E_XACCIT     = XACCIT
          TABLES
            T_VBRP       = XVBRP
            T_KOMV       = XKOMV.
      ENDIF.

      IF J1B_ACTIVE_COMPONENT = 'AR'.  "<<<< INSERT - NOTE 126562
*** Start of note 2020052
        DATA: LV_STRLEN(2) TYPE N.
        LV_STRLEN = STRLEN( VBRK-XBLNR ).
        IF LV_STRLEN = 14.
          XACCIT-BRNCH = VBRK-XBLNR+1(4).
        ELSE.
*** End of note 2020052
          XACCIT-BRNCH = VBRK-XBLNR(4).  "<<<< INSERT - NOTE 126562
        ENDIF.                             " Note 2020052
        CLEAR: LV_STRLEN.                      " Note 2020052
      ENDIF.                           "<<<< INSERT - NOTE 126562

    ENDIF.

ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_06 SPOTS ES_SAPLV60B.
* BADI call for customer line item
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
        CALL BADI GR_SD_ACCOUNTING_BADI->ACCOUNTING_HEAD_LINE
          EXPORTING
            FVBRK       = VBRK
            FDOC_NUMBER = XVBRP_KEY-VBELN
            FVBRP       = XVBRP
            FKOMV       = XKOMV
            F_TVFK      = TVFK
            FACCIT_DEB  = XACCIT_DEB
          CHANGING
            FXACCIT     = XACCIT
            FXACCCR     = XACCCR
            FXACCHD     = XACCHD[].
      ENDIF.
    ENDIF.

* old userexits are executed due to upward compatibility
    MOVE-CORRESPONDING XACCIT TO XKOMK2.
    PERFORM USEREXIT_FILL_XKOMK2.
    MOVE-CORRESPONDING XKOMK2 TO XACCIT.

* userexit customer line item
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      CALL CUSTOMER-FUNCTION '002'
        EXPORTING
          XACCIT     = XACCIT
          VBRK       = VBRK
          DOC_NUMBER = XVBRP_KEY-VBELN
        IMPORTING
          XACCIT     = XACCIT
        TABLES
          CVBRP      = XVBRP
          CKOMV      = XKOMV.
    ELSE.
      CALL CUSTOMER-FUNCTION '002'
        EXPORTING
          XACCIT = XACCIT
          VBRK   = VBRK
        IMPORTING
          XACCIT = XACCIT
        TABLES
          CKOMV  = XKOMV.
    ENDIF.
* Zero value customer lines for FI-CA
    IF NOT XACCCR-WRBTR IS INITIAL
    OR ( NOT VBRK-VKONT IS INITIAL
    AND VBRK-FKTYP NE CON_FKTYP_P
    AND XACCIT_DEB-MWSK1 NE SPACE ) .
* document currency
      XACCCR-CURTP = '00'.
      XACCCR-WAERS = VBRK-WAERK.
      XACCCR-KURSF = VBRK-KURRF.

* several customer line items for installment plan
      DATA: DIFFERENZ LIKE VBRK-NETWR.
      DATA: LAST_SKFBT LIKE VBRK-NETWR.
*     secure value
      DATA: DIFFERENZ_SV LIKE VBRK-NETWR.
      DATA: LAST_ABSBT LIKE VBRK-NETWR.
*     tax for downpayment requests
      DATA: DIFFERENZ_MWST LIKE VBRK-NETWR.

      CASHDISCOUNT = XACCCR-SKFBT.
      SECUREVALUE  = XACCIT-ABSBT.
      CALL FUNCTION 'BILLING_SCHEDULE_CREATE_T052S'
        EXPORTING
          ZTERM               = VBRK-ZTERM
          WERT                = XACCCR-WRBTR  "Warenwert + Tax
          WAERK               = VBRK-WAERK
          FKDAT               = VBRK-FKDAT
          VALTG               = VBRK-VALTG
          VALDT               = VBRK-VALDT
          I_COMPANY_CODE      = VBRK-BUKRS
        TABLES
          ZFPLT               = RFPLT
        EXCEPTIONS
          NO_ENTRY_T052S      = 01
          NO_ZFBDT            = 02
          NO_ENTRY_T052       = 03
          NO_BILLING_SCHEDULE = 04.

      DATA: DA_SUBRC LIKE SY-SUBRC.
      DA_SUBRC = SY-SUBRC.
* installment plan: restore old baseline date
      IF DA_SUBRC IS INITIAL AND NOT RFPLT[] IS INITIAL.
        XACCIT-ZFBDT = SAVE_ZFBDT.
      ENDIF.

      CALL CUSTOMER-FUNCTION '007'
        EXPORTING
          VBRK   = VBRK
          XACCIT = XACCIT
        IMPORTING
          XACCIT = XACCIT
        TABLES
          XFPLT  = RFPLT.

* save the perhaps changed ZFBDT because it is the calculation basis
* for the installment plan dates
      SAVE_ZFBDT = XACCIT-ZFBDT.

      CASE DA_SUBRC.
        WHEN 0.
          POSNR = POSNR - 1.
          IF J1B_ACTIVE = 'X'          " reset line number
          AND NOT XACCIT-REBZG IS INITIAL.        " related document

*           Argentina/ Brazil: invoice reference
            XACCIT-REBZZ = XACCIT-REBZZ - 1.      "instalments
          ENDIF.
          LOOP AT RFPLT.
* withholding tax key
            IF XACCIT_DEB-WT_KEY NE 0.
              LT_KEY = LT_KEY + 1.
              IF WOLT_KEY NE SPACE.
                LT_KEY = WOLT_KEY + 1.
                CLEAR WOLT_KEY.
              ENDIF.
            ELSE.
              IF SY-TABIX EQ 1.
                OLDLT_KEY = LT_KEY.
                LT_KEY = XACCIT_DEB-WT_KEY.
              ELSE.
                WOLT_KEY = OLDLT_KEY.
              ENDIF.
            ENDIF.
            POSNR = POSNR + 1.
            XACCIT-POSNR = POSNR.
            IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
              IF XVBPA-XCPDK = 'X'.
                LT_ACCFI-POSNR = POSNR.    " CPD TAX.
                APPEND LT_ACCFI.           " CPD TAX
              ENDIF.
            ENDIF.

            IF J1B_ACTIVE = 'X'        " Fill line number
            AND NOT XACCIT-REBZG IS INITIAL.
*           Argentina/ Brazil: invocie reference
              XACCIT-REBZZ = XACCIT-REBZZ + 1.

              TABLES BSEG.

              SELECT SINGLE * FROM BSEG
                            WHERE BUKRS = XACCIT-BUKRS
                              AND BELNR = XACCIT-REBZG
                              AND GJAHR = XACCIT-REBZJ
                              AND BUZEI = XACCIT-REBZZ.

              IF BSEG-KOART <> 'D'.
                XACCIT-REBZZ = XACCIT-REBZZ - 1.
              ENDIF.


            ENDIF.
            XACCIT-ZTERM = RFPLT-ZTERM.
* restore the previously saved ZFBDT, so that for all rates the
* due dates are calculated from the same basis
            XACCIT-ZFBDT = SAVE_ZFBDT.
            DATA : LD_ZFBDT LIKE ACCIT-ZFBDT.
            CALL FUNCTION 'FI_TERMS_OF_PAYMENT_PROPOSE'
              EXPORTING
                I_BLDAT       = XACCIT-ZFBDT
                I_BUDAT       = XACCIT-ZFBDT
                I_CPUDT       = XACCIT-ZFBDT
                I_ZFBDT       = XACCIT-ZFBDT
                I_ZTERM       = RFPLT-ZTERM
              IMPORTING
                E_ZBD1T       = XACCIT-ZBD1T
                E_ZBD1P       = XACCIT-ZBD1P
                E_ZBD2T       = XACCIT-ZBD2T
                E_ZBD2P       = XACCIT-ZBD2P
                E_ZBD3T       = XACCIT-ZBD3T
                E_ZFBDT       = XACCIT-ZFBDT
                E_ZLSCH       = XACCIT-ZLSCH
                E_T052        = X_T052
              EXCEPTIONS
                ERROR_MESSAGE = 4
                OTHERS        = 4.

            IF NOT VBRK-ZLSCH IS INITIAL.
              XACCIT-ZLSCH = VBRK-ZLSCH.
            ENDIF.

            XACCIT-ABSBT = RFPLT-FPROZ * SECUREVALUE  / 100.

            XACCCR-POSNR = POSNR.
            XACCCR-WRBTR = RFPLT-FAKWR.
            XACCCR-SKFBT = RFPLT-FPROZ * CASHDISCOUNT / 100.
*           downpayment requests with installment plan: tax value is
*           filled in the respective customer line item
            IF VBRK-FKTYP EQ CON_FKTYP_P OR NOT VBRK-VKONT IS INITIAL.
              IF POS_NETTO IS NOT INITIAL.
                XACCCR-WMWST =
                 ( XACCCR-WRBTR * ( TAX * 100 / POS_NETTO ) ) /
                       ( ( TAX * 100 / POS_NETTO ) + 100 ).
              ELSE.
                XACCCR-WMWST =
                 ( XACCCR-WRBTR * ( TAX * 100 / POS_BRUTTO ) ) /
                       ( ( TAX * 100 / POS_BRUTTO ) + 100 ).
              ENDIF.
*             Downpayment request only with due date
              IF VBRK-FKTYP EQ CON_FKTYP_P.
*               Clearing the FI-terms of payments in the RW-Interface
                PERFORM CLEAR_TERMS_OF_PAYMENT.
              ENDIF.
            ENDIF.
            AT LAST.
*          last table entry: compensate rounding difference for cash
*          discount base value
              LAST_SKFBT = CASHDISCOUNT - DIFFERENZ.
              XACCCR-SKFBT = LAST_SKFBT.
              LAST_ABSBT = SECUREVALUE  - DIFFERENZ_SV.
              XACCIT-ABSBT = LAST_ABSBT.
*             compensate difference for downpayment requests
              IF VBRK-FKTYP EQ CON_FKTYP_P OR NOT VBRK-VKONT IS INITIAL.
                XACCCR-WMWST = TAX - DIFFERENZ_MWST.
              ENDIF.
            ENDAT.
            XACCIT-WT_KEY = LT_KEY.
            APPEND XACCCR.
            APPEND XACCIT.

            ADD XACCCR-SKFBT TO DIFFERENZ.
            ADD XACCIT-ABSBT TO DIFFERENZ_SV.
*           determine difference for downpayment requests
            IF VBRK-FKTYP EQ CON_FKTYP_P OR NOT VBRK-VKONT IS INITIAL.
              ADD XACCCR-WMWST TO DIFFERENZ_MWST.
            ENDIF.
          ENDLOOP.

        WHEN OTHERS.

          XACCIT-POSNR = POSNR.
          APPEND XACCIT.

          XACCCR-POSNR = POSNR.
          APPEND XACCCR.
          IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
            IF XVBPA-XCPDK = 'X'.
              LT_ACCFI-POSNR = POSNR.      "CPD TAX
              APPEND LT_ACCFI.
            ENDIF.
          ENDIF.
      ENDCASE.
    ELSE.
      IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
        IF XVBPA-XCPDK = 'X'.
          LT_ACCFI-POSNR = POSNR - 1.      "CPD TAX
          APPEND LT_ACCFI.
        ENDIF.
      ENDIF.
    ENDIF.

  ELSE.

* G/L account item instead of customer line item ( cash sale )
    XACCIT-VALUT = SY-DATLO.
    XACCIT-GSBER = XGSBER.
    XACCCR-WRBTR = POS_BRUTTO.
    XACCCR-SKFBT = CASHDISCOUNT.

ENHANCEMENT-POINT ACCOUNTING_HEAD_LINE_05 SPOTS ES_SAPLV60B.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCCR-WRBTR = XACCCR-WRBTR * -1.
      XACCCR-SKFBT = XACCCR-SKFBT * -1.
    ENDIF.
    CLEAR XACCIT-SHKZG.
    CLEAR XACCIT-ZTERM.
    IF XACCCR-WRBTR LT 0.
      XACCIT-BSCHL = '50'.
      IF XACCCR-SKFBT GT 0.
        XACCCR-SKFBT = 0.
      ENDIF.
    ELSE.
      XACCIT-BSCHL = '40'.
      IF XACCCR-SKFBT LT 0.
        XACCCR-SKFBT = 0.
      ENDIF.
    ENDIF.

* old userexits are executed due to upward compatibility
    MOVE-CORRESPONDING XACCIT TO XKOMK3.
    PERFORM USEREXIT_FILL_XKOMK3_CASH.
    MOVE-CORRESPONDING XKOMK3 TO XACCIT.

* userexit cash sale
    CALL CUSTOMER-FUNCTION '003'
      EXPORTING
        XACCIT     = XACCIT
        VBRK       = VBRK
        DOC_NUMBER = XVBRP_KEY-VBELN
      IMPORTING
        XACCIT     = XACCIT
      TABLES
        CVBRP      = XVBRP
        CKOMV      = XKOMV.

    XACCIT-POSNR = POSNR.
    LT_ACCFI-POSNR = POSNR.            "CPD TAX

    IF NOT XACCCR-WRBTR IS INITIAL.
      APPEND XACCIT.
      IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
        IF XVBPA-XCPDK = 'X'.
          APPEND LT_ACCFI.                 "CPD TAX
        ENDIF.
      ENDIF.
* document currency
      XACCCR-CURTP = '00'.
      XACCCR-WAERS = VBRK-WAERK.
      XACCCR-KURSF = VBRK-KURRF.

      XACCCR-POSNR = POSNR.
      APPEND XACCCR.

    ENDIF.
  ENDIF.

ENDFORM.                    "accounting_head_line

*---------------------------------------------------------------------
*
*       FORM ACCOUNTING_ITEM_LINE
*
*---------------------------------------------------------------------
*
*       fill accounting document G/L account item
*
*---------------------------------------------------------------------
*
*  -->  XVPRP           table invoice items
*
*  -->  XKOMV           table conditions
*
*  <--  XACCIT, XACCCR  tables G/L account item
*
*---------------------------------------------------------------------
*
FORM ACCOUNTING_ITEM_LINE.

* internal structures
  DATA: OLDXACCCR    LIKE XACCCR,
        LVS_KONVFLAG LIKE KONVFLAG.

* Differential billing.
  DATA: LV_IS_DIFF_CAPABLE TYPE XFELD.

* Flag XACCIT-KRUEK set by differential billing.
  DATA: LV_KRUEK_FROM_BILL_DIFF       TYPE XFELD.

* Business function switches for differential billing
* and period-end valuation. Statics for performance.
* Initialised with 'I' different from the valid values ' ' and 'X'.
  STATICS: SV_BILL_DIFF_ACTIVE        TYPE CHAR1 VALUE 'I'.
  STATICS: SV_PERIOD_END_ACTIVE       TYPE CHAR1 VALUE 'I'.
* Self Billing
  DATA: LS_TVAU TYPE TVAU.

ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_01 SPOTS ES_SAPLV60B.

* Determine business function switch once only.
  IF SV_BILL_DIFF_ACTIVE = 'I'.
    SV_BILL_DIFF_ACTIVE = CL_SD_BILL_SWITCH_CHECK=>SD_SFWS_BILL_DIFF_1( ).
  ENDIF.
  IF SV_PERIOD_END_ACTIVE = 'I'.
    SV_PERIOD_END_ACTIVE = CL_SD_BILL_SWITCH_CHECK=>SD_SFWS_MEV_01( ).
  ENDIF.

  POSNR = POSNR + 1.

  CLEAR XACCIT.
  CLEAR XACCCR.

  MOVE-CORRESPONDING VBRK TO XACCIT.
  CLEAR: XACCIT-LAND1, XACCIT-KDGRP.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.         " no invoice list
    MOVE-CORRESPONDING XVBRP TO XACCIT.
ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_08 SPOTS ES_SAPLV60B.
    IF NOT JCDACTIVE IS INITIAL AND XACCIT-TXJCD IS INITIAL.
      XACCIT-TXJCD = TXJCD.
      XACCIT-TXJDP = TXJDP.
      XACCIT-TXJLV = TXJLV.
    ENDIF.
    XACCIT-TXJCD = TXJCD.              "fill with current TXJCD value
* fill argentinian fields
    MOVE XVBRP-J_1AREGIO TO XACCIT-GRIRG.
    MOVE XVBRP-J_1AGICD  TO XACCIT-GRICD.
    MOVE XVBRP-J_1ADTYP  TO XACCIT-GITYP.

    IF XVBRP-SHKZG CA 'BX' OR ( NOT VBRK-KNUMA IS INITIAL
                                AND VBRK-KAPPL = 'V' ).

      XACCIT-SHKZG_VA = 'X'.
    ENDIF.
ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_02 SPOTS ES_SAPLV60B.
    IF XVBRP-POSAR = 'D'.
      CLEAR: XACCIT-MATNR.
    ENDIF.

    IF XVBRP-TRANSIT_PLANT IS NOT INITIAL AND XVBRP-VCM_CHAIN_CATEGORY = 'ICSL' AND VBRK-VBTYP = IF_SD_DOC_CATEGORY=>INVOICE .
      XACCIT-WERKS = XVBRP-TRANSIT_PLANT.
    ENDIF.

  ENDIF.

  MOVE-CORRESPONDING XKOMV TO XACCIT.
  IF GO_BIL_ENRICH_ACCOUNTING IS BOUND.
    XACCIT-TAX_COUNTRY = GO_BIL_ENRICH_ACCOUNTING->FILL_EMPTY_TAXCOUNTRY(
                  IV_MWSK1 = XACCIT-MWSK1
                  IV_TAX_COUNTRY = XACCIT-TAX_COUNTRY
                  IS_VBRK = VBRK
                  IS_T001 = T001
                  IS_VBRP = XVBRP ).
  ENDIF.

* fields that are only filled in customer line item
  CLEAR XACCIT-LIFNR.
  CLEAR XACCIT-KUNNR.
  CLEAR XACCIT-ZTERM.
  CLEAR XACCIT-ZLSCH.
  CLEAR XACCIT-MABER.
  CLEAR XACCIT-MANSP.
  CLEAR XACCIT-MSCHL.
  CLEAR XACCIT-VBUND.
  CLEAR XACCIT-VBELN.
  CLEAR XACCIT-XBLNR.
  CLEAR XACCIT-ZUONR.
  CLEAR XACCIT-GJAHR.

*--- SD-SEPA
  INCLUDE SD_SEPA_FAKTURA_007.
*--- SD-SEPA

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ).           "WIA
    CLEAR XACCIT-STCEG.
  ENDIF.

* fields that are only filled in tax lines
  CLEAR XACCIT-KBETR.

  XACCIT-AWTYP = CON_AWTYP_VBRK.
  XACCIT-AWREF = VBRK-VBELN.
  XACCIT-BELNR = VBRK-VBELN.
  IF MODE NE ' '.
    XACCIT-AWORG = 'CORR'.
*** item typetyp = value item for rebate correction
    XACCIT-POSAR = 'A'.
  ENDIF.
  MOVE-CORRESPONDING XKOMK1 TO XACCIT.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    XACCIT-POSNR_SD = XVBRP-POSNR.
  ENDIF.

* cancellation
  XACCIT-AWREF_REV = VBRK-SFAKN.

  XACCCR-MANDT = VBRK-MANDT.
  XACCCR-AWTYP = CON_AWTYP_VBRK.
  XACCCR-AWREF = VBRK-VBELN.
  XACCCR-AWORG = SPACE.
  IF MODE NE ' '.
    XACCCR-AWORG = 'CORR'.
  ENDIF.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
* Delivery based on purchase order, transfer purchase order no.
    IF XVBRP-AUTYP EQ IF_SD_DOC_CATEGORY=>PURCHASE_ORDER.
      CLEAR: XACCIT-AUBEL, XACCIT-AUPOS.
      XACCIT-EBELN = XVBRP-AUBEL.
      XACCIT-EBELP = XVBRP-AUPOS.
    ENDIF.
  ELSE.
    XACCIT-GSBER = TVTA-GSBER.
    CLEAR XACCIT-TXJCD.
  ENDIF.
  XACCCR-KURSF = VBRK-KURRF.

* transfer prices (TP)
  IF XACCIT-KNTYP CA 'bchn'.

    CLEAR XACCIT-MWSKZ.                " TPs not relevant for tax
    XACCIT-XSKRL = 'X'.                " and cash discount

* transfer prices non-statistical, transfer cost statistical
    IF  XACCIT-KNTYP = 'c'          OR
      ( XACCIT-KNTYP = 'b'          AND
        VBRK-VBUND IS NOT INITIAL   AND
       ( XVBRP-NETWR IS NOT INITIAL OR
*        For batch split items consider the net value of the higher level
         ( XVBRP-UECHA IS NOT INITIAL AND UECHA_NETWR IS NOT INITIAL ) ) ).

      IF MODE_TYPES NE '2'.
        CLEAR XACCIT-KSTAT.
      ENDIF.
      XACCIT-XMFRW = 'X'.              " must be set for CO-PA
    ENDIF.

    DATA: KONZERN_CURTP LIKE ACCCR-CURTP,
          PCA_CURTP     LIKE ACCCR-CURTP.

    CALL FUNCTION 'ECPCA_CVTYP_FOR_SD_GET'
      EXPORTING
        I_BUKRS           = XACCIT-BUKRS
      IMPORTING
        E_KONZERN_CURTP   = KONZERN_CURTP
        E_PCA_CURTP       = PCA_CURTP
      EXCEPTIONS
        KOKRS_NOT_FOUND   = 4
        CVTYP_NOT_FOUND   = 8
        CVPROF_NOT_FOUND  = 1
        CVPROF_NOT_ACTIVE = 8
        OTHERS            = 1.

    IF SY-SUBRC = 0.
      IF XACCIT-KNTYP = 'b'.
* group value
        XACCCR-CURTP = KONZERN_CURTP.
        XACCCR-WAERS = XKOMV-KWAEH.
      ELSEIF XACCIT-KNTYP CA 'chn'.
* profit center value
        XACCCR-CURTP = PCA_CURTP.
        XACCCR-WAERS = XKOMV-KWAEH.
      ENDIF.
    ELSE.
* no valuation type found
      MESSAGE E160 RAISING ERROR_01.
    ENDIF.
  ELSE.
* document currency
    XACCCR-CURTP = '00'.
    XACCCR-WAERS = VBRK-WAERK.
  ENDIF.

  IF XACCIT-KNTYP CA 'bchn'.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCCR-WRBTR = XKOMV-KWERT_K.
    ELSE.
      XACCCR-WRBTR = XKOMV-KWERT_K * -1.
    ENDIF.
  ELSE.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCCR-WRBTR = XKOMV-KWERT.
    ELSE.
      XACCCR-WRBTR = XKOMV-KWERT * -1.
    ENDIF.
  ENDIF.

  CLEAR XACCIT-SHKZG.
  IF XACCCR-WRBTR LE 0.
    IF MODE = 'A'.
      XACCIT-BSCHL = CON_BSCHL_40.
      XACCCR-WRBTR = XACCCR-WRBTR * -1.
    ELSE.
      XACCIT-BSCHL = CON_BSCHL_50.
    ENDIF.
  ELSE.
    IF MODE = 'A'.
      XACCIT-BSCHL = CON_BSCHL_50.
      XACCCR-WRBTR = XACCCR-WRBTR * -1.
    ELSE.
      XACCIT-BSCHL = CON_BSCHL_40.
    ENDIF.
  ENDIF.
  XACCIT-HKONT = XKOMV-SAKN1.
  XACCIT-MWSKZ = XKOMV-MWSK1.

* Time dependent taxes: Determine the tax rate validity start date and set the tax calculation date
  IF  CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.

    IF VBRK-LANDTX IS NOT INITIAL. "See Developer Memo of CM 51251 (2020)

      IF CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
        IV_BUKRS = XACCIT-BUKRS
        IV_LAND1 = COND #( WHEN CL_FOT_TXA_UTILITIES=>AGENT->IS_TAX_ABROAD_ACTIVE( T001-BUKRS ) EQ ABAP_TRUE
                           THEN VBRK-LANDTX ELSE T001-LAND1 )
      )  EQ ABAP_TRUE .
*   Note: The tax code is the same for all item pricing elements (xvbrp-mwsk1). So we can just use the start date from the item (xvbrp-txdat_from)
        XACCIT-TXDAT_FROM = XVBRP-TXDAT_FROM.
        XACCIT-TXDAT = COND #( WHEN XVBRP-FBUDA IS NOT INITIAL THEN XVBRP-FBUDA ELSE XVBRP-PRSDT ).
      ENDIF.

    ENDIF.

  ELSE.

    IF CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
      IV_BUKRS = XACCIT-BUKRS
      IV_LAND1 = T001-LAND1
    )  EQ ABAP_TRUE .

      READ TABLE XKOMV[]
      WITH KEY KPOSN = XKOMV-KPOSN
               KOAID = 'D'
      INTO DATA(TAX_LINE).

      IF TAX_LINE IS NOT  INITIAL.
        XACCIT-TXDAT = TAX_LINE-KDATU.
        XACCIT-TXDAT_FROM = CL_WLF_TDT_SERVICE=>GET_TAX_CAL_DATE_FROM( I_BUKRS = XACCIT-BUKRS
                                                          I_MWSKZ = XACCIT-MWSKZ
                                                          I_FBUDA = XACCIT-TXDAT
                                                          I_TAX_COUNTRY = T001-LAND1
                                                        ).
      ENDIF.

    ENDIF.

  ENDIF.

  IF GO_BIL_DFLOW_ORIG_BD_ACCESS IS BOUND AND GO_BIL_DFLOW_ORIG_BD_ACCESS->IS_RELEVANT_FOR_CALCULATION( CORRESPONDING #( VBRK ) ) = ABAP_TRUE
    AND VBRK-FKTYP NE CON_FKTYP_P.
    DATA(LS_ORIGINAL_DOCUMENT) = GO_BIL_DFLOW_ORIG_BD_ACCESS->GET_REFERENCE_INVOICE( IV_BILLING_DOCUMENT      = VBRK-VBELN
                                                                                     IV_BILLING_DOCUMENT_ITEM = XVBRP-POSNR ).
    IF LS_ORIGINAL_DOCUMENT IS NOT INITIAL.
      XACCIT-PREC_AWREF  = LS_ORIGINAL_DOCUMENT-REFERENCE_INVOICE.
      XACCIT-PREC_AWITEM = LS_ORIGINAL_DOCUMENT-REFERENCE_INVC_ITEM.
      XACCIT-PREC_AWTYP  = CON_AWTYP_VBRK.
    ENDIF.
    CLEAR XACCIT-PREC_AWORG.
  ENDIF.

  IF VBRK-KAPPL = 'M ' AND
     ( NOT VBRK-KNUMA IS INITIAL ) AND
     XKOMV-KOAID =  CON_C AND
     ( NOT ( XVBRP-PAOBJNR =  IF_FCO_COPA_PAOBJNR=>C_INIT  OR
             XVBRP-PAOBJNR =  IF_FCO_COPA_PAOBJNR=>C_ZERO ) ) AND
     XKOMV-KSCHL = 'S000'.
* customer billing document from purchasing
* settlement document subsequent settlement
* condition class subsequent settlement
* item relevant for CO-PA
* condition type clearing from revenue from previous settlements
    CALL FUNCTION 'MM_ARRANG_COND_TYPE_BY_KOMV'
      EXPORTING
        I_KNUMV                = XKOMV-KNUMV
        I_VBELN                = VBRK-VBELN
        I_KPOSN                = XVBRP-POSNR
      IMPORTING
        E_KSCHL                = XACCIT-KSCHL
      TABLES
        T_KOMV                 = XKOMV
      EXCEPTIONS
        INTERNAL_ERROR_PRICING = 1
        OTHERS                 = 2.
* then change condition type for COPA
* no error handling
  ENDIF.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    XACCIT-KDAUF = XVBRP-VBELV.
    XACCIT-KDPOS = XVBRP-POSNV.
    XACCIT-PS_PSP_PNR = XVBRP-PS_PSP_PNR.
    IF XVBRP-SKTOF = CON_X.
      XACCIT-XSKRL = ' '.
    ELSE.
      XACCIT-XSKRL = 'X'.
    ENDIF.
  ELSE.
    CLEAR XACCIT-KOSTL.
    XACCIT-ZUONR = XVBRL-VBELN_VF.
  ENDIF.

* cash management
  XACCIT-FIPOS = FMII1-FIPOS.
  XACCIT-FISTL = FMII1-FISTL.
  XACCIT-GEBER = FMII1-FONDS.
  XACCIT-GRANT_NBR = FMII1-GRANT_NBR.
  XACCIT-FKBER     = FMII1-FAREA.
  XACCIT-MEASURE   = FMII1-MEASURE.

* Take over Budget Period if the switch is on:
  IF CL_PSM_CORE_SWITCH_CHECK=>PSM_FM_CORE_BUD_PER_REV_1( ) IS NOT INITIAL.
    MOVE FMII1-BUDGET_PD TO XACCIT-BUDGET_PD.
  ENDIF.


* accruals ( expens account )
  IF XKOMV-KRUEK = 'X' OR
    " period-end billing document
    ( SV_PERIOD_END_ACTIVE                                  = ABAP_TRUE AND
      CL_SD_DOC_CATEGORY_UTIL=>IS_BILL_PERIOD_END( VBRK-VBTYP ) = ABAP_TRUE AND
      NOT XKOMV-SAKN2 IS INITIAL  ).

    CLEAR: XACCIT-KRUEK,               " Initialisation of the flag
           XACCIT-MWSKZ,               " nicht MWSt-relevant
           XACCIT-KSTAT,               " nicht statistisch
           XACCIT-FIPOS,               " nicht haushaltsrelevant
           XACCIT-FISTL,               " nicht haushaltsrelevant
           XACCIT-GEBER,               " nicht haushaltsrelevant
           XACCIT-GRANT_NBR,
           XACCIT-FKBER,
           XACCIT-MEASURE.
    XACCIT-XSKRL = CON_X.             " nicht skontofähig (P30K013866)
*   Set flag "Accrual reversal" for stock price
    IF ( NOT XKOMV-KWERT_K IS INITIAL AND
             XKOMV-KSTEU   CA 'EFH'   AND
             XKOMV-KNTYP   EQ CON_G   AND
             XKOMV-KSTAT   EQ CON_X   AND
             XKOMV-KWAEH   EQ T001-WAERS ).
      XACCIT-KRUEK = CON_X.            " Accrual reversal
    ENDIF.
  ENDIF.

* take item accounting indicator, when KTREL = 'B'
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    CLEAR XACCIT-BEMOT.
    IF XKOMV-KTREL EQ CON_B.
      XACCIT-BEMOT = XVBRP-BEMOT.
    ENDIF.
  ENDIF.

  "{ Begin ENHO J3GD_SAPLV60B IS-EC-CEM /SAPCEM/ECO_ETM }
*$*$* IBU A&D/E&C, project CEM
* vbrk-j_3gkbaul = 3 bedeutet EDI-buchungskreisübergreifende Verrechnung
*                  (Kennzeichen hängt nicht an der Domäne, sondern wird
*                  beim Generieren der SD-Aufträge erzeugt!).
*                  d.h. es wird nur auf der Leistenden Seite (Absender)
*                  die Kontierung ermittelt. Die Debitorenzeile bleibt
*                  wie im Standard
  IF VBRK-J_3GKBAUL = '1' OR           "CEM
*    vbrk-j_3gkbaul = '2' or           "CEM
     VBRK-J_3GKBAUL = '3'.             "CEM
    PERFORM J_3G_XACCIT_MODIFY.        "CEM commented because it is not transported yet
  ENDIF.                               "CEM
* Document is CEM relevant and should be converted (customizing)
* than the unit should be converted into a time dependant unit.
  IF  NOT ( VBRK-J_3GKBAUL IS INITIAL ).
    PERFORM J_3G_XACCIT_MODIFY_UNIT.
  ENDIF.
  "{ End ENHO J3GD_SAPLV60B IS-EC-CEM /SAPCEM/ECO_ETM }

ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_06 SPOTS ES_SAPLV60B.

* old userexits are executed due to upward compatibility
  MOVE-CORRESPONDING XACCIT TO XKOMK3.

  PERFORM USEREXIT_FILL_XKOMK3.
  MOVE-CORRESPONDING XKOMK3 TO XACCIT.

ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_03 SPOTS ES_SAPLV60B.
** start Revenue Recognition Project                         "75170
*  IF cl_sd_doc_category_util=>is_any_invoice_list( vbrk-vbtyp ) EQ abap_false.
*    IF xvbrp-rrrel NE space AND
*       xkomv-kstat IS INITIAL.
*
**     Translate KBFLAG into LVS_KONVFLAG
*      CALL FUNCTION 'SD_BITS_TO_WORKAREA'
*        EXPORTING
*          i_bits     = xkomv-kbflag
*          i_wa_type  = 'KONVFLAG'
*        IMPORTING
*          e_workarea = lvs_konvflag
*        EXCEPTIONS
*          OTHERS     = 1.
*
**     Directly to revenue account ?
**     Checking field2, because of older code version. For future times
**     field5 will be used, because code must be upwardly compatible
*      IF lvs_konvflag-field2 IS INITIAL AND
*         lvs_konvflag-field5 IS INITIAL.
**       use deferred revenue account instead of revenue account
*        xaccit-hkont = xkomv-sakn2.
*        xacccr-mandt = xaccit-mandt.
*        xaccit-ktosl = xkomv-kvsl1.
*      ELSE.
**       Not relevant for revenue recognition
*        CLEAR: xaccit-rrrel.
*      ENDIF.
*    ENDIF.
*  ENDIF.
** end Revenue Recognition Project "75170
* special treatment for self billing
  IF NOT CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( IV_VBTYP = VBRK-VBTYP ).
    IF NOT XVBRP-AUGRU_AUFT IS INITIAL.
      SELECT SINGLE * FROM TVAU INTO LS_TVAU WHERE AUGRU = XVBRP-AUGRU_AUFT.
      IF SY-SUBRC IS INITIAL.
        IF LS_TVAU-VAUGV = 'X'.
          XACCIT-POSAR = 'A'.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
* to make sure that quantities are correct in CO-PA, the flag XMFRW
* must be set for the first non-statistical line per invoice item

  IF   XACCIT-KSTAT = 'X'
  OR ( XACCIT-POSAR <> SPACE AND
     XACCIT-POSAR <> 'C' ).
    XACCIT-XMFRW = ' '.
  ELSE.
    IF  XACCIT-AWREF    = GD_LAST_AWREF
    AND XACCIT-POSNR_SD = GD_LAST_POSNR_SD.
      XACCIT-XMFRW = ' '.
    ELSE.
      XACCIT-XMFRW = 'X'.
    ENDIF.
    GD_LAST_AWREF = XACCIT-AWREF.
    GD_LAST_POSNR_SD = XACCIT-POSNR_SD.
  ENDIF.

* Statistical Conditions with 'Relevant for Account Based CO-PA'-Flag
  IF  XACCIT-KSTAT = 'X' AND XACCIT-KRUEK IS INITIAL AND XKOMV-IS_ACCT_DETN_RELEVANT = 'X'.
    XACCIT-GKONT = XKOMV-SAKN2.
  ENDIF.

* Differential billing: Additional checks and data processing.
  CLEAR LV_KRUEK_FROM_BILL_DIFF.
  IF SV_BILL_DIFF_ACTIVE = ABAP_TRUE.
    CALL FUNCTION 'SD_BILL_DIFF_CHCK_DIFF_CAPABLE'
      EXPORTING
        IS_VBRK            = VBRK
      IMPORTING
        EV_IS_DIFF_CAPABLE = LV_IS_DIFF_CAPABLE.
    IF LV_IS_DIFF_CAPABLE IS NOT INITIAL.
      PERFORM ACCOUNTING_ITEM_LINE_BILL_DIFF CHANGING LV_KRUEK_FROM_BILL_DIFF.
    ENDIF.
  ENDIF.

* call badi_sd_accounting method: accounting_item_line
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
      CALL BADI GR_SD_ACCOUNTING_BADI->ACCOUNTING_ITEM_LINE
        EXPORTING
          FVBRK  = VBRK
          FVBRP  = XVBRP
          FKOMV  = XKOMV
          FFMII1 = FMII1
        CHANGING
          FACCIT = XACCIT.
    ENDIF.
  ENDIF.

* userexit G/L account item
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    CALL CUSTOMER-FUNCTION '004'
      EXPORTING
        XACCIT = XACCIT
        VBRK   = VBRK
        XVBRP  = XVBRP
        XKOMV  = XKOMV
      IMPORTING
        XACCIT = XACCIT.
  ENDIF.

* set Debit/Credit Indicator
  CLEAR XACCIT-KZZUAB.
  IF XACCCR-WRBTR LT 0.
    MOVE 'X' TO XACCIT-KZZUAB.
  ENDIF.

  XACCIT-POSNR = POSNR.
ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_05 SPOTS ES_SAPLV60B.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
* start Revenue Recognition Project
*    IF xvbrp-rrrel         NE space   AND
*       xkomv-kstat         IS INITIAL AND
*       lvs_konvflag-field2 IS INITIAL AND
*       lvs_konvflag-field5 IS INITIAL AND
*       xkomv-kruek         IS INITIAL.
*      rr_accit = xaccit.
*      APPEND rr_accit.
*    ELSEIF xvbrp-rrrel      IS INITIAL OR
* end Revenue Recognition Project

    IF XVBRP-RRREL          IS INITIAL OR
       NOT XVBRP-VBELV      IS INITIAL OR
       NOT XVBRP-POSNV      IS INITIAL OR
       NOT XVBRP-AUFNR      IS INITIAL OR
       NOT XVBRP-PS_PSP_PNR IS INITIAL OR
       XKOMV-KSTAT          IS INITIAL OR
       XKOMV-KNTYP          NA 'Ghn'   OR
     ( XKOMV-KNTYP EQ CON_G AND NOT XKOMV-KRUEK IS INITIAL ).
      APPEND XACCIT.
    ENDIF.
  ELSE.
    APPEND XACCIT.
  ENDIF.

  XACCCR-POSNR = POSNR.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
* start revenue recognition project
*    IF xvbrp-rrrel         NE space   AND
*       xkomv-kstat         IS INITIAL AND
*       lvs_konvflag-field2 IS INITIAL AND
*       lvs_konvflag-field5 IS INITIAL AND
*       xkomv-kruek         IS INITIAL.
*      rr_acccr = xacccr.
*      APPEND rr_acccr.
*    ELSEIF xvbrp-rrrel          IS INITIAL OR
* end revenue recognition project

    IF     XVBRP-RRREL      IS INITIAL OR
       NOT XVBRP-VBELV      IS INITIAL OR
       NOT XVBRP-POSNV      IS INITIAL OR
       NOT XVBRP-AUFNR      IS INITIAL OR
       NOT XVBRP-PS_PSP_PNR IS INITIAL OR
       XKOMV-KSTAT          IS INITIAL OR
       XKOMV-KNTYP          NA 'Ghn'    OR
      ( XKOMV-KNTYP EQ CON_G AND NOT XKOMV-KRUEK IS INITIAL ).
      APPEND XACCCR.

* Begin extra solution stock price in company code currency
      IF ( NOT XKOMV-KWERT_K IS INITIAL AND
               XKOMV-KSTEU   CA 'EFH'   AND
               XKOMV-KNTYP   EQ CON_G   AND
               XKOMV-KSTAT   EQ CON_X   AND
               XKOMV-KWAEH   EQ T001-WAERS ).
        OLDXACCCR    = XACCCR.
        XACCCR-CURTP = '10'.
        XACCCR-WAERS = XKOMV-KWAEH.
        XACCCR-KURSF = VBRK-KURRF.
        IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
          XACCCR-WRBTR = XKOMV-KWERT_K.
        ELSE.
          XACCCR-WRBTR = XKOMV-KWERT_K * -1.
        ENDIF.
*       Third party case
        IF MODE EQ CON_A.
          XACCCR-WRBTR = XACCCR-WRBTR * -1.
        ENDIF.
        APPEND XACCCR.
        XACCCR = OLDXACCCR.
      ENDIF.
* End extra solution stock price in company code currency

* rebate accrual clearing: value in local currency too
      IF NOT VBRK-KNUMA IS INITIAL AND
         VBRK-FKTYP = 'B' AND
         XKOMV-KRECH = 'B' AND
         XKOMV-KRUEK = 'X'.
        OLDXACCCR = XACCCR.
        XACCCR-CURTP = '10'.
        XACCCR-WAERS = XKOMV-WAERS.
        XACCCR-KURSF = VBRK-KURRF.
        IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
          XACCCR-WRBTR = XKOMV-KBETR.
        ELSE.
          XACCCR-WRBTR = XKOMV-KBETR * -1.
        ENDIF.
        IF ( XKOMV-KBETR > 0 AND XKOMV-KWERT < 0 ) OR
           ( XKOMV-KBETR < 0 AND XKOMV-KWERT > 0 ).
          XACCCR-WRBTR = XACCCR-WRBTR * -1.
        ENDIF.
*       append xacccr.
        XACCCR = OLDXACCCR.
      ENDIF.
* end of extra solution rebate accrual clearing

    ENDIF.
  ELSE.
    APPEND XACCCR.
  ENDIF.

  "{ Begin ENHO J3GD_SAPLV60B IS-EC-CEM /SAPCEM/ECO_ETM }
*$*$* IBU A&D/E&C, project CEM
* Achtung: wenn vbrk-j_3gkbaul = 3
*                  (=EDI-buchungskreisübergreifende Verrechnung)
* soll die Belastete Seite nicht erzeugt werden, da die Belastung mit
* der Standard-Debitorenzeile (Buchungsschlüssel 01) erfolgt.
  IF VBRK-J_3GKBAUL = '1'.               "CEM
    PERFORM J_3G_XACCIT_DUPLICATE.       "CEM Include not transported yet, after transport remove comment sign
    PERFORM J_3G_XACCCR_DUPLICATE.       "CEM Include not transported yet, after transport remove comment sign
  ENDIF.                                 "CEM

  "{ End ENHO J3GD_SAPLV60B IS-EC-CEM /SAPCEM/ECO_ETM }

ENHANCEMENT-POINT ACCOUNTING_ITEM_LINE_07 SPOTS ES_SAPLV60B.

* posting for accrual balance sheet account
  IF XKOMV-KRUEK = 'X' OR
    " month end invoice
    ( SV_PERIOD_END_ACTIVE                                  = ABAP_TRUE AND
      CL_SD_DOC_CATEGORY_UTIL=>IS_BILL_PERIOD_END( VBRK-VBTYP ) = ABAP_TRUE AND
      NOT XKOMV-SAKN2 IS INITIAL ).

    POSNR = POSNR + 1.
    XACCIT-HKONT = XKOMV-SAKN2.
    CLEAR XACCIT-KSTAT.                "non statistical
    CLEAR XACCIT-SHKZG.                "determined by FI
    IF XACCIT-BSCHL = CON_BSCHL_40.    "change posting key
      XACCIT-BSCHL = CON_BSCHL_50.
    ELSE.
      XACCIT-BSCHL = CON_BSCHL_40.
    ENDIF.

* userexit accruals
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
        CALL BADI GR_SD_ACCOUNTING_BADI->ACCOUNTING_ITEM_LINE_ACCRUALS
          EXPORTING
            FVBRK  = VBRK
            FVBRP  = XVBRP
            FKOMV  = XKOMV
          CHANGING
            FACCIT = XACCIT.
      ENDIF.
      CALL CUSTOMER-FUNCTION '005'
        EXPORTING
          XACCIT = XACCIT
          VBRK   = VBRK
          XVBRP  = XVBRP
          XKOMV  = XKOMV
        IMPORTING
          XACCIT = XACCIT.
    ENDIF.

    XACCIT-POSNR = POSNR.
    APPEND XACCIT.

* document currency
    XACCCR-CURTP = '00'.
    XACCCR-WAERS = VBRK-WAERK.

    IF XACCIT-KNTYP = 'b'.
* value in group currency
      XACCCR-CURTP = KONZERN_CURTP.
      XACCCR-WAERS = XKOMV-KWAEH.
    ELSEIF XACCIT-KNTYP CA 'ch'.
* profit center value
      XACCCR-CURTP = PCA_CURTP.
      XACCCR-WAERS = XKOMV-KWAEH.
    ENDIF.

* posting for accrual balance sheet account
    XACCCR-WRBTR = ( -1 ) * XACCCR-WRBTR.

    XACCCR-POSNR = POSNR.
    APPEND XACCCR.

*   Begin extra solution stock price in company code currency
    IF XACCIT-KRUEK EQ CON_X
*   Only, if the XACCIT-KRUEK was not set by differential billing.
    AND LV_KRUEK_FROM_BILL_DIFF IS INITIAL.
      OLDXACCCR    = XACCCR.
      XACCCR-CURTP = '10'.
      XACCCR-WAERS = XKOMV-KWAEH.
      XACCCR-KURSF = VBRK-KURRF.
      IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
        XACCCR-WRBTR = XKOMV-KWERT_K * -1.
      ELSE.
        XACCCR-WRBTR = XKOMV-KWERT_K.
      ENDIF.
*     Third party case
      IF MODE EQ CON_A.
        XACCCR-WRBTR = XACCCR-WRBTR * -1.
      ENDIF.
      APPEND XACCCR.
      XACCCR = OLDXACCCR.
    ENDIF.
*   End extra solution stock price in company code currency

* rebate accrual clearing: value in local currency too
    IF NOT VBRK-KNUMA IS INITIAL AND
       VBRK-FKTYP = 'B' AND
       XKOMV-KRECH = 'B' AND
       XKOMV-KRUEK = 'X'.
      XACCCR-CURTP = '10'.
      XACCCR-WAERS = XKOMV-WAERS.
      XACCCR-KURSF = VBRK-KURRF.
      IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ) EQ ABAP_FALSE.
        XACCCR-WRBTR = XKOMV-KBETR.
      ELSE.
        XACCCR-WRBTR = XKOMV-KBETR * -1.
      ENDIF.
      IF ( XKOMV-KBETR > 0 AND XKOMV-KWERT < 0 ) OR
         ( XKOMV-KBETR < 0 AND XKOMV-KWERT > 0 ).
        XACCCR-WRBTR = XACCCR-WRBTR * -1.
      ENDIF.
*       append xacccr.
    ENDIF.
* end of extra solution rebate accrual clearing
  ENDIF.

ENDFORM.                    "accounting_item_line

*---------------------------------------------------------------------
*
*       FORM ACCOUNTING_TAX_LINE
*
*---------------------------------------------------------------------
*
*       fill accounting document tax line item
*
*---------------------------------------------------------------------
*
*  -->  XVPRP           table invoice items
*
*  -->  XKOMV           table conditions
*
*  <--  XACCIT, XACCCR  tables tax line item
*
*---------------------------------------------------------------------
*
FORM ACCOUNTING_TAX_LINE.

  DATA: TX_KURRF LIKE VBRK-KURRF.
  DATA: L_T001 LIKE T001.                                   "N1007336

  POSNR = POSNR + 1.

  CLEAR XACCIT.
  CLEAR XACCCR.

  MOVE-CORRESPONDING VBRK TO XACCIT.
  CLEAR: XACCIT-LAND1, XACCIT-KDGRP.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.         " no invoice list
    MOVE-CORRESPONDING XVBRP TO XACCIT.
ENHANCEMENT-POINT ACCOUNTING_TAX_LINE_01 SPOTS ES_SAPLV60B.

    IF XVBRP-TRANSIT_PLANT IS NOT INITIAL AND XVBRP-VCM_CHAIN_CATEGORY = 'ICSL' AND VBRK-VBTYP = IF_SD_DOC_CATEGORY=>INVOICE .
      XACCIT-WERKS = XVBRP-TRANSIT_PLANT.
    ENDIF.

* when xauto is set, FI checks whether there is a revenue line with
* the same tax code as the tax line
* for tax only documents, no check can be made, xauto must be cleared
  ENDIF.
  XACCIT-XAUTO = 'X'.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    IF XVBRP-NETWR IS INITIAL AND NOT XVBRP-MWSBP IS INITIAL
    AND NOT XKOMV-KAWRT IS INITIAL.
      CLEAR: XACCIT-XAUTO.
    ENDIF.
  ENDIF.

  IF NOT T001-XFMCA IS INITIAL.
    XACCIT-FKBER = FMII1-FAREA.
  ENDIF.

  MOVE-CORRESPONDING XKOMV TO XACCIT.
  IF GO_BIL_ENRICH_ACCOUNTING IS BOUND.
    XACCIT-TAX_COUNTRY = GO_BIL_ENRICH_ACCOUNTING->FILL_EMPTY_TAXCOUNTRY(
                  IV_MWSK1 = XACCIT-MWSK1
                  IV_TAX_COUNTRY = XACCIT-TAX_COUNTRY
                  IS_VBRK = VBRK
                  IS_T001 = T001
                  IS_VBRP = XVBRP ).
  ENDIF.

* fields that are only filled in customer line item
  CLEAR XACCIT-LIFNR.
  CLEAR XACCIT-KUNNR.
  CLEAR XACCIT-ZTERM.
  CLEAR XACCIT-ZLSCH.
  CLEAR XACCIT-MABER.
  CLEAR XACCIT-MANSP.
  CLEAR XACCIT-MSCHL.
  CLEAR XACCIT-VBUND.
  CLEAR XACCIT-VBELN.
  CLEAR XACCIT-XBLNR.
  CLEAR XACCIT-ZUONR.
  CLEAR XACCIT-GJAHR.

*--- SD-SEPA
  INCLUDE SD_SEPA_FAKTURA_007.
*--- SD-SEPA

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ).         "WIA
    CLEAR XACCIT-STCEG.
  ENDIF.

* fields that are only filled in line items
  CLEAR XACCIT-SHKZG.                  "determined by FI

* account assignments are not relevant in tax lines
  CLEAR XACCIT-PRCTR.
  CLEAR XACCIT-PPRCTR.
  CLEAR XACCIT-GSBER.
  CLEAR XACCIT-PARGB.
  CLEAR XACCIT-KOSTL.
  CLEAR XACCIT-PS_PSP_PNR.
  CLEAR XACCIT-SERVICE_DOC_ID.
  CLEAR XACCIT-SERVICE_DOC_ITEM_ID.
  CLEAR XACCIT-SERVICE_DOC_TYPE.
  CLEAR XACCIT-AUFNR.
  CLEAR XACCIT-PAOBJNR.

  XACCIT-AWTYP = CON_AWTYP_VBRK.
  XACCIT-AWREF = VBRK-VBELN.
  XACCIT-BELNR = VBRK-VBELN.
  XACCIT-TAXIT = CON_TAXIT_X.
  MOVE-CORRESPONDING XKOMK1 TO XACCIT.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    XACCIT-POSNR_SD = XVBRP-POSNR.
  ENDIF.

* cancellation
  XACCIT-AWREF_REV = VBRK-SFAKN.

  XACCCR-MANDT = VBRK-MANDT.
  XACCCR-AWTYP = CON_AWTYP_VBRK.
  XACCCR-AWREF = VBRK-VBELN.
  XACCCR-AWORG = SPACE.

  XACCCR-WAERS = VBRK-WAERK.
  XACCCR-KURSF = VBRK-KURRF.
  XACCCR-FWBAS = XKOMV-KAWRT.
  XACCCR-WRBTR = XKOMV-KWERT.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ) EQ ABAP_FALSE.
    XACCCR-FWBAS = XACCCR-FWBAS * -1.
    XACCCR-WRBTR = XACCCR-WRBTR * -1.
  ENDIF.

* posting key is set according to the sign of the tax amount
* the tax base is not taken into account ( note 437983 )
  IF EXTERNAL IS INITIAL AND NOT XACCCR-WRBTR EQ 0.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCIT-BSCHL = CON_BSCHL_50.
      IF XACCCR-WRBTR GE 0.
        XACCIT-BSCHL =  CON_BSCHL_40.
      ENDIF.
    ELSE.
      XACCIT-BSCHL =  CON_BSCHL_40.
      IF XACCCR-WRBTR LE 0.
        XACCIT-BSCHL = CON_BSCHL_50.
      ENDIF.
    ENDIF.
  ELSE.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ).
      XACCIT-BSCHL = CON_BSCHL_50.
      IF XACCCR-WRBTR GE 0 AND XACCCR-FWBAS GE 0.
        XACCIT-BSCHL =  CON_BSCHL_40.
      ENDIF.
    ELSE.
      XACCIT-BSCHL =  CON_BSCHL_40.
      IF XACCCR-WRBTR LE 0 AND XACCCR-FWBAS LE 0.
        XACCIT-BSCHL = CON_BSCHL_50.
      ENDIF.
    ENDIF.
  ENDIF.

  IF XKOMV-KRECH = CON_H.
    SUBTRACT XACCCR-WRBTR FROM XACCCR-FWBAS.
  ENDIF.

  XACCIT-MWSKZ = XKOMV-MWSK1.
  XACCIT-KTOSL = XKOMV-KVSL1.
  CLEAR XACCIT-TXJCD.
  CASE XKOMV-KNTYP.
    WHEN '1'.
      XACCIT-TXJCD = TXJCD1.
    WHEN '2'.
      XACCIT-TXJCD = TXJCD2.
    WHEN '3'.
      XACCIT-TXJCD = TXJCD3.
    WHEN '4'.
      XACCIT-TXJCD = TXJCD.
  ENDCASE.
  IF XKOMV-KNTYP CA '1234' AND XKOMV-TXJLV IS INITIAL.
    XACCIT-TXJLV = XKOMV-KNTYP.
  ENDIF.
  IF NOT XACCIT-TXJLV IS INITIAL.
    XACCIT-TXJDP = TXJCD.
  ENDIF.
  IF NOT JCDACTIVE IS INITIAL AND XACCIT-TXJCD IS INITIAL.
    XACCIT-TXJCD = TXJCD.
    XACCIT-TXJDP = TXJDP.
* use the level determined for the company code only as fallback
    IF XACCIT-TXJLV IS INITIAL.
      XACCIT-TXJLV = TXJLV.
    ENDIF.
  ENDIF.

* set TXDAT for FI-CA
  IF NOT XACCIT-TXJCD IS INITIAL AND NOT VBRK-VKONT IS INITIAL.
    IF NOT XVBRP-FBUDA IS INITIAL.
      XACCIT-TXDAT = XVBRP-FBUDA.
    ELSE.
      XACCIT-TXDAT = XVBRP-PRSDT.
    ENDIF.
  ENDIF.
* end set TXDAT

* Time dependent taxes: Determine the tax rate validity start date
* Note:
* - The tax code is the same for all item pricing elements (xvbrp-mwsk1). So we can just use the start date from the item (xvbrp-txdat_from)
* - The tax calculation date must not be filled for tax lines
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.

    IF VBRK-LANDTX IS NOT INITIAL. "See Developer Memo of CM 51251 (2020)

      IF CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
        IV_BUKRS = XACCIT-BUKRS
        IV_LAND1 = COND #( WHEN CL_FOT_TXA_UTILITIES=>AGENT->IS_TAX_ABROAD_ACTIVE( T001-BUKRS ) EQ ABAP_TRUE
                           THEN VBRK-LANDTX ELSE T001-LAND1 )
      ) EQ ABAP_TRUE.
        XACCIT-TXDAT_FROM = XVBRP-TXDAT_FROM.
        CLEAR XACCIT-TXDAT.
      ENDIF.

    ENDIF.

  ELSE.
    IF CL_FOT_TDT_CMN_UTIL=>GET( )->IS_TIME_DEP_AND_NO_TAXJUR(
      IV_BUKRS = XACCIT-BUKRS
      IV_LAND1 = T001-LAND1
    ) EQ ABAP_TRUE .
      XACCIT-TXDAT_FROM = CL_WLF_TDT_SERVICE=>GET_TAX_CAL_DATE_FROM( I_BUKRS = XACCIT-BUKRS
                                                                         I_MWSKZ = XACCIT-MWSKZ
                                                                         I_FBUDA = XKOMV-KDATU
                                                                         I_TAX_COUNTRY = T001-LAND1
                                                                       ).
      CLEAR XACCIT-TXDAT.
    ENDIF.
  ENDIF.

  IF GO_BIL_DFLOW_ORIG_BD_ACCESS IS BOUND AND GO_BIL_DFLOW_ORIG_BD_ACCESS->IS_RELEVANT_FOR_CALCULATION( CORRESPONDING #( VBRK ) ) = ABAP_TRUE
     AND VBRK-FKTYP NE CON_FKTYP_P.
    DATA(LS_ORIGINAL_DOCUMENT) = GO_BIL_DFLOW_ORIG_BD_ACCESS->GET_REFERENCE_INVOICE( IV_BILLING_DOCUMENT      = VBRK-VBELN
                                                                                     IV_BILLING_DOCUMENT_ITEM = XVBRP-POSNR ).
    IF LS_ORIGINAL_DOCUMENT IS NOT INITIAL.
      XACCIT-PREC_AWREF  = LS_ORIGINAL_DOCUMENT-REFERENCE_INVOICE.
      XACCIT-PREC_AWITEM = LS_ORIGINAL_DOCUMENT-REFERENCE_INVC_ITEM.
      XACCIT-PREC_AWTYP  = CON_AWTYP_VBRK.
    ENDIF.
    CLEAR XACCIT-PREC_AWORG.
  ENDIF.

  XACCIT-OLD_DOC_NUMBER = DOCUMENT_OLD.
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE AND EXTERNAL = 'X'.
    XACCIT-TAXPS = XVBRP-POSNR.
  ENDIF.

* fill fields for downpayment request
  IF VBRK-FKTYP = CON_FKTYP_P.
    XACCIT-VBEL2 = ANZ_VGBEL.          "Belegnummer
    XACCIT-POSN2 = ANZ_VGPOS.          "Positionsnummer
    XACCIT-BSTAT = CON_BSTAT_S.        "Belegstatus
  ENDIF.

  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
* fill fields for purchase order
    IF XVBRP-AUTYP EQ IF_SD_DOC_CATEGORY=>PURCHASE_ORDER.
      CLEAR: XACCIT-AUBEL, XACCIT-AUPOS.
      XACCIT-EBELN = XVBRP-AUBEL.
      XACCIT-EBELP = XVBRP-AUPOS.
    ENDIF.
  ENDIF.

* BADI call for tax line item
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
      CALL BADI GR_SD_ACCOUNTING_BADI->ACCOUNTING_TAX_LINE
        EXPORTING
          FVBRK   = VBRK
          FVBRP   = XVBRP
          FKOMV   = XKOMV
          F_TXJCD = TXJCD
        CHANGING
          FXACCIT = XACCIT.
    ENDIF.
  ENDIF.

* begin of note 1050398
  DATA: L_BKPF TYPE BKPF.
  DATA: L_ACCHD TYPE ACCHD.
  CLEAR L_ACCHD.
  IF VBRK-WAERK NE T001-WAERS                               "N2053544
  AND VBRK-VKONT IS INITIAL.
* only for posting in foreign currency                      "N2053544
    READ TABLE XACCHD INTO L_ACCHD INDEX 1.
    MOVE-CORRESPONDING XACCHD TO L_BKPF.
    MOVE-CORRESPONDING XACCIT TO L_BKPF.
* begin of note 2053544
    IF VBRK-VBTYP NE IF_SD_DOC_CATEGORY=>INVOICE_CANCEL.
      CALL FUNCTION 'FI_TAX_GET_TXKRS'
        EXPORTING
          I_BUKRS                     = VBRK-BUKRS
          I_CURR_FORGN                = VBRK-WAERK
          I_CURR_LOCAL                = T001-WAERS
*         i_bldat                     = xaccit-bldat            "N1953410
*         i_budat                     = xaccit-bldat            "N1953410
          I_BLDAT                     = L_BKPF-BLDAT            "N1953410
          I_BUDAT                     = L_BKPF-BUDAT            "N1953410
*         I_VATDATE                   =
          I_BKPF                      = L_BKPF
          I_KURST                     = VBRK-KURST              "N1953410
        IMPORTING
          E_TXKRS                     = TX_KURRF
        EXCEPTIONS
          ERROR_READING_BUKRS         = 1
          ERROR_READING_EXCHANGE_RATE = 2
          OTHERS                      = 3.
      IF ( SY-SUBRC IS INITIAL AND TX_KURRF NE 0 ).
        XACCCR-KURSF = TX_KURRF.
      ENDIF.
    ELSE.
* for cancellations.
      IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE AND
         VBRK-SFAKN IS NOT INITIAL. " not item cancellation
* keine Rechnungsliste
        CLEAR LOC_BKPF.
        CALL FUNCTION 'SD_DETERMINE_ACCOUNT_INVOICE'
          EXPORTING
            LOC_VBRP = XVBRP
            LOC_VBRK = VBRK
          IMPORTING
            LOC_BKPF = LOC_BKPF
          EXCEPTIONS
            OTHERS   = 0.
        IF NOT LOC_BKPF-TXKRS IS INITIAL.
* take exchange rate for tax items from BKPF if available
          XACCCR-KURSF = LOC_BKPF-TXKRS.
        ELSEIF NOT LOC_BKPF-KURSF IS INITIAL.
          XACCCR-KURSF = LOC_BKPF-KURSF.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
* end of note 2053544

* userexit tax line item
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    CALL CUSTOMER-FUNCTION '006'
      EXPORTING
        XACCIT = XACCIT
        VBRK   = VBRK
        XVBRP  = XVBRP
        XKOMV  = XKOMV
      IMPORTING
        XACCIT = XACCIT.
  ENDIF.

  XACCIT-POSNR = POSNR.

  APPEND XACCIT.

* document currency
  XACCCR-CURTP = '00'.
  XACCCR-WAERS = VBRK-WAERK.

  XACCCR-POSNR = POSNR.
  APPEND XACCCR.

ENDFORM.                    "accounting_tax_line

*---------------------------------------------------------------------
*
*       FORM READ_FI_PERIODE_INVOICEDATE
*
*---------------------------------------------------------------------
*
*  -->  determine posting period from invoice date
*
*---------------------------------------------------------------------
*
FORM READ_FI_PERIODE_INVOICEDATE.

  CALL FUNCTION 'FI_PERIOD_DETERMINE'
    EXPORTING
      I_BUDAT        = VBRK-FKDAT
      I_BUKRS        = VBRK-BUKRS
    IMPORTING
      E_GJAHR        = FI_PERIODE-GJAHR
      E_MONAT        = FI_PERIODE-MONAT
    EXCEPTIONS
      FISCAL_YEAR    = 1
      PERIOD         = 2
      PERIOD_VERSION = 3
      POSTING_PERIOD = 4
      SPECIAL_PERIOD = 5
      VERSION        = 6
      POSTING_DATE   = 7
      OTHERS         = 8.

ENDFORM.                    "read_fi_periode_invoicedate

*---------------------------------------------------------------------
*
*       FORM FREE_OF_CHARGE_LINE
*
*---------------------------------------------------------------------
*
*---------------------------------------------------------------------
*
FORM FREE_OF_CHARGE_LINE.

  CHECK FREE-OF-CHARGE = 'X'.
  CHECK XVBRP-PAOBJNR <> IF_FCO_COPA_PAOBJNR=>C_INIT  AND  XVBRP-PAOBJNR <> IF_FCO_COPA_PAOBJNR=>C_ZERO .
  CHECK XVBRP-FKIMG NE 0.
  CHECK XVBRP-NETWR = 0.

  POSNR = POSNR + 1.

  CLEAR XACCIT.
  CLEAR XACCCR.
  XACCIT-KSTAT = 'X'.
  MOVE-CORRESPONDING VBRK  TO XACCIT.
  CLEAR: XACCIT-LAND1, XACCIT-KDGRP.
  MOVE-CORRESPONDING XVBRP TO XACCIT.

ENHANCEMENT-POINT FREE_OF_CHARGE_LINE_01 SPOTS ES_SAPLV60B.
  IF XVBRP-SHKZG CA 'BX' OR ( NOT VBRK-KNUMA IS INITIAL
               AND VBRK-KAPPL = 'V' AND XKOMV-KSTEU = 'E' ).
    XACCIT-SHKZG_VA = 'X'.
  ENDIF.

  CLEAR XACCIT-ZTERM.
  CLEAR XACCIT-MABER.
  CLEAR XACCIT-MANSP.
  CLEAR XACCIT-MSCHL.
  CLEAR XACCIT-STCEG.
  CLEAR XACCIT-VBELN.
  CLEAR XACCIT-XBLNR.
  CLEAR XACCIT-ZUONR.
  CLEAR XACCIT-GJAHR.
  CLEAR XACCIT-RRREL.

  XACCIT-AWTYP = CON_AWTYP_VBRK.
  XACCIT-AWREF = VBRK-VBELN.
  XACCIT-BELNR = VBRK-VBELN.
  MOVE-CORRESPONDING XKOMK1 TO XACCIT.
  XACCIT-POSNR_SD = XVBRP-POSNR.
  XACCIT-AWREF_REV = VBRK-SFAKN.

  IF XVBRP-AUTYP EQ IF_SD_DOC_CATEGORY=>PURCHASE_ORDER.
    CLEAR: XACCIT-AUBEL, XACCIT-AUPOS.
    XACCIT-EBELN = XVBRP-AUBEL.
    XACCIT-EBELP = XVBRP-AUPOS.
  ENDIF.
  XACCIT-KDAUF = XVBRP-VBELV.
  XACCIT-KDPOS = XVBRP-POSNV.
  XACCIT-PS_PSP_PNR = XVBRP-PS_PSP_PNR.

  IF XVBRP-TRANSIT_PLANT IS NOT INITIAL AND XVBRP-VCM_CHAIN_CATEGORY = 'ICSL' AND VBRK-VBTYP = IF_SD_DOC_CATEGORY=>INVOICE .
    XACCIT-WERKS = XVBRP-TRANSIT_PLANT.
  ENDIF.

* old userexits are executed due to upward compatibility
  MOVE-CORRESPONDING XACCIT TO XKOMK3.
  PERFORM USEREXIT_FILL_XKOMK3.
  MOVE-CORRESPONDING XKOMK3 TO XACCIT.

* call badi_sd_accounting method: accounting_item_line
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
    IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
      CALL BADI GR_SD_ACCOUNTING_BADI->ACCOUNTING_ITEM_LINE
        EXPORTING
          FVBRK  = VBRK
          FVBRP  = XVBRP
          FKOMV  = XKOMV
        CHANGING
          FACCIT = XACCIT.
    ENDIF.
  ENDIF.

* userexit G/L account item
  CALL CUSTOMER-FUNCTION '004'
    EXPORTING
      XACCIT = XACCIT
      VBRK   = VBRK
      XVBRP  = XVBRP
      XKOMV  = XKOMV
    IMPORTING
      XACCIT = XACCIT.

  XACCIT-POSNR = POSNR.
  APPEND XACCIT.

  XACCCR-MANDT = VBRK-MANDT.
  XACCCR-AWTYP = CON_AWTYP_VBRK.
  XACCCR-AWREF = VBRK-VBELN.
  XACCCR-AWORG = SPACE.
  XACCCR-WAERS = VBRK-WAERK.
  XACCCR-KURSF = VBRK-KURRF.
  XACCCR-CURTP = '00'.
  XACCCR-POSNR = POSNR.
  APPEND XACCCR.

ENDFORM.                    "free_of_charge_line

*---------------------------------------------------------------------
*
*       FORM TXJCD_AUFBEREITEN
*
*---------------------------------------------------------------------
*
*       separate TXJCD in TXJCD1-TXJCD3
*
*---------------------------------------------------------------------
*
FORM TXJCD_AUFBEREITEN.

  DATA: TXJCD_INVALID.
* if tax jurisdictions are active and the ship-to is in a country
* different from the company code country of the Jurisdictioncode
* take a default standard code for the company code ( OBCL )
* if there is already a jurisdiction with the correct length,
* leave it
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE AND NOT JCDACTIVE IS INITIAL
        AND ( T001-LAND1 NE VBRK-LAND1 ).
* check if jurisdiction from XVBRP can be posted
    IF NOT XVBRP-TXJCD IS INITIAL.
      CALL FUNCTION 'TXJCD_CHECK'
        EXPORTING
          I_BUKRS                 = T001-BUKRS
          I_TXJCD                 = XVBRP-TXJCD
          I_CHECK_EXTERNAL        = 'X'                     "N840125
        EXCEPTIONS
          INPUT_PARAMETER_MISSING = 1
          NOT_EXIST               = 2
          PARAMETER_CONFLICT      = 3
          OTHERS                  = 4.
      IF SY-SUBRC <> 0.
        TXJCD_INVALID = CON_X.
      ENDIF.
    ENDIF.
    PERFORM USEREXIT_TXJCD IN PROGRAM SAPLV60B IF FOUND
                    CHANGING TXJCD_INVALID.
    IF XVBRP-TXJCD IS INITIAL OR NOT TXJCD_INVALID IS INITIAL.
      CALL FUNCTION 'FI_TAX_GET_DEFAULT_TXJCD'
        EXPORTING
          I_BUKRS       = VBRK-BUKRS
        IMPORTING
          E_TXJCD       = TXJCD
          E_TXJDP       = TXJDP
          E_TXJLV       = TXJLV
        EXCEPTIONS
          SYSTEM_ERROR  = 1
          ERROR_MESSAGE = 2.
      IF SY-SUBRC NE 0.
        IF CHECK NE CON_X.
*      write error to protocol
          RAISE ERROR_01.
        ELSE.
*      display error directly
          MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
        ENDIF.
      ELSE.
        EXIT.
      ENDIF.
    ENDIF.
  ENDIF.

* read t005
  IF T005-LAND1 NE T001-LAND1.
    CLEAR T005.
    T005-LAND1 = T001-LAND1.
    SELECT SINGLE * FROM T005 WHERE LAND1 = T001-LAND1.
  ENDIF.

* read TTXD
  IF TTXD-KALSM NE T005-KALSM.
    CLEAR TTXD.
    TXJCD = SPACE.
    TTXD-KALSM = T005-KALSM.
    SELECT SINGLE * FROM TTXD WHERE KALSM = T005-KALSM.
  ENDIF.


* only split tax jurisdiction code if a valid split rule
* is maintained in table TTXD.
  IF XVBRP-TXJCD NE SPACE AND
     NOT ( ( TTXD-LENG1 IS INITIAL ) AND
           ( TTXD-LENG2 IS INITIAL ) AND
           ( TTXD-LENG3 IS INITIAL ) AND
           ( TTXD-LENG4 IS INITIAL ) ).
* split KOMK-TXJCD into TXJCD1-3
    IF XVBRP-TXJCD NE TXJCD.
      TXJCD  = XVBRP-TXJCD.
      TXJCD1 = SPACE.
      SY-TABIX = TTXD-LENG1 + TTXD-LENG2 + TTXD-LENG3 + TTXD-LENG4.
      IF SY-TABIX > 0.
        WRITE '000000000000000' TO TXJCD1(SY-TABIX).
      ENDIF.
      TXJCD2 = TXJCD1.
      TXJCD3 = TXJCD1.
      SY-TABIX = TTXD-LENG1.
      IF SY-TABIX > 0.
        WRITE TXJCD TO TXJCD1(SY-TABIX).
        ADD TTXD-LENG2 TO SY-TABIX.
        WRITE TXJCD TO TXJCD2(SY-TABIX).
        ADD TTXD-LENG3 TO SY-TABIX.
        WRITE TXJCD TO TXJCD3(SY-TABIX).
      ENDIF.
    ENDIF.
  ELSE.
    TXJCD  = SPACE.
    TXJCD1 = SPACE.
    TXJCD2 = SPACE.
    TXJCD3 = SPACE.
  ENDIF.

ENDFORM.                    "txjcd_aufbereiten

*---------------------------------------------------------------------
*
*       FORM XACCIT_XNEGP_SET
*
*---------------------------------------------------------------------
*
*       Set negative posting indicator and cancellation reason
*
*---------------------------------------------------------------------
*
FORM XACCIT_XNEGP_SET.

* negative posting for credit memos and cancellations
  IF CL_SD_DOC_CATEGORY_UTIL=>IS_INVOICE_NEGATIVE( VBRK-VBTYP ) OR
     VBRK-VBTYP EQ IF_SD_DOC_CATEGORY=>CREDIT_MEMO_CANCEL.

    " globalization can overwrite the core standard customizing
    PERFORM XACCIT_XNEGP_SET_GLO CHANGING TVFK-XNEGP.

    CASE TVFK-XNEGP.
* no negative posting
      WHEN ' '.
        CLEAR XACCIT-XNEGP.
* negative posting if reference invoice belongs to the same posting
* period
      WHEN 'A'.
* translate invoice date to posting period
        PERFORM READ_FI_PERIODE_INVOICEDATE.
        READ TABLE XVBRP WITH KEY XVBRP_KEY BINARY SEARCH. "#EC CI_SORTED
        CLEAR LOC_BKPF.
        IF SY-SUBRC EQ 0.
          CALL FUNCTION 'SD_DETERMINE_ACCOUNT_INVOICE'
            EXPORTING
              LOC_VBRP = XVBRP
              LOC_VBRK = VBRK
            IMPORTING
              LOC_BKPF = LOC_BKPF
            EXCEPTIONS
              OTHERS   = 0.
          IF LOC_BKPF-MONAT EQ FI_PERIODE-MONAT.
            XACCIT-XNEGP = 'X'.
          ELSE.
            CLEAR XACCIT-XNEGP.
          ENDIF.
        ENDIF.
* always negative posting
      WHEN 'B'.
        XACCIT-XNEGP = 'X'.
* negative posting from reversal reason on VF11
      WHEN 'C'.
        CLEAR XACCIT-XNEGP.
        IF NOT VBRK-STGRD IS INITIAL.
          DATA: L_XNEGP TYPE T041C-XNEGP.
          " detect negative posting flag from T041C
          SELECT SINGLE XNEGP FROM T041C
            INTO L_XNEGP
            WHERE STGRD = VBRK-STGRD.
          IF SY-SUBRC IS INITIAL.
            " fill stgrd and xnegp
            XACCIT-STGRD = VBRK-STGRD.
            XACCIT-XNEGP = L_XNEGP.
          ENDIF.
        ENDIF.
    ENDCASE.
  ENDIF.

ENDFORM.                    "xaccit_xnegp_set

*---------------------------------------------------------------------
*
*       FORM XACCIT_XVALGS_ZFBDT_SET
*
*---------------------------------------------------------------------
*
*       get accounting reference and baseline date info
*
*---------------------------------------------------------------------
*
FORM XACCIT_XVALGS_ZFBDT_SET USING UV_ASSIGN_BASEDOC         TYPE ABAP_BOOL
                                   UV_DETERMINE_BASELINEDATE TYPE ABAP_BOOL.

  " credit memo with value date: use baseline date of the reference
  " invoice if this is not cleared
  IF UV_ASSIGN_BASEDOC EQ ABAP_TRUE OR
     UV_DETERMINE_BASELINEDATE EQ ABAP_TRUE.
    " negative posting and credit memos with value date may be used in
    " parallel therefore the period determination need not be done again
    IF LOC_BKPF-MONAT IS INITIAL.
      READ TABLE XVBRP WITH KEY XVBRP_KEY BINARY SEARCH. "#EC CI_SORTED
      IF SY-SUBRC EQ 0.
        CALL FUNCTION 'SD_DETERMINE_ACCOUNT_INVOICE'
          EXPORTING
            LOC_VBRK = VBRK
            LOC_VBRP = XVBRP
          IMPORTING
            LOC_BKPF = LOC_BKPF
          EXCEPTIONS
            OTHERS   = 0.
      ENDIF.                         " sy-subrc eq 0
    ENDIF.                           " bkpf-monat is initial
    CALL FUNCTION 'SD_READ_ACC_DOC_NOT_CLEARING'
      EXPORTING
        LOC_BKPF = LOC_BKPF
      IMPORTING
        LOC_BSID = LOC_BSID
      EXCEPTIONS
        OTHERS   = 0.

    " FI-CA: adjust baseline date of original invoice
    IF NOT VBRK-VKONT IS INITIAL.
      DATA: LD_FKDAT LIKE VBRK-FKDAT,
            LD_VALTG LIKE VBRK-VALTG,
            LD_VBELN LIKE VBRK-VBELN.

      LD_VBELN = LOC_BKPF-AWKEY.

      SELECT SINGLE FKDAT VALTG FROM VBRK INTO (LD_FKDAT, LD_VALTG)
                                          WHERE VBELN = LD_VBELN.

      LOC_BSID-ZFBDT = LD_FKDAT + LD_VALTG.
    ENDIF.

    IF LOC_BSID-GJAHR IS INITIAL OR LOC_BSID-BUZEI IS INITIAL
    OR LOC_BSID-UMSKZ NE SPACE OR LOC_BSID-UMSKS NE SPACE
    OR LOC_BKPF-BUKRS NE VBRK-BUKRS.
      CLEAR LOC_BKPF.
      CLEAR LOC_BSID.
    ENDIF.

  ENDIF.

ENDFORM.                    "xaccit_xvalgs_zfbdt_set

*---------------------------------------------------------------------
*
*       FORM  CURRENCY_CONVERSION
*
*---------------------------------------------------------------------
*
FORM CURRENCY_CONVERSION.

* convert securevalue to credit currency
  DATA: DOC_CURR_IS_EURO.
  DATA: CREDIT_CURR_IS_EURO.
  DATA: ABSBT LIKE XACCIT-ABSBT.
  DATA: LD_KURST LIKE VBRK-KURST.

  CLEAR: DOC_CURR_IS_EURO, CREDIT_CURR_IS_EURO, ABSBT.
* check whether currencies involved are euro currencies

  CALL FUNCTION 'PRICING_CHECK_EURO_CURRENCY'
    EXPORTING
      CURRENCY_TO_CHECK = VBRK-CMWAE
    IMPORTING
      EURO_FLAG         = CREDIT_CURR_IS_EURO
    EXCEPTIONS
      EURO_CURRENCIES   = 1
      OTHERS            = 2.

  CALL FUNCTION 'PRICING_CHECK_EURO_CURRENCY'
    EXPORTING
      CURRENCY_TO_CHECK = VBRK-WAERK
    IMPORTING
      EURO_FLAG         = DOC_CURR_IS_EURO
    EXCEPTIONS
      EURO_CURRENCIES   = 1
      OTHERS            = 2.

  LD_KURST = VBRK-KURST.
  IF LD_KURST IS INITIAL.
    LD_KURST = 'M'.
  ENDIF.

* Not both currencies involved are EURO currencies
  IF NOT ( CREDIT_CURR_IS_EURO = 'X' AND DOC_CURR_IS_EURO = 'X' ).
* Try also this way if TRY_DIRECT_CONV flag is set
    IF TRY_DIRECT_CONV = 'X'.
* First try direct currency conversion
      CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
        EXPORTING
          DATE             = VBRK-FKDAT
          FOREIGN_AMOUNT   = XACCIT-ABSBT
          FOREIGN_CURRENCY = VBRK-WAERK
          LOCAL_CURRENCY   = VBRK-CMWAE
          TYPE_OF_RATE     = LD_KURST
        IMPORTING
          LOCAL_AMOUNT     = ABSBT
        EXCEPTIONS
          NO_RATE_FOUND    = 1
          OVERFLOW         = 2.

      IF SY-SUBRC = 0.
        XACCIT-ABSBT = ABSBT.
        EXIT.
      ENDIF.
    ENDIF.
    CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
      EXPORTING
        DATE             = VBRK-FKDAT
        FOREIGN_AMOUNT   = XACCIT-ABSBT
        FOREIGN_CURRENCY = VBRK-WAERK
        LOCAL_CURRENCY   = T001-WAERS
        TYPE_OF_RATE     = LD_KURST
      IMPORTING
        LOCAL_AMOUNT     = ABSBT
      EXCEPTIONS
        ERROR_MESSAGE    = 4
        OTHERS           = 4.

    IF NOT SY-SUBRC IS INITIAL.
      IF CHECK NE CON_X.
*      Fehler ins Protokoll schreiben
        RAISE ERROR_01.
      ELSE.
*     Fehler direkt ausgeben
        MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.
    ENDIF.

    CALL FUNCTION 'CONVERT_TO_FOREIGN_CURRENCY'
      EXPORTING
        DATE             = VBRK-FKDAT
        LOCAL_AMOUNT     = ABSBT
        FOREIGN_CURRENCY = VBRK-CMWAE
        LOCAL_CURRENCY   = T001-WAERS
        RATE             = VBRK-CMKUF
        TYPE_OF_RATE     = LD_KURST
      IMPORTING
        FOREIGN_AMOUNT   = ABSBT
      EXCEPTIONS
        ERROR_MESSAGE    = 4
        OTHERS           = 4.

    IF SY-SUBRC = 0.
      XACCIT-ABSBT = ABSBT.
      EXIT.
    ELSE.

      IF CHECK NE CON_X.
*      Fehler ins Protokoll schreiben
        RAISE ERROR_01.
      ELSE.
*     Fehler direkt ausgeben
        MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.
    ENDIF.

  ELSE.
* Both currencies involved are EURO currencies
    CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
      EXPORTING
        DATE             = VBRK-FKDAT
        FOREIGN_AMOUNT   = XACCIT-ABSBT
        FOREIGN_CURRENCY = VBRK-WAERK
        LOCAL_CURRENCY   = VBRK-CMWAE
        TYPE_OF_RATE     = LD_KURST
      IMPORTING
        LOCAL_AMOUNT     = ABSBT
      EXCEPTIONS
        NO_RATE_FOUND    = 1
        OVERFLOW         = 2.
    IF SY-SUBRC = 0.
      XACCIT-ABSBT = ABSBT.
      EXIT.
    ELSE.

      IF CHECK NE CON_X.
*      Fehler ins Protokoll schreiben
        RAISE ERROR_01.
      ELSE.
*     Fehler direkt ausgeben
        MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
              WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.
    ENDIF.

  ENDIF.

ENDFORM.                               " currency_conversion

*---------------------------------------------------------------------
*
*       FORM  FILL_ACCIT_DEB
*
* --------------------------------------------------------------------
*
FORM FILL_ACCIT_DEB USING I_CALL_ID TYPE C
                    CHANGING DA_POS_NETWR LIKE VBRP-NETWR
                    CV_COLLECT_PROCESSED LIKE BOOLE-BOOLE.

  DATA: LV_ACTIVE_BADI_SD_TO_FM TYPE XFELD.

* Payment Service Provider
  DATA: LR_PAYSP TYPE REF TO CL_SLS_PAYSP_INTEGRATION.

  CASE I_CALL_ID.

    WHEN 'C'.
* Coding is only processed in case of:
* => Condition record is not statistical
      CHECK XKOMV-KSTAT IS INITIAL.

* => Calculation step (condition) is active
      CHECK XKOMV-KINAK IS INITIAL.
      CHECK NOT VBRK-VKONT IS INITIAL OR
            NOT DEB_LINES_FOR_COND IS INITIAL.

* In case of non tax condition
      IF XKOMV-KNTYP NE CON_KNTYP_POS.
        IF XKOMV-KOAID NE CON_D OR
         MODE NE SPACE.
*    Transfering condition value to net amount
          XACCIT_DEB-NETWR = XKOMV-KWERT.
*    In case of pricing condition
*    Transferring amount eligible for cash discount
          IF XKOMV-KOAID =  CON_B.
            XACCIT_DEB-SKFBP = XVBRP-SKFBP.
          ENDIF.
        ENDIF.
      ENDIF.

    WHEN 'I'.
* Coding is processed generally,
* if collect wasn't already processed.
      CHECK CV_COLLECT_PROCESSED IS INITIAL.
      XACCIT_DEB-NETWR = XVBRP-NETWR.
      XACCIT_DEB-SKFBP = XVBRP-SKFBP.
      CLEAR: XKOMV.

  ENDCASE.

  CV_COLLECT_PROCESSED = 'X'.

  PERFORM FREE_OF_CHARGE_LINE. " P30K123459

* Fill table for creation of the customer line items.
* Determine G/L accounts for cash sales or reconciliation account
* in case of cancellations, use invoice type customizing of the
* original invoice.
  IF NOT VBRK-SFAKN IS INITIAL.
    SELECT SINGLE * FROM VBRK
                    INTO *VBRK
                    WHERE VBELN = VBRK-SFAKN.
    SELECT SINGLE * FROM TVFK
                    INTO *TVFK
                                    WHERE FKART = *VBRK-FKART.
    IF NOT *TVFK-KALSMCB IS INITIAL OR
       NOT *TVFK-KALSMCC IS INITIAL.
      TVFK-KALSMCC = *TVFK-KALSMCC.
      TVFK-KALSMCB = *TVFK-KALSMCB.
    ENDIF.
  ENDIF.

  "Localization Brazil
  CALL FUNCTION 'J_1BSA_COMPONENT_ACTIVE'
    EXPORTING
      BUKRS                = T001-BUKRS
      COMPONENT            = 'BR'
    EXCEPTIONS
      COMPONENT_NOT_ACTIVE = 1
      OTHERS               = 2.

  IF SY-SUBRC = 0 AND CL_COS_UTILITIES=>IS_CLOUD( ) = ABAP_TRUE.
    "For Cloud, override some Billing Type configuration with country specific settings
    CL_LOGBR_BILLING_TYPE_SETTINGS=>GET_INSTANCE( )->IF_LOGBR_BILLING_TYPE_SETTINGS~EXECUTE(
      EXPORTING
        IV_COUNTRY_KEY         = T001-LAND1
      CHANGING
        CS_BILLING_TYPE_CONFIG = TVFK
    ).
  ENDIF.

  IF TVFK-KALSMCB NE SPACE  OR
     TVFK-KALSMCC NE SPACE.
    CLEAR KOMKCV.
    MOVE-CORRESPONDING VBRK TO KOMKCV.
    KOMKCV-KTOPL = T001-KTOPL.
* Reconciliation account from account determination procedure
    IF TVFK-KALSMCB NE SPACE.
      KOMKCV-KAPPL = 'VB'.
      KOMKCV-KALSMC = TVFK-KALSMCB.
    ELSE.
* Reconciliation account for cash sales
      KOMKCV-KAPPL = 'VC'.
      KOMKCV-KALSMC = TVFK-KALSMCC.
    ENDIF.

    READ TABLE KOMKCV WITH KEY KOMKCV-KEY_UC.

    IF SY-SUBRC NE 0.
      APPEND KOMKCV.
    ENDIF.

    MOVE-CORRESPONDING XVBRP TO KOMPCV.

* Userexit for the komkcv- and kompcv-structures
    CALL CUSTOMER-FUNCTION '011'
      EXPORTING
        I_XVBRP      = XVBRP
        I_VBRK       = VBRK
        I_KOMKCV     = KOMKCV
        I_KOMPCV     = KOMPCV
        I_DOC_NUMBER = XVBRP_KEY-VBELN
        I_XKOMV      = XKOMV
      IMPORTING
        E_KOMKCV     = KOMKCV
        E_KOMPCV     = KOMPCV
      TABLES
        T_XVBPA      = XVBPA.

    CALL FUNCTION 'ACCOUNT_ALLOCATION_GENERAL'
      EXPORTING
        I_APPLICATION          = KOMKCV-KAPPL
        I_SCHEME               = KOMKCV-KALSMC
        I_HEADER_COMMUNICATION = KOMKCV
        I_ITEM_COMMUNICATION   = KOMPCV
        I_PROTOKOLL            = ' '
      IMPORTING
        E_C000                 = C000
      EXCEPTIONS
        OTHERS                 = 0.

    XACCIT_DEB-HKONT = C000-SAKN1.

    IF TVFK-KALSMCC NE SPACE AND
       NOT XACCIT_DEB-HKONT IS INITIAL.

* BADI_SD_TO_FM, check if Public Sector is active
* no cash accounts allowed for PSM-FG and down payment request
      CALL FUNCTION 'GET_HANDLE_SD_TO_FM'
        IMPORTING
          ACTIVE = LV_ACTIVE_BADI_SD_TO_FM.

      IF LV_ACTIVE_BADI_SD_TO_FM EQ 'X' AND
        VBRK-FKTYP EQ CON_FKTYP_P.
        CLEAR XACCIT_DEB-HKONT.
        CLEAR XACCIT_DEB-CASH.
      ELSE.
        XACCIT_DEB-CASH = 'X'.
      ENDIF.

    ENDIF.

  ELSE.
    IF CASH_SALE_ACCOUNT IS INITIAL.
      CLEAR XACCIT_DEB-HKONT.
    ELSE.
      XACCIT_DEB-HKONT = CASH_SALE_ACCOUNT.
      XACCIT_DEB-CASH = 'X'.
    ENDIF.
  ENDIF.

* Settlement2Invoice
  IF CL_ERP_EHP_SWITCH_CHECK=>ERP_CF_SFWS_1( ) EQ ABAP_TRUE.
    IF NOT VBRK-FK_SOURCE_SYS IS INITIAL.
      XACCIT_DEB-CASE_GUID_CORE = XVBRP-DISPUTE_CASE.
      IF NOT VBRK-SFAKN IS INITIAL.
        XACCIT_DEB-DISPUTE_IF_TYPE = CON_B.
      ELSE.
        XACCIT_DEB-DISPUTE_IF_TYPE = CON_A.
      ENDIF.
    ENDIF.
    IF NOT CLEARING_ACCOUNT IS INITIAL.
      XACCIT_DEB-HKONT = CLEARING_ACCOUNT.
      XACCIT_DEB-CASH = CON_X.
    ENDIF.
  ENDIF.
* Fill fields for downpayment requests
  IF VBRK-FKTYP EQ CON_FKTYP_P.
    XACCIT_DEB-VBEL2      = ANZ_VGBEL.        " Belegnummer
    XACCIT_DEB-POSN2      = ANZ_VGPOS.        " Positionsnummer
    XACCIT_DEB-PS_PSP_PNR = XVBRP-PS_PSP_PNR. " PSP-Element
    XACCIT_DEB-PRCTR      = XVBRP-PRCTR.      " Profit Center
    XACCIT_DEB-POSNR_SD   = XACCIT-POSNR_SD.  " POSNR_SD
* Fill fields for cash management
    XACCIT_DEB-FIPOS     = FMII1-FIPOS.           " Finanzposition
    XACCIT_DEB-FISTL     = FMII1-FISTL.           " Finanzstelle
    XACCIT_DEB-GEBER     = FMII1-FONDS.           " Fonds
    XACCIT_DEB-GRANT_NBR = FMII1-GRANT_NBR.       " Grant
    XACCIT_DEB-FKBER     = FMII1-FAREA.           " Functional Area
    XACCIT_DEB-MEASURE   = FMII1-MEASURE.         " Funded Program
* Take over Budget Period if the switch is on:
    IF CL_PSM_CORE_SWITCH_CHECK=>PSM_FM_CORE_BUD_PER_REV_1( ) IS NOT INITIAL.
      MOVE FMII1-BUDGET_PD TO XACCIT_DEB-BUDGET_PD.
    ENDIF.

  ENDIF.

* Payment Service Provider
  IF CL_OPS_SWITCH_CHECK=>SD_SFWS_SC4( ) EQ ABAP_TRUE.
    IF VBRK-SPPAYM EQ '01'. "cv_sppaym01.
      LR_PAYSP = CL_SLS_PAYSP_INTEGRATION=>GET_INSTANCE( ).
      IF LR_PAYSP IS BOUND.
        CALL METHOD LR_PAYSP->FILL_FI_DATA
          EXPORTING
            IV_DOCUMENT_ID = XVBRP-AUBEL
          IMPORTING
            EV_PAYS_PROV   = XACCIT_DEB-PAYS_PROV
            EV_PAYS_TRAN   = XACCIT_DEB-PAYS_TRAN.
      ENDIF.
    ENDIF.
  ENDIF.

  XACCIT_DEB-VERTT = XVBRP-VERTT.
  XACCIT_DEB-VERTN = XVBRP-VERTN.
  XACCIT_DEB-SGTXT = XVBRP-SGTXT.
* xaccit_deb-netwr = xvbrp-netwr + da_pos_netwr.
* xaccit_deb-skfbp = xvbrp-skfbp.

* Change due to FI-CA
  XACCIT_DEB-NETWR = XACCIT_DEB-NETWR + DA_POS_NETWR.
  XACCIT_DEB-BRTWR = XACCIT_DEB-NETWR + XACCIT_DEB-MWSBP.

* Update withholding tax and set index QST table
  PERFORM WT_XACCIT_WT_FILL CHANGING XACCIT_DEB-WT_KEY.

* BADI call for FI-CA
  IF NOT VBRK-VKONT IS INITIAL OR
     NOT DEB_LINES_FOR_COND IS INITIAL.

    DATA: LD_SUBRC LIKE SY-SUBRC.
    IF CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INVOICE_LIST( VBRK-VBTYP ) EQ ABAP_FALSE.
      IF BADI_SD_ACCOUNTING_ACTIVE = 'X'.
        CALL BADI GR_SD_ACCOUNTING_BADI->FILL_ACCIT_DEB
          EXPORTING
            FVBRK       = VBRK
            FVBRP       = XVBRP
            FKOMV       = XKOMV
            FFMII1      = FMII1
          CHANGING
            FXACCIT_DEB = XACCIT_DEB
            FSUBRC      = LD_SUBRC.
        IF LD_SUBRC NE 0.
          IF CHECK NE CON_X.
*           Write message to protocol
            RAISE ERROR_01.
          ELSE.
*            Send message
            MESSAGE ID SY-MSGID TYPE 'E' NUMBER SY-MSGNO
                WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.

* Userexit items for creation of customer line items
  CALL CUSTOMER-FUNCTION '010'
    EXPORTING
      I_XVBRP      = XVBRP
      I_XACCIT_DEB = XACCIT_DEB
      I_XKOMV      = XKOMV
    IMPORTING
      E_XACCIT_DEB = XACCIT_DEB.
  COLLECT XACCIT_DEB.

* FI-CA: clear, in case of deb lines per condition
  IF NOT VBRK-VKONT IS INITIAL OR
     NOT DEB_LINES_FOR_COND IS INITIAL.
    CLEAR: DA_POS_NETWR,
           XACCIT_DEB.
  ENDIF.

ENDFORM.                    " fill_accit_deb

*---------------------------------------------------------------------
*
*       FORM  ENHANCE_CONTR_REFERENCE
*
* --------------------------------------------------------------------
*
FORM ENHANCE_CONTR_REFERENCE.

  DATA: LR_VBAP TYPE REF TO VBAP.

  CHECK XVBRP-CONTR_DP_ID IS INITIAL.

  CL_SD_SALESORDER_ACCESS=>GET_INSTANCE( )->GET_ITEMS(
    EXPORTING
      IV_SALESORDER_ID    = XVBRP-AUBEL              " Sales and Distribution Document Number
*      it_salesorder_id    =                         " SD Document Numbers, Not Sorted
    IMPORTING
      ET_SALESORDER_ITEMS = DATA(LT_VBAP)            " Table type sales document: Item data
      EV_RETURN_CODE      = DATA(LS_RETURN_CODE) ).  " Return Code

  IF LS_RETURN_CODE EQ 0.

    READ TABLE LT_VBAP WITH KEY VBELN = XVBRP-AUBEL POSNR = XVBRP-AUPOS REFERENCE INTO LR_VBAP.
    IF SY-SUBRC EQ 0.
      XVBRP-CONTR_DP_ID      = LR_VBAP->VGBEL.
      XVBRP-CONTR_DP_ITEM_ID = LR_VBAP->VGPOS.
    ENDIF.
  ENDIF.

ENDFORM.                    " enhance_contr_reference

*---------------------------------------------------------------------*
*
*       FORM map_extensibility_flow_bill
*
*---------------------------------------------------------------------*
*       Map ExtensibilityFlows from Billing to FIN
*---------------------------------------------------------------------*
*
FORM MAP_EXTENSIBILITY_FLOW_BILL.

  DATA: LS_VBRP_CUST_FIELD TYPE SDBILLGDOCITEM_INCL_EEW_PS.
  FIELD-SYMBOLS: <FS_VBRP>  TYPE VBRPVB,
                 <FS_ACCIT> TYPE ACCIT.

  LOOP AT XVBRP ASSIGNING <FS_VBRP> WHERE VBELN = DOCUMENT_OLD.

*   Check if any custom fields are used
    MOVE-CORRESPONDING <FS_VBRP> TO LS_VBRP_CUST_FIELD.
    CHECK LS_VBRP_CUST_FIELD IS NOT INITIAL.

*   modify all lines in XACCIT, which are related to the current XVBRP-POSNR
    LOOP AT XACCIT ASSIGNING <FS_ACCIT> WHERE POSNR_SD = <FS_VBRP>-POSNR.
*     source of these flows
      GET REFERENCE OF <FS_VBRP> INTO DATA(LR_VBRP).
*     target of these flows
      GET REFERENCE OF <FS_ACCIT> INTO DATA(LR_ACCIT).

*     call the data transfer routines.
      TRY.
          CL_CFD_DATA_TRANSFER_FACTORY=>GET_DATA_TRANSFER_RUNTIME( )->TRANSFER_DATA(
          IV_DATA_TRANSFER_NAME = 'SD_BILDOC_ITEM_2_FINS_CODG_BLK'
          IR_SOURCE_STRUCTURE   = LR_VBRP
          IR_TARGET_STRUCTURE   = LR_ACCIT ).

          CL_CFD_DATA_TRANSFER_FACTORY=>GET_DATA_TRANSFER_RUNTIME( )->TRANSFER_DATA(
          IV_DATA_TRANSFER_NAME = 'SD_BIL_ITM_2_FINS_JRNL_ENT_ITM'
          IR_SOURCE_STRUCTURE   = LR_VBRP
          IR_TARGET_STRUCTURE   = LR_ACCIT ).

        CATCH CX_CFD_DATA_TRANSFER INTO DATA(LX_CFD_DATA_TRANSFER).
*       catch potential exception without error handling
      ENDTRY.

    ENDLOOP.

  ENDLOOP.

ENDFORM.

FORM XACCIT_CLEARING_RELEVANT CHANGING EV_CLRST TYPE STRING.

  DATA: BEGIN OF XACCIT_KUNNR OCCURS 100.
          INCLUDE STRUCTURE ACCIT.
  DATA: END OF XACCIT_KUNNR.

  LOOP AT XACCIT.
    IF XACCIT-KUNNR IS NOT INITIAL.
      APPEND XACCIT TO XACCIT_KUNNR.
    ENDIF.
  ENDLOOP.

  DESCRIBE TABLE XACCIT_KUNNR LINES DATA(LV_KUNNR).

  IF LV_KUNNR > 0.

    SELECT KOART FROM TBSL WHERE BSCHL = @XACCIT_KUNNR-BSCHL
      INTO TABLE @DATA(LT_TBSL).

    LOOP AT LT_TBSL REFERENCE INTO DATA(LR_TBSL).
      IF LR_TBSL->KOART = 'D'.
        DATA(LV_CUSTOMERKOART) = 'X'.
      ENDIF.
    ENDLOOP.

    IF NOT LV_CUSTOMERKOART = 'X'.
      EV_CLRST = 'not_relevant'.
    ENDIF.

  ELSE.
    EV_CLRST = 'not_relevant'.
  ENDIF.

  CLEAR XACCIT_KUNNR.
  REFRESH XACCIT_KUNNR.
  CLEAR LV_CUSTOMERKOART.
  CLEAR LV_KUNNR.

ENDFORM.
