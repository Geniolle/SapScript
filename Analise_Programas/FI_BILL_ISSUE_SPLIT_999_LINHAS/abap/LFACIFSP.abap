*----------------------------------------------------------------------*
***INCLUDE LFACIFSP.
*----------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&      Form DOCUMENT_SPLIT_GOODS_MOVEMENT
*&---------------------------------------------------------------------*
*       Split FI documents from goods movement postings (-> reference
*       procedure 'MKPF') if they exceed the maximum of 999 line items
*       per FI document and if no summarization in FI is active.
*
*       As this split logic takes care that the cut is not done within
*       logical transactions, which stand for single goods movements,
*       the balance of the split FI documents is 0 without the need to
*       create technical split clearing line items: MM-IM provides a
*       balance of 0 in transaction currency and first local currency
*       for each logical transaction (-> field LOGVO), and with note
*       1295856 also FI takes care of that for a possible second and
*       third local currency.
*----------------------------------------------------------------------*

FORM DOCUMENT_SPLIT_GOODS_MOVEMENT.

  DATA: BEGIN OF LT_DOCUMENTS OCCURS 1,
          DOCUMENT_NUMBER      TYPE BELNR_D,
          NUMBER_OF_LINE_ITEMS TYPE I,
          TABIX_FROM           TYPE SYTABIX,
          TABIX_TO             TYPE SYTABIX,
        END OF LT_DOCUMENTS,
        LV_BELNR                 TYPE BELNR_D VALUE 1,
        LV_BUZEI                 TYPE I,
        LV_TABIX_FROM_LAST_LOGVO TYPE SYTABIX,
        LV_PS_US_ACTIVE.

* only for goods movement postings
  CHECK ACCHD_FI-AWTYP = 'MKPF'
   OR ( ( ACCHD_FI-AWTYP = 'CF3P' OR ACCHD_FI-AWTYP = 'CF3PG' ) AND
          ACCHD_FI-GLVOR(2) = 'RM' AND ACCHD_FI-GLVOR <> 'RMRP' ).

* check whether the complete posting consists of more than
* 999 line items
  DESCRIBE TABLE ACCIT_FI LINES SY-TFILL.
  CHECK SY-TFILL > 999.

* do not perform this kind of split if summarization in FI is active
* Begin of note 1779136
  PERFORM READ_TTYPV.

* begin of note 2400530
  DESCRIBE TABLE LOGDN LINES SY-TFILL.

  IF SY-TFILL <> 0.
    LOOP AT LOGDN.
      LOOP AT GT_TTYPV WHERE AWTYP = 'MKPF'
                       AND ( BUKRS = LOGDN-BUKRS
                       OR    BUKRS = SPACE )
                       AND ( BLART = LOGDN-BLART
                       OR    BLART = SPACE ).
        EXIT.
      ENDLOOP.
      IF SY-SUBRC IS INITIAL.
        RETURN.
      ENDIF.
    ENDLOOP.
  ELSE.
    CLEAR: SAVE.
    LOOP AT ACCIT_FI.
      IF ACCIT_FI-BUKRS <> SAVE-BUKRS
      OR ACCIT_FI-BLART <> SAVE-BLART.
        SAVE-BUKRS = ACCIT_FI-BUKRS.
        SAVE-BLART = ACCIT_FI-BLART.
        LOOP AT GT_TTYPV WHERE AWTYP = 'MKPF'
                         AND ( BUKRS = SAVE-BUKRS
                         OR    BUKRS = SPACE )
                         AND ( BLART = SAVE-BLART
                         OR    BLART = SPACE ).
          EXIT.
        ENDLOOP.
        IF SY-SUBRC IS INITIAL.
          RETURN.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDIF.
* end of note 2400530
* End of note 1779136

* do not perform this kind of split if US Federal Government ISPS
* is active -> different split for JFMIP
  CALL FUNCTION 'FM00_CHECK_ISPS'
    EXPORTING
      I_CLIENT       = SY-MANDT
    IMPORTING
      E_PS_US_ACTIVE = LV_PS_US_ACTIVE.
  CHECK LV_PS_US_ACTIVE IS INITIAL.

  SORT ACCIT_FI BY BELNR LOGVO POSNR.

* determine the rough structure of the posting by determining the to be
* created FI documents and the numer of their line items
  READ TABLE ACCIT_FI INDEX 1.
  SAVE-BELNR = ACCIT_FI-BELNR.

  LOOP AT ACCIT_FI.
    IF ACCIT_FI-BELNR NE SAVE-BELNR.
      LT_DOCUMENTS-DOCUMENT_NUMBER = SAVE-BELNR.
      LT_DOCUMENTS-TABIX_FROM =
        SY-TABIX - LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS.
      LT_DOCUMENTS-TABIX_TO = SY-TABIX - 1.
      SAVE-BELNR = ACCIT_FI-BELNR.
      APPEND LT_DOCUMENTS.
      CLEAR LT_DOCUMENTS.
    ENDIF.
    LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS =
      LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS + 1.
  ENDLOOP.
  LT_DOCUMENTS-DOCUMENT_NUMBER = ACCIT_FI-BELNR.
  DESCRIBE TABLE ACCIT_FI LINES SY-TFILL.
  LT_DOCUMENTS-TABIX_FROM =
    SY-TFILL - LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS + 1.
  LT_DOCUMENTS-TABIX_TO = SY-TFILL.
  APPEND LT_DOCUMENTS.

* check whether a split is necessary at all
  LOOP AT LT_DOCUMENTS TRANSPORTING NO FIELDS
    WHERE NUMBER_OF_LINE_ITEMS > 999.
    EXIT.
  ENDLOOP.
  CHECK SY-SUBRC EQ 0.

* execute the split
  LOOP AT LT_DOCUMENTS.
    IF LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS > 999.
      READ TABLE ACCIT_FI INDEX LT_DOCUMENTS-TABIX_FROM.
      SAVE-LOGVO = ACCIT_FI-LOGVO.
      LOOP AT ACCIT_FI FROM LT_DOCUMENTS-TABIX_FROM
                         TO LT_DOCUMENTS-TABIX_TO.
        LV_BUZEI = LV_BUZEI + 1.
        IF ACCIT_FI-LOGVO NE SAVE-LOGVO.
          IF LV_BUZEI > 999.
            LV_BELNR = LV_BELNR + 1.
            LV_BUZEI = SY-TABIX - LV_TABIX_FROM_LAST_LOGVO + 1.
          ENDIF.
          ACCIT_FI-BELNR    = LV_BELNR.
          ACCIT_FI-BELNR(2) = '$$'.
          ACCIT_FI-XSPLIT = 'X'.
          MODIFY ACCIT_FI FROM ACCIT_FI TRANSPORTING BELNR XSPLIT
            WHERE LOGVO = SAVE-LOGVO
              AND BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
          SAVE-LOGVO = ACCIT_FI-LOGVO.
          LV_TABIX_FROM_LAST_LOGVO = SY-TABIX.
        ENDIF.
      ENDLOOP.
      IF LV_BUZEI > 999.
        LV_BELNR = LV_BELNR + 1.
      ENDIF.
      ACCIT_FI-BELNR    = LV_BELNR.
      ACCIT_FI-BELNR(2) = '$$'.
      ACCIT_FI-XSPLIT = 'X'.
      MODIFY ACCIT_FI FROM ACCIT_FI TRANSPORTING BELNR XSPLIT
        WHERE LOGVO = SAVE-LOGVO
          AND BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
    ELSE.
      ACCIT_FI-BELNR    = LV_BELNR.
      ACCIT_FI-BELNR(2) = '$$'.
      ACCIT_FI-XSPLIT = 'X'.
      MODIFY ACCIT_FI FROM ACCIT_FI TRANSPORTING BELNR XSPLIT
        WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
    ENDIF.
    LV_BELNR = LV_BELNR + 1.
    LV_BUZEI = 0.
  ENDLOOP.

* check the balance of all to be created FI documents
  CLEAR PRUEF.
  PERFORM CHECK_BALANCE.

ENDFORM.                    " DOCUMENT_SPLIT_GOODS_MOVEMENT

*&---------------------------------------------------------------------*
*&      Form DOCUMENT_SPLIT_DPR
*&---------------------------------------------------------------------*
*       split of down payment requests
*----------------------------------------------------------------------*

FORM DOCUMENT_SPLIT_DPR.

  DATA: LV_BADI_BILL_ISSUE_SPLIT
          TYPE REF TO FI_BILL_ISSUE_SPLIT,
        LV_IS_PS_NOT_ACTIVE,
        LV_PS_US_ACTIVE,
        LS_T001                    TYPE T001,
        LV_INTCA                   TYPE INTCA,
        LS_T003                    TYPE T003,
        LS_NRIV                    TYPE NRIV,
        LT_ACCIT_FI                TYPE TABLE OF ACCIT_FI,
        LT_ACCCR_FI                TYPE TABLE OF ACCCR_FI,
        LT_BSET                    TYPE BSET_TAB,
        LV_AUTOMATIC_SPLIT         TYPE XFELD,
        LV_NUMBER_OF_INVOICE_ITEMS TYPE I,
        LV_COUNT                   TYPE I,
        LV_BELNR                   TYPE BELNR_D VALUE 1,
        BEGIN OF LT_BELNR_TAB OCCURS 1,
          TAX_COUNTRY TYPE FOT_TAX_COUNTRY,
          MWSKZ       TYPE MWSKZ,
          TXDAT_FROM  TYPE FOT_TXDAT_FROM,
          VBEL2       TYPE VBELN_VA,
          POSN2       TYPE POSNR_VA,
          BELNR       TYPE BELNR_D,
        END OF LT_BELNR_TAB,
        LV_SAVE_VBEL2       TYPE VBELN_VA,
        LV_SAVE_POSN2       TYPE POSNR_VA,
        LV_SAVE_TAX_COUNTRY TYPE FOT_TAX_COUNTRY,
        LV_SAVE_MWSKZ       TYPE MWSKZ,
        LV_SAVE_TXDAT_FROM  TYPE FOT_TXDAT_FROM,
        LV_SAVE_TABIX       TYPE SYTABIX.

* execute this split only for down payment requests from SD billing
  CHECK ACCHD_FI-AWTYP = 'VBRK' AND ACCHD_FI-GLVOR = 'RFST'.
  READ TABLE ACCIT_FI INDEX 1.
  CHECK ACCIT_FI-BSTAT EQ CHAR_S.

* this document split in FI is only executed if it has been activated
* via a corresponding BAdI implementation
  TRY.
      GET BADI LV_BADI_BILL_ISSUE_SPLIT.
    CATCH CX_BADI_NOT_IMPLEMENTED.
      CLEAR LV_BADI_BILL_ISSUE_SPLIT.
  ENDTRY.
  CHECK LV_BADI_BILL_ISSUE_SPLIT IS BOUND.

* do not perform this kind of split if US Federal Government ISPS
* is active -> different split for JFMIP
* and not for funds management either
  CALL FUNCTION 'FM00_CHECK_ISPS'
    EXPORTING
      I_CLIENT              = SY-MANDT
    IMPORTING
      E_FLG_ISPS_NOT_ACTIVE = LV_IS_PS_NOT_ACTIVE
      E_PS_US_ACTIVE        = LV_PS_US_ACTIVE.

  CHECK LV_PS_US_ACTIVE IS INITIAL AND NOT
        LV_IS_PS_NOT_ACTIVE IS INITIAL.

  SAVE-BELNR = ACCIT_FI-BELNR.
  SAVE-BLART = ACCIT_FI-BLART.
  SAVE-BUKRS = ACCIT_FI-BUKRS.
  LOOP AT ACCIT_FI WHERE BELNR NE SAVE-BELNR
                      OR BLART NE SAVE-BLART
                      OR BUKRS NE SAVE-BUKRS.
    EXIT.
  ENDLOOP.

  CHECK SY-SUBRC NE 0.

  LOOP AT ACCIT_FI WHERE KOART NE CHAR_D
                     AND TAXIT NE CHAR_X.
    EXIT.
  ENDLOOP.

  CHECK SY-SUBRC NE 0.

  CALL FUNCTION 'FI_COMPANY_CODE_DATA'
    EXPORTING
      I_BUKRS = ACCIT_FI-BUKRS
    IMPORTING
      E_T001  = LS_T001.

  CALL FUNCTION 'COUNTRY_CODE_SAP_TO_ISO'
    EXPORTING
      SAP_CODE = LS_T001-LAND1
    IMPORTING
      ISO_CODE = LV_INTCA.

* do not execute this split if Joint-Venture-Accounting (JVA) is active
* in the company code
  CHECK LS_T001-XJVAA IS INITIAL.

* This kind of split is only allowed for the countries listed
* below. It can be extended to other countries as well after all
* implications on country-specific (VAT) reporting have been
* clarified and resolved by the corresponding country versions.
  CHECK LV_INTCA EQ 'US' OR LV_INTCA EQ 'CA' OR
        LV_INTCA EQ 'MX' OR LV_INTCA EQ 'BM' OR
        LV_INTCA EQ 'DE' OR LV_INTCA EQ 'AT' OR
        LV_INTCA EQ 'CH' OR LV_INTCA EQ 'DK' OR
        LV_INTCA EQ 'SE' OR LV_INTCA EQ 'FI' OR
        LV_INTCA EQ 'GB' OR LV_INTCA EQ 'FR' OR
        LV_INTCA EQ 'NL' OR LV_INTCA EQ 'BE' OR
        LV_INTCA EQ 'LU' OR
*        LV_INTCA EQ 'ES' OR LV_INTCA EQ 'PT' OR
        LV_INTCA EQ 'LT' OR
*        LV_INTCA EQ 'RO' OR
        LV_INTCA EQ 'AU' OR LV_INTCA EQ 'NZ' OR
        LV_INTCA EQ 'IE' OR LV_INTCA EQ 'ZA' OR
        LV_INTCA EQ 'SA' OR LV_INTCA EQ 'AE' OR
        LV_INTCA EQ 'IN' OR
        LV_INTCA EQ 'SG' OR
        LV_INTCA EQ 'TH' OR
        LV_INTCA EQ 'MA' OR LV_INTCA EQ 'TN' OR
        LV_INTCA EQ 'AO'.

  CALL FUNCTION 'FI_DOCUMENT_TYPE_DATA'
    EXPORTING
      I_BLART = ACCIT_FI-BLART
    IMPORTING
      E_T003  = LS_T003.

  CALL FUNCTION 'NUMBER_GET_INFO'
    EXPORTING
      NR_RANGE_NR = LS_T003-NUMKR
      OBJECT      = 'RF_BELEG'
      SUBOBJECT   = ACCIT_FI-BUKRS
      TOYEAR      = ACCIT_FI-GJAHR
    IMPORTING
      INTERVAL    = LS_NRIV.

* do not split in case of external document number assignment in FI
  CHECK LS_NRIV-EXTERNIND IS INITIAL.

  LT_ACCIT_FI[] = ACCIT_FI[].
  LT_ACCCR_FI[] = ACCCR_FI[].
  LT_BSET[] = XBSET[].

  CALL BADI LV_BADI_BILL_ISSUE_SPLIT->ACTIVATE_AUTOMATIC_SPLIT
    EXPORTING
      I_ACCHD_FI        = ACCHD_FI
      IT_ACCIT_FI       = LT_ACCIT_FI
      IT_ACCCR_FI       = LT_ACCCR_FI
      IT_BSET           = LT_BSET
    IMPORTING
      E_AUTOMATIC_SPLIT = LV_AUTOMATIC_SPLIT.

  CALL BADI LV_BADI_BILL_ISSUE_SPLIT->SET_NUMBER_OF_INVOICE_ITEMS
    EXPORTING
      I_ACCHD_FI                = ACCHD_FI
      IT_ACCIT_FI               = LT_ACCIT_FI
      IT_ACCCR_FI               = LT_ACCCR_FI
      IT_BSET                   = LT_BSET
    IMPORTING
      E_NUMBER_OF_INVOICE_ITEMS = LV_NUMBER_OF_INVOICE_ITEMS.

  REFRESH: LT_ACCIT_FI, LT_ACCCR_FI, LT_BSET.

  LOOP AT ACCIT_FI WHERE KOART EQ CHAR_D.
    LV_COUNT = LV_COUNT + 1.
  ENDLOOP.

* there is no or no valid BAdI implementation or split is not necessary
* as the maximum number of invoice items is not exceeded
  CHECK ( LV_AUTOMATIC_SPLIT = CHAR_X AND
          LV_COUNT > MAX_LINE_ITEMS )
     OR ( LV_NUMBER_OF_INVOICE_ITEMS >= 1 AND
          LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS ).

  LV_COUNT = 0.

  CLEAR LOGDN.
  REFRESH LOGDN.

  SORT ACCIT_FI BY POSNR.

* execute the split for the customer line items
  LOOP AT ACCIT_FI WHERE KOART EQ CHAR_D.
    LV_COUNT = LV_COUNT + 1.
    IF ( LV_AUTOMATIC_SPLIT = CHAR_X AND
         LV_COUNT > MAX_LINE_ITEMS )
    OR ( LV_NUMBER_OF_INVOICE_ITEMS >= 1 AND
         LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS ).
      LV_BELNR = LV_BELNR + 1.
      LV_COUNT = 0.
    ENDIF.
    ACCIT_FI-BELNR    = LV_BELNR.
    ACCIT_FI-BELNR(2) = '$$'.
    ACCIT_FI-XSPLIT = 'X'.
    MODIFY ACCIT_FI TRANSPORTING BELNR XSPLIT.
    MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
    COLLECT LOGDN.
    MOVE-CORRESPONDING ACCIT_FI TO LT_BELNR_TAB.
    COLLECT LT_BELNR_TAB.
  ENDLOOP.

  SORT LT_BELNR_TAB BY VBEL2 POSN2 TAX_COUNTRY MWSKZ TXDAT_FROM.

* consistency check :
* if line items with the same values in fields VBEL2, POSN2 and MWSKZ
* would be put into different FI documents, terminate with an error
  LOOP AT LT_BELNR_TAB.
    IF LV_SAVE_VBEL2 EQ LT_BELNR_TAB-VBEL2 AND
       LV_SAVE_POSN2 EQ LT_BELNR_TAB-POSN2 AND
       LV_SAVE_TAX_COUNTRY = LT_BELNR_TAB-TAX_COUNTRY AND
       LV_SAVE_MWSKZ EQ LT_BELNR_TAB-MWSKZ AND
       LV_SAVE_TXDAT_FROM EQ LT_BELNR_TAB-TXDAT_FROM.
      MESSAGE E898 WITH 'LFACIFSP'.
    ENDIF.
    LV_SAVE_VBEL2 = LT_BELNR_TAB-VBEL2.
    LV_SAVE_POSN2 = LT_BELNR_TAB-POSN2.
    LV_SAVE_TAX_COUNTRY = LT_BELNR_TAB-TAX_COUNTRY.
    LV_SAVE_MWSKZ = LT_BELNR_TAB-MWSKZ.
    LV_SAVE_TXDAT_FROM = LT_BELNR_TAB-TXDAT_FROM.
  ENDLOOP.

* execute the split for the VAT tax line items
  LOOP AT ACCIT_FI WHERE TAXIT = CHAR_X.
    LV_SAVE_TABIX = SY-TABIX.
    READ TABLE LT_BELNR_TAB WITH KEY VBEL2 = ACCIT_FI-VBEL2
                                     POSN2 = ACCIT_FI-POSN2
                                     TAX_COUNTRY = ACCIT_FI-TAX_COUNTRY
                                     MWSKZ = ACCIT_FI-MWSKZ
                                     TXDAT_FROM = ACCIT_FI-TXDAT_FROM
      BINARY SEARCH.
    IF SY-SUBRC NE 0.
      MESSAGE E898 WITH 'LFACIFSP'.
    ENDIF.
    ACCIT_FI-BELNR = LT_BELNR_TAB-BELNR.
    ACCIT_FI-XSPLIT = CHAR_X.
    MODIFY ACCIT_FI INDEX LV_SAVE_TABIX TRANSPORTING BELNR XSPLIT.
  ENDLOOP.

ENDFORM.                    " DOCUMENT_SPLIT_DPR

*&---------------------------------------------------------------------*
*&      Form DOCUMENT_SPLIT_CDPC
*&---------------------------------------------------------------------*
*       split of customer down payment clearings
*----------------------------------------------------------------------*
FORM DOCUMENT_SPLIT_CDPC TABLES   T_ACCIT STRUCTURE ACCIT
                         CHANGING CV_TABIX TYPE SYTABIX.

  DATA: LV_COUNT_DPC TYPE SYTABIX.

  DATA: LV_BLART_DPC TYPE BLART,
        LS_TSTCP     TYPE TSTCP,
        LV_PARAM     TYPE TCDPARAM,
        LS_T003      TYPE T003,
        LV_GJAHR     TYPE GJAHR,
        LS_NRIV      TYPE NRIV.

  DATA: LV_BELNR                 TYPE BELNR_D VALUE 1,
        LV_BUZEI                 TYPE I,
        LV_TABIX_FROM_LAST_LOGVO TYPE SYTABIX.

  FIELD-SYMBOLS: <ACCIT> TYPE ACCIT.

  DESCRIBE TABLE T_ACCIT LINES SY-TFILL.
  CHECK SY-TFILL > CV_TABIX.

  LV_COUNT_DPC = SY-TFILL - CV_TABIX.
  CHECK LV_COUNT_DPC > 980.

* determine the default FI document type
* for customer down payment clearings
  SELECT SINGLE * FROM TSTCP INTO LS_TSTCP
    WHERE TCODE = 'F-39'.

  IF SY-SUBRC = 0.
    LV_PARAM = LS_TSTCP-PARAM.
    SHIFT LV_PARAM UP TO 'BKPF-BLART'.
    IF SY-SUBRC = 0.
      SHIFT LV_PARAM UP TO '='.
      SHIFT LV_PARAM.
      IF SY-SUBRC = 0.
        SHIFT LV_PARAM LEFT DELETING LEADING ' '.
        LV_BLART_DPC = LV_PARAM.
      ENDIF.
    ENDIF.
  ENDIF.

  CHECK NOT LV_BLART_DPC IS INITIAL.

  CALL FUNCTION 'FI_DOCUMENT_TYPE_DATA'
    EXPORTING
      I_BLART = LV_BLART_DPC
    IMPORTING
      E_T003  = LS_T003.

  IF T_ACCIT-POSNR NE '0000000001'.
    READ TABLE T_ACCIT WITH KEY POSNR = '0000000001'.
  ENDIF.

* authority check for the default FI document type used
* for the down payment clearing
  IF LS_T003-BRGRU IS NOT INITIAL.                          "n3198223
    AUTHORITY-CHECK OBJECT 'F_BKPF_BLA'
      ID 'ACTVT' FIELD '01'
      ID 'BRGRU' FIELD LS_T003-BRGRU.

    IF SY-SUBRC NE 0.
      MESSAGE E087 WITH LV_BLART_DPC.
    ENDIF.
  ENDIF.                                                    "n3198223
  CALL FUNCTION 'FI_PERIOD_DETERMINE'
    EXPORTING
      I_BUDAT = T_ACCIT-BUDAT
      I_BUKRS = T_ACCIT-BUKRS
    IMPORTING
      E_GJAHR = LV_GJAHR.

  CALL FUNCTION 'NUMBER_GET_INFO'
    EXPORTING
      NR_RANGE_NR = LS_T003-NUMKR
      OBJECT      = 'RF_BELEG'
      SUBOBJECT   = T_ACCIT-BUKRS
      TOYEAR      = LV_GJAHR
    IMPORTING
      INTERVAL    = LS_NRIV.

  IF LS_NRIV-EXTERNIND EQ CHAR_X.
    MESSAGE E221(F5A) WITH LV_BLART_DPC T_ACCIT-BLART.
  ENDIF.

  CV_TABIX = CV_TABIX + 1.

  LOOP AT T_ACCIT FROM CV_TABIX ASSIGNING <ACCIT>.      "#EC CI_NOORDER
    CLEAR <ACCIT>-BELNR.
    <ACCIT>-BLART = LV_BLART_DPC.
  ENDLOOP.

* execute the split
  READ TABLE T_ACCIT INDEX CV_TABIX.
  SAVE-LOGVO = T_ACCIT-LOGVO.
  LOOP AT T_ACCIT FROM CV_TABIX.                        "#EC CI_NOORDER
    LV_BUZEI = LV_BUZEI + 1.
    IF T_ACCIT-LOGVO NE SAVE-LOGVO.
      IF LV_BUZEI > 999.
        LV_BELNR = LV_BELNR + 1.
        LV_BUZEI = SY-TABIX - LV_TABIX_FROM_LAST_LOGVO + 1.
      ENDIF.
      T_ACCIT-BELNR    = LV_BELNR.
      T_ACCIT-BELNR(3) = '$$$'.
*      T_ACCIT-XSPLIT = 'X'.
      MODIFY T_ACCIT FROM T_ACCIT TRANSPORTING BELNR "XSPLIT
        WHERE LOGVO = SAVE-LOGVO.
      SAVE-LOGVO = T_ACCIT-LOGVO.
      LV_TABIX_FROM_LAST_LOGVO = SY-TABIX.
    ENDIF.
  ENDLOOP.
  IF LV_BUZEI > 999.
    LV_BELNR = LV_BELNR + 1.
  ENDIF.
  T_ACCIT-BELNR    = LV_BELNR.
  T_ACCIT-BELNR(3) = '$$$'.
*  T_ACCIT-XSPLIT = 'X'.
  MODIFY T_ACCIT FROM T_ACCIT TRANSPORTING BELNR "XSPLIT
    WHERE LOGVO = SAVE-LOGVO.

ENDFORM.