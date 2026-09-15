*&---------------------------------------------------------------------*
*&  Include           LFACIFSI
*&---------------------------------------------------------------------*


  TYPES: BEGIN OF ST_INVOICE_ITEMS,
           INVOICE_ITEM            TYPE CK_ZEILE,
           NUMBER_OF_FI_LINE_ITEMS TYPE I,
         END OF ST_INVOICE_ITEMS,
         TT_INVOICE_ITEMS TYPE TABLE OF ST_INVOICE_ITEMS WITH KEY INVOICE_ITEM,
         BEGIN OF ST_TAX_ITEMS,
           TAX_ITEM               TYPE TAX_POSNR,
           NUMBER_OF_BSET_ENTRIES TYPE I,
         END OF ST_TAX_ITEMS,
         TT_TAX_ITEMS TYPE TABLE OF ST_TAX_ITEMS WITH KEY TAX_ITEM,
         BEGIN OF ST_DOC_STRUCTURE,
           DOCUMENT_NUMBER         TYPE BELNR_D,
           NUMBER_OF_LINE_ITEMS    TYPE I,
           NUMBER_OF_BSET_ENTRIES  TYPE I,
           NUMBER_OF_TAX_ITEMS     TYPE I,
           NUMBER_OF_INVOICE_ITEMS TYPE I,
           INVOICE_ITEMS           TYPE TT_INVOICE_ITEMS,
           TAX_ITEMS               TYPE TT_TAX_ITEMS,
           TABIX_FROM              TYPE SYTABIX,
           TABIX_TO                TYPE SYTABIX,
           LOGVO                   TYPE LOGVO,
         END OF ST_DOC_STRUCTURE,
         TT_DOC_STRUCTURE TYPE TABLE OF ST_DOC_STRUCTURE,
         BEGIN OF ST_SPLIT_CLEARING,
           BUKRS TYPE BUKRS,
           BELNR TYPE BELNR_D,
           TAX_COUNTRY TYPE FOT_TAX_COUNTRY,
           MWSKZ TYPE MWSKZ,
           TXDAT_FROM TYPE FOT_TXDAT_FROM,
           TXJCD TYPE TXJCD,
           XSKRL TYPE XSKRL,
           CURTP TYPE CURTP,
           WAERS TYPE WAERS,
           WRBTR TYPE WRBTR,
         END OF ST_SPLIT_CLEARING,
         TT_SPLIT_CLEARING TYPE TABLE OF ST_SPLIT_CLEARING.


*&---------------------------------------------------------------------*
*&
*&      Form SPLIT_INVOICE
*&
*&---------------------------------------------------------------------*
*
*       split FI documents from invoice receipt and bill issue
*       postings with business transactions 'RMRP' and 'SD00'
*
*       - if this document split in FI is activated for
*         invoice receipt postings via an implementation of BAdI
*         'FI_INVOICE_RECEIPT_SPLIT' or for
*         bill issue postings via an implementation of BAdI
*         'FI_BILL_ISSUE_SPLIT',
*       - if no summarization in FI is active,
*       - if Funds Management (PSM-FM) is not active,
*       - if Joint-Venture-Accounting (JVA) is not active and
*       - if the FI documents are not distributed to other systems,
*         where the document split of newGL is active, via IDOC/ALE
*         with message types 'FIDCC1' or 'FIDCC2'
*
*----------------------------------------------------------------------*
  FORM SPLIT_INVOICE.

    DATA: LS_INVOICE_ITEM               TYPE ST_INVOICE_ITEMS,
          LS_TAX_ITEM                   TYPE ST_TAX_ITEMS,
          LT_DOCUMENTS                  TYPE TABLE OF ST_DOC_STRUCTURE WITH HEADER LINE,
          LV_BELNR                      TYPE BELNR_D VALUE 1,
          LX_MAIN_DOC                   TYPE XFELD,                 "note 3269293
          LV_MAIN_BELNR                 TYPE BELNR_D,
          LV_SPLIT_ALLOWED_BY_PS,
          LV_XTXIT                      TYPE XTXIT_TXD,
          LV_SAVE_XTXIT                 TYPE XTXIT_TXD,
          LS_BUKRS                      TYPE FAGL_S_BUKRS,
          LT_BUKRS                      TYPE FAGL_T_BUKRS,
          LV_BLART_SUBSEQ               TYPE BLART,
          LS_T001                       TYPE T001,
          LV_EXTERNIND                  TYPE XFELD,
          LV_INTCA                      TYPE INTCA,
          LV_SAVE_ZEILE                 TYPE CK_ZEILE,
          LV_SAVE_POSNR_SD              TYPE POSNR,
          LV_SAVE_TAXPS                 TYPE TAX_POSNR,
          LV_COUNT                      TYPE I,
          LV_BUZEI                      TYPE BUZEI,
          LV_BUZEI_BSET                 TYPE BUZEI,
          LV_BUZEI_SUM                  TYPE I,
          LV_BUZEI_BSET_SUM             TYPE I,
          LV_SAVE_TABIX                 TYPE SYTABIX,
          LV_BADI_INVOICE_RECEIPT_SPLIT
            TYPE REF TO FI_INVOICE_RECEIPT_SPLIT,
          LV_BADI_BILL_ISSUE_SPLIT
            TYPE REF TO FI_BILL_ISSUE_SPLIT,
          LV_NUMBER_OF_INVOICE_ITEMS    TYPE I,
          LV_AUTOMATIC_SPLIT            TYPE XFELD,
          LT_SPLIT_CLEARING             TYPE TABLE OF ST_SPLIT_CLEARING WITH HEADER LINE,
          LS_ACCIT_FI                   TYPE ACCIT_FI,
          LT_ACCIT_FI                   TYPE TABLE OF ACCIT_FI,
          LS_ACCCR_FI                   TYPE ACCCR_FI,
          LT_ACCCR_FI                   TYPE TABLE OF ACCCR_FI,
          LT_BSET                       TYPE BSET_TAB,
          LV_POSNR                      TYPE POSNR_ACC,
          LV_NO_SPLIT                   TYPE XFELD,
          LV_NO_CUST_VEND               TYPE XFELD,                 "note 3060320
          LV_SKIP_ITEM                  TYPE XFELD.

    DATA: LV_MSG_TYPE TYPE MSGTS,
          LV_ERR_MODE TYPE XFELD.

*only for invoice receipt and bill issue postings
    CHECK ACCHD_FI-AWTYP = 'RMRP'
     OR ( ACCHD_FI-AWTYP = 'BKPFF' AND ACCHD_FI-GLVOR = 'RMRP' )
     OR ( ACCHD_FI-AWTYP = 'BEBD' AND ACCHD_FI-GLVOR = 'RMRP' )     "note 2931261
     OR ( ACCHD_FI-AWTYP = 'WBRK' AND ACCHD_FI-GLVOR = 'RMRP' )
     OR ( ACCHD_FI-AWTYP = 'CF3P' AND ACCHD_FI-GLVOR = 'RMRP' )
     OR ( ACCHD_FI-AWTYP = 'CF3PM' AND ACCHD_FI-GLVOR = 'RMRP' )
     OR ( ACCHD_FI-AWTYP = 'VBRK' AND ACCHD_FI-GLVOR = 'SD00' )
     OR ( ACCHD_FI-AWTYP = 'BKPFF' AND ACCHD_FI-GLVOR = 'SD00' )    "note 3030853
     OR ( ACCHD_FI-AWTYP = 'BEBD' AND ACCHD_FI-GLVOR = 'SD00' )
     OR ( ACCHD_FI-AWTYP = 'WBRK' AND ACCHD_FI-GLVOR = 'SD00' )
     OR ( ACCHD_FI-AWTYP = 'CF3P' AND ACCHD_FI-GLVOR = 'SD00' )
     OR ( ACCHD_FI-AWTYP = 'CF3PS' AND ACCHD_FI-GLVOR = 'SD00' ).

*not for document parking, except from LIV
    CHECK ( ACCHD_FI-STATUS_NEW NE '2' AND ACCHD_FI-STATUS_NEW NE '3' )
       OR ACCHD_FI-AWTYP = 'RMRP'.

*check whether error mode is active
    CALL FUNCTION 'READ_CUSTOMIZED_MESSAGE'
      EXPORTING
        I_ARBGB = 'FACI_ANA'
        I_DTYPE = '-'
        I_MSGNR = '100'
      IMPORTING
        E_MSGTY = LV_MSG_TYPE.

    IF LV_MSG_TYPE <> '-'.
      LV_ERR_MODE = 'X'.
    ENDIF.

*this document split in FI is only executed if it has been activated
*via a corresponding BAdI implementation
    IF ACCHD_FI-GLVOR = 'RMRP'.
      TRY.
          GET BADI LV_BADI_INVOICE_RECEIPT_SPLIT.
        CATCH CX_BADI_NOT_IMPLEMENTED.
          IF LV_ERR_MODE = 'X'.
            MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '101' WITH 'FI_INVOICE_RECEIPT_SPLIT'.
          ENDIF.
          CLEAR LV_BADI_INVOICE_RECEIPT_SPLIT.
      ENDTRY.
      CHECK LV_BADI_INVOICE_RECEIPT_SPLIT IS BOUND.
    ELSEIF ACCHD_FI-GLVOR = 'SD00'.
      TRY.
          GET BADI LV_BADI_BILL_ISSUE_SPLIT.
        CATCH CX_BADI_NOT_IMPLEMENTED.
          IF LV_ERR_MODE = 'X'.
            MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '101' WITH 'FI_BILL_ISSUE_SPLIT'.
          ENDIF.
          CLEAR LV_BADI_BILL_ISSUE_SPLIT.
      ENDTRY.
      CHECK LV_BADI_BILL_ISSUE_SPLIT IS BOUND.
    ENDIF.

* begin of note 2861298
    READ TABLE ACCIT_FI INDEX 1.
    IF NOT ACCIT_FI-BELNR_SENDER IS INITIAL.
* inbound processing in a Central Finance system
* -> check whether the posting was split in the sending system
      LOOP AT ACCIT_FI TRANSPORTING NO FIELDS WHERE KTOSL = 'SPL'.
        EXIT.
      ENDLOOP.
* the posting was split in the sending system
      IF SY-SUBRC IS INITIAL.
* Indicator XSPLIT, which indicates that the FI document results
* from a posting, which was split in FI, is not transferred to a
* Central Finance system. So set it in this Central Finance system
* again so that BKPF-XSPLIT and BSPL are updated there accordingly.
        ACCIT_FI-XSPLIT = CHAR_X.
        MODIFY ACCIT_FI TRANSPORTING XSPLIT WHERE XSPLIT NE CHAR_X.
        EXIT.
      ENDIF.
    ENDIF.
* end of note 2861298

    PERFORM CHECK_ALE_INBOUND_N CHANGING LV_NO_SPLIT.
    IF NOT LV_NO_SPLIT IS INITIAL.                        "note 2693737
      IF LV_ERR_MODE = 'X'.                               "note 2693737
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '102'.      "note 2693737
      ENDIF.                                              "note 2693737
      EXIT.                                               "note 2693737
    ENDIF.                                                "note 2693737

*do not perform this kind of split if summarization in FI is active
    PERFORM IS_SUMMARIZATION_ACTIVE CHANGING LV_NO_SPLIT.
    IF NOT LV_NO_SPLIT IS INITIAL.
      IF LV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '103' WITH ACCHD_FI-AWTYP.
      ENDIF.
      EXIT.
    ENDIF.

*This kind of split is only allowed if EITHER in ALL involved
*company codes tax jurisdictions with line item based VAT tax
*calculation is active OR in NONE of the involved company codes
*tax jurisdictions with line item based VAT tax calculation is
*active so that the integrity of the split FI documents from a
*VAT tax point of view is ensured.
*This kind of split is only allowed for the countries listed
*below. It can be extended to other countries as well after all
*implications on country-specific (VAT) reporting have been
*clarified and resolved by the corresponding country versions.
*This kind of split is not allowed for Joint-Venture-Accounting
*(JVA).
    SORT ACCIT_FI BY BUKRS.

    CLEAR SAVE-BUKRS.

    READ TABLE ACCIT_FI INDEX 1.
    PERFORM IS_XTXIT_ACTIVE USING    ACCHD_FI
                                     ACCIT_FI
                            CHANGING LV_SAVE_XTXIT.

    LOOP AT ACCIT_FI ASSIGNING <ACCIT_FI>.
      CHECK SAVE-BUKRS <> <ACCIT_FI>-BUKRS.
      SAVE-BUKRS = <ACCIT_FI>-BUKRS.
      LS_BUKRS = <ACCIT_FI>-BUKRS.
      COLLECT LS_BUKRS INTO LT_BUKRS.

      CALL FUNCTION 'FM_DOCUMENT_CHECK_FI_SPLIT'
        EXPORTING
          I_BUKRS         = <ACCIT_FI>-BUKRS
        IMPORTING
          E_SPLIT_ALLOWED = LV_SPLIT_ALLOWED_BY_PS.
      IF LV_SPLIT_ALLOWED_BY_PS <> CHAR_X.
        IF LV_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '104'.
        ENDIF.
        LV_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

      PERFORM IS_XTXIT_ACTIVE USING    ACCHD_FI
                                       <ACCIT_FI>
                              CHANGING LV_XTXIT.

      IF LV_XTXIT NE LV_SAVE_XTXIT.
        IF LV_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '105'.
        ENDIF.
        LV_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

      CALL FUNCTION 'FI_COMPANY_CODE_DATA'
        EXPORTING
          I_BUKRS = <ACCIT_FI>-BUKRS
        IMPORTING
          E_T001  = LS_T001.
      IF NOT LS_T001-XJVAA IS INITIAL.
        IF CL_JVA_SWITCH=>GET_INSTANCE( )->JVA_ON_ACDOCA( <ACCIT_FI>-BUDAT ) <> ABAP_TRUE. "< 3397676
          IF LV_ERR_MODE = 'X'.
            MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '106' WITH <ACCIT_FI>-BUKRS.
          ENDIF.
          LV_NO_SPLIT = 'X'.
          EXIT.
        ENDIF.
      ENDIF.

      CALL FUNCTION 'COUNTRY_CODE_SAP_TO_ISO'
        EXPORTING
          SAP_CODE = LS_T001-LAND1
        IMPORTING
          ISO_CODE = LV_INTCA.

      PERFORM IS_COUNTRY_ALLOWED USING LV_INTCA
                                 CHANGING LV_NO_SPLIT.
      IF NOT LV_NO_SPLIT IS INITIAL.
        IF LV_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '107' WITH LV_INTCA.
        ENDIF.
        EXIT.
      ENDIF.
    ENDLOOP.

    CHECK LV_NO_SPLIT IS INITIAL.

    PERFORM CHECK_PARKING_XTXIT CHANGING LV_NO_SPLIT      "note 3041292
                                         LV_XTXIT.        "note 3041292

    CHECK LV_NO_SPLIT IS INITIAL.                         "note 3041292

*do not perform this kind of split if the FI documents are distributed
*to other systems, where the document split of newGL is active, via
*IDOC/ALE with message types 'FIDCC1' or 'FIDCC2'
    PERFORM CHECK_ALE_OUTBOUND_N USING LT_BUKRS           "note 2825257
                                 CHANGING LV_NO_SPLIT.    "note 2825257
    IF NOT LV_NO_SPLIT IS INITIAL.
      IF LV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '102'.
      ENDIF.
      EXIT.
    ENDIF.

    PERFORM IS_DOC_NUMBER_EXTERN CHANGING LV_EXTERNIND.
    PERFORM CHECK_DOC_TYPE_SUBSEQ USING    LV_BADI_INVOICE_RECEIPT_SPLIT
                                           LV_BADI_BILL_ISSUE_SPLIT
                                           LV_EXTERNIND
                                           LV_INTCA
                                           LT_BUKRS
                                           LV_ERR_MODE
                                  CHANGING LV_BLART_SUBSEQ
                                           LV_NO_SPLIT.
    CHECK LV_NO_SPLIT IS INITIAL.

    PERFORM CHECK_DOC_STRUCTURE TABLES LT_BUKRS
                                USING  LV_XTXIT
                                       LV_ERR_MODE
                                CHANGING LV_NO_SPLIT
                                         LV_NO_CUST_VEND. "note 3060320

    CHECK LV_NO_SPLIT IS INITIAL.

    PERFORM DET_SPLIT_DOC_STRUCTURE USING    LV_XTXIT
                                             LT_BUKRS[]
                                             LV_ERR_MODE
                                    CHANGING LT_DOCUMENTS[]
                                             LV_NO_SPLIT.
    CHECK LV_NO_SPLIT IS INITIAL.

    LT_ACCIT_FI[] = ACCIT_FI[].
    LT_ACCCR_FI[] = ACCCR_FI[].
    LT_BSET[] = XBSET[].

*let the system automatically determine the maximum
*number of invoice items per split FI document and
*execute the split accordingly ?
    IF ACCHD_FI-GLVOR = 'RMRP'.
      CALL BADI LV_BADI_INVOICE_RECEIPT_SPLIT->ACTIVATE_AUTOMATIC_SPLIT
        EXPORTING
          I_ACCHD_FI        = ACCHD_FI
          IT_ACCIT_FI       = LT_ACCIT_FI
          IT_ACCCR_FI       = LT_ACCCR_FI
          IT_BSET           = LT_BSET
        IMPORTING
          E_AUTOMATIC_SPLIT = LV_AUTOMATIC_SPLIT.
    ELSEIF ACCHD_FI-GLVOR = 'SD00'.
      CALL BADI LV_BADI_BILL_ISSUE_SPLIT->ACTIVATE_AUTOMATIC_SPLIT
        EXPORTING
          I_ACCHD_FI        = ACCHD_FI
          IT_ACCIT_FI       = LT_ACCIT_FI
          IT_ACCCR_FI       = LT_ACCCR_FI
          IT_BSET           = LT_BSET
        IMPORTING
          E_AUTOMATIC_SPLIT = LV_AUTOMATIC_SPLIT.
    ENDIF.

*the number of invoice items in a split FI document can be set
*via a BAdI implementation
    IF ACCHD_FI-GLVOR = 'RMRP'.
      CALL BADI LV_BADI_INVOICE_RECEIPT_SPLIT->SET_NUMBER_OF_INVOICE_ITEMS
        EXPORTING
          I_ACCHD_FI                = ACCHD_FI
          IT_ACCIT_FI               = LT_ACCIT_FI
          IT_ACCCR_FI               = LT_ACCCR_FI
          IT_BSET                   = LT_BSET
        IMPORTING
          E_NUMBER_OF_INVOICE_ITEMS = LV_NUMBER_OF_INVOICE_ITEMS.
    ELSEIF ACCHD_FI-GLVOR = 'SD00'.
      CALL BADI LV_BADI_BILL_ISSUE_SPLIT->SET_NUMBER_OF_INVOICE_ITEMS
        EXPORTING
          I_ACCHD_FI                = ACCHD_FI
          IT_ACCIT_FI               = LT_ACCIT_FI
          IT_ACCCR_FI               = LT_ACCCR_FI
          IT_BSET                   = LT_BSET
        IMPORTING
          E_NUMBER_OF_INVOICE_ITEMS = LV_NUMBER_OF_INVOICE_ITEMS.
    ENDIF.

    REFRESH: LT_ACCIT_FI, LT_ACCCR_FI, LT_BSET.

    IF LV_ERR_MODE = 'X'
    AND LV_AUTOMATIC_SPLIT IS INITIAL
    AND LV_NUMBER_OF_INVOICE_ITEMS IS INITIAL.
      MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '108'.
    ENDIF.

*check whether a split is necessary at all
    IF LV_AUTOMATIC_SPLIT = CHAR_X.
      LOOP AT LT_DOCUMENTS TRANSPORTING NO FIELDS
        WHERE NUMBER_OF_LINE_ITEMS > MAX_LINE_ITEMS
           OR NUMBER_OF_BSET_ENTRIES > MAX_LINE_ITEMS.
        EXIT.
      ENDLOOP.
    ELSEIF LV_NUMBER_OF_INVOICE_ITEMS >= 1.
      LOOP AT LT_DOCUMENTS TRANSPORTING NO FIELDS
        WHERE NUMBER_OF_INVOICE_ITEMS > LV_NUMBER_OF_INVOICE_ITEMS
           OR NUMBER_OF_TAX_ITEMS > LV_NUMBER_OF_INVOICE_ITEMS.
        EXIT.
      ENDLOOP.
    ENDIF.

*there is no or no valid BAdI implementation or no split is necessary
*as the maximum number of invoice items is not exceeded
    IF SY-SUBRC NE 0 OR ( LV_NUMBER_OF_INVOICE_ITEMS < 1 AND
                          LV_AUTOMATIC_SPLIT NE CHAR_X )
    OR ( LV_XTXIT IS INITIAL AND
         ACCHD_FI-GLVOR EQ 'RMRP' AND
         LV_AUTOMATIC_SPLIT NE CHAR_X AND
         LV_NUMBER_OF_INVOICE_ITEMS < 167 ).
      IF NOT LV_XTXIT IS INITIAL.
        PERFORM INIT_TAXPS.
      ENDIF.
      EXIT.
    ENDIF.

*Fields MWSKZ or MWSK1/DMBT1, MWSK2/DMBT2 and MWSK3/DMBT3 need to be
*populated in the vendor/customer line item(s) so that the logic of
*FORMs DOCUMENT_MW2TAB/DOCUMENT_MW2TAB_SUBST is needed here.
*But FORMs DOCUMENT_MW2TAB/DOCUMENT_MW2TAB_SUBST should not
*be processed for such split postings.

    IF ACCHD_FI-STATUS_NEW NE '2'.
      PERFORM VAT_BREAKDOWN_N.
    ENDIF.

    CLEAR LOGDN.
    REFRESH LOGDN.

*execute the split

    LOOP AT LT_DOCUMENTS.

      IF ( ( ( LT_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS > LV_NUMBER_OF_INVOICE_ITEMS OR
               LT_DOCUMENTS-NUMBER_OF_TAX_ITEMS > LV_NUMBER_OF_INVOICE_ITEMS ) AND
             LV_NUMBER_OF_INVOICE_ITEMS >= 1 )
        OR ( ( LT_DOCUMENTS-NUMBER_OF_LINE_ITEMS > MAX_LINE_ITEMS OR
               LT_DOCUMENTS-NUMBER_OF_BSET_ENTRIES > MAX_LINE_ITEMS ) AND
             LV_AUTOMATIC_SPLIT = CHAR_X ) )
      AND ( ( ACCHD_FI-AWTYP EQ 'RMRP' AND
            ( LT_DOCUMENTS-LOGVO EQ 'MAIN' OR
              LT_DOCUMENTS-LOGVO EQ SPACE  OR             "note 3440208
              LT_DOCUMENTS-LOGVO EQ 'VAL'  OR
              LT_DOCUMENTS-LOGVO EQ 'PUR'  OR             "note 3206651
              LT_DOCUMENTS-LOGVO EQ 'RET' ) ) OR
          ( ( ACCHD_FI-AWTYP EQ 'BKPFF' OR
              ACCHD_FI-AWTYP EQ 'WBRK'  OR
              ACCHD_FI-AWTYP EQ 'VBRK'  OR
              ACCHD_FI-AWTYP EQ 'BEBD'  OR
              ACCHD_FI-AWTYP EQ 'CF3P'  OR
              ACCHD_FI-AWTYP EQ 'CF3PM' OR
              ACCHD_FI-AWTYP EQ 'CF3PS') AND
              LT_DOCUMENTS-LOGVO IS INITIAL ) ).

        IF LV_XTXIT IS INITIAL.
*put the vendor/customer line item(s) and the VAT tax line items
*into the first split FI document
          LOOP AT ACCIT_FI FROM LT_DOCUMENTS-TABIX_FROM
                             TO LT_DOCUMENTS-TABIX_TO
            WHERE ( KOART CA 'DVK' AND
                    KTOSL NE 'EGX' ) OR NOT
                  MWART IS INITIAL OR
*the withholding tax line items, too
                  KTOSL EQ 'WIT' OR
                  KTOSL EQ 'OFF' OR
                  KTOSL EQ 'GRU' OR
*put all automatic summary line items into the first
*split FI document as well
*('KDM' and 'PRD' line items are always on item level)
                  KTOSL EQ 'RKA' OR
                  KTOSL EQ 'DIF' OR
                  KTOSL EQ 'UPF' OR
                  KTOSL EQ 'VVA' OR
                  KTOSL EQ 'MVA' OR
                  KTOSL EQ 'KDT' OR
                  KTOSL EQ 'KDF' OR
                  KTOSL EQ 'RDF' OR
*the cash discount clearing line items as well
                  KTOSL EQ 'SKV' OR
*and the clearing line item in case of prepayment processing as well
                  KTOSL EQ 'PPX' OR                       "note 3477973
*and the credit card line items as well
                  CCINS NE SPACE OR
*and the line item with the cash account in case of cash sale as well
                ( POSNR EQ '0000000001' AND               "note 2978879
                  KOART EQ 'S' AND                        "note 2978879
                  AWTYP EQ 'VBRK' ) OR                    "note 2978879
                ( POSNR_SD IS INITIAL AND                 "note 3060320
                  KOART EQ 'S' AND                        "note 3060320
                  AWTYP EQ 'WBRK' ).                      "note 3060320
            IF ACCIT_FI-POSNR_SD IS INITIAL AND           "note 3060320
               ACCIT_FI-MWART IS INITIAL AND NOT          "note 3070418
             ( ACCIT_FI-KTOSL EQ 'WIT' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'OFF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'GRU' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'RKA' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'DIF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'UPF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'VVA' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'MVA' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'KDT' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'KDF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'RDF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'SKV' OR                 "note 3476873
               ACCIT_FI-CCINS NE SPACE ) AND              "note 3476873
               ACCIT_FI-KOART EQ 'S' AND                  "note 3060320
               ACCIT_FI-AWTYP EQ 'WBRK'.                  "note 3060320
              CHECK ACCHD_FI-GLVOR = 'SD00'               "note 3060320
              AND ( LV_NO_CUST_VEND = CHAR_X OR           "note 3060320
                  ( ACCIT_FI-XBILK IS NOT INITIAL AND     "note 3077047
                    ACCIT_FI-COCO_NUM IS NOT INITIAL ) ). "note 3077047
            ENDIF.                                        "note 3060320
            LX_MAIN_DOC = CHAR_X.                         "note 3269293
            IF LV_EXTERNIND IS INITIAL
            OR ACCIT_FI-LOGVO EQ 'RET'.
              ACCIT_FI-BELNR    = LV_BELNR.
              ACCIT_FI-BELNR(2) = '$$'.
            ENDIF.
            ACCIT_FI-XSPLIT = 'X'.
            MODIFY ACCIT_FI TRANSPORTING BELNR XSPLIT.
            MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
            COLLECT LOGDN.
          ENDLOOP.
          IF LX_MAIN_DOC IS INITIAL.                      "note 3354728
            DESCRIBE TABLE XBSET LINES SY-TFILL.          "note 3354728
            IF SY-TFILL NE 0.                             "note 3354728
              READ TABLE ACCIT_FI                         "note 3354728
                INDEX LT_DOCUMENTS-TABIX_FROM.            "note 3354728
              MOVE-CORRESPONDING ACCIT_FI TO LOGDN.       "note 3354728
              IF LV_EXTERNIND IS INITIAL.                 "note 3354728
                LOGDN-BELNR = LV_BELNR.                   "note 3354728
                LOGDN-BELNR(2) = '$$'.                    "note 3354728
              ENDIF.                                      "note 3354728
              COLLECT LOGDN.                              "note 3354728
              LX_MAIN_DOC = CHAR_X.                       "note 3354728
            ENDIF.                                        "note 3354728
          ENDIF.                                          "note 3354728
          IF LX_MAIN_DOC = CHAR_X.                        "note 3269293
            IF LV_EXTERNIND IS INITIAL
            OR ( ACCIT_FI-AWTYP EQ 'RMRP' AND             "note 3449394
                 ACCIT_FI-LOGVO NE 'MAIN' ).              "note 3449394
              LV_MAIN_BELNR = LV_BELNR.
              LV_MAIN_BELNR(2) = '$$'.
              LV_BELNR = LV_BELNR + 1.
            ELSE.
              LV_MAIN_BELNR = ACCIT_FI-BELNR.
            ENDIF.
          ENDIF.
          LOOP AT XBSET
            WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
            IF LV_EXTERNIND IS INITIAL
            OR ( ACCIT_FI-AWTYP EQ 'RMRP' AND             "note 3449394
                 ACCIT_FI-LOGVO NE 'MAIN' ).              "note 3449394
              XBSET-BELNR = LV_BELNR - 1.
              XBSET-BELNR(2) = '$$'.
            ELSE.
              XBSET-BELNR = ACCIT_FI-BELNR.
            ENDIF.
            MODIFY XBSET TRANSPORTING BELNR.
          ENDLOOP.
          IF SY-SUBRC IS INITIAL AND LV_MAIN_BELNR IS INITIAL.
            MESSAGE E855.
          ENDIF.
*put the G/L line items into the subsequent split FI documents
          LOOP AT ACCIT_FI FROM LT_DOCUMENTS-TABIX_FROM
                             TO LT_DOCUMENTS-TABIX_TO
            WHERE ( NOT KOART CA 'DVK' OR
                        KTOSL EQ 'EGX' ) AND
                      MWART IS INITIAL AND NOT
*put withholding tax line items into the first split FI document
                    ( KTOSL EQ 'WIT' OR
                      KTOSL EQ 'OFF' OR
                      KTOSL EQ 'GRU' OR
*put all automatic summary line items into the first split FI document
*('KDM' and 'PRD' line items are always on item level)
                      KTOSL EQ 'RKA' OR
                      KTOSL EQ 'DIF' OR
                      KTOSL EQ 'UPF' OR
                      KTOSL EQ 'VVA' OR
                      KTOSL EQ 'MVA' OR
                      KTOSL EQ 'KDT' OR
                      KTOSL EQ 'KDF' OR
                      KTOSL EQ 'RDF' OR
*the cash discount clearing line items as well
                      KTOSL EQ 'SKV' OR
*and the clearing line item in case of prepayment processing as well
                      KTOSL EQ 'PPX' OR                   "note 3477973
*and the credit card line items as well
                      CCINS NE SPACE OR
*and the line item with the cash account in case of cash sale as well
                    ( POSNR EQ '0000000001' AND           "note 2978879
                      KOART EQ 'S' AND                    "note 2978879
                      AWTYP EQ 'VBRK' ) ).                "note 2978879
             CHECK NOT ( ACCHD_FI-GLVOR = 'SD00' AND      "note 3060320
                       ( LV_NO_CUST_VEND = CHAR_X OR      "note 3060320
* begin of note 3077047
                       ( ACCIT_FI-XBILK IS NOT INITIAL AND
                         ACCIT_FI-COCO_NUM IS NOT INITIAL ) ) AND
* end of note 3077047
                         ACCIT_FI-POSNR_SD IS INITIAL AND "note 3060320
                         ACCIT_FI-KOART EQ 'S' AND        "note 3060320
                         ACCIT_FI-AWTYP EQ 'WBRK' ).      "note 3060320
            LV_SAVE_TABIX = SY-TABIX.
            IF ( ACCHD_FI-GLVOR = 'RMRP' AND
                 ACCIT_FI-ZEILE NE LV_SAVE_ZEILE AND NOT
               ( ACCIT_FI-ZEILE IS INITIAL OR
                 ACCIT_FI-ZEILE EQ '999999' ) )
            OR ( ACCHD_FI-GLVOR = 'SD00' AND
                 ACCIT_FI-POSNR_SD NE LV_SAVE_POSNR_SD AND NOT
               ( ACCIT_FI-POSNR_SD IS INITIAL OR
                 ACCIT_FI-POSNR_SD EQ '999999' ) ).
              LV_COUNT = LV_COUNT + 1.
              LV_SAVE_ZEILE = ACCIT_FI-ZEILE.
              LV_SAVE_POSNR_SD = ACCIT_FI-POSNR_SD.
              IF LV_AUTOMATIC_SPLIT = CHAR_X.
                IF ACCHD_FI-GLVOR = 'RMRP'.
                  READ TABLE LT_DOCUMENTS-INVOICE_ITEMS
                    BINARY SEARCH
                    WITH KEY INVOICE_ITEM = ACCIT_FI-ZEILE
                    INTO LS_INVOICE_ITEM.
                ELSEIF ACCHD_FI-GLVOR = 'SD00'.
                  READ TABLE LT_DOCUMENTS-INVOICE_ITEMS
                    BINARY SEARCH
                    WITH KEY INVOICE_ITEM = ACCIT_FI-POSNR_SD
                    INTO LS_INVOICE_ITEM.
                ENDIF.
                LV_BUZEI_SUM = LV_BUZEI
                + LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS.
                IF LV_BUZEI_SUM > MAX_LINE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_COUNT = 1.
                ENDIF.
              ELSE.
                IF LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_COUNT = 1.
                ENDIF.
              ENDIF.
            ENDIF.
            LV_BUZEI = LV_BUZEI + 1.
            ACCIT_FI-BELNR    = LV_BELNR.
            ACCIT_FI-BELNR(2) = '$$'.
            ACCIT_FI-XSPLIT = 'X'.
            IF ( NOT LV_EXTERNIND IS INITIAL OR
                 LV_INTCA EQ 'IT' OR
                 LV_INTCA EQ 'ES' OR
                 LV_INTCA EQ 'PT' OR
             ( ( LV_INTCA EQ 'RO' OR
                 LV_INTCA EQ 'LT' OR
                 LV_INTCA EQ 'EE' ) AND NOT
                 LV_BLART_SUBSEQ IS INITIAL ) )
            AND NOT ( ACCHD_FI-AWTYP EQ 'RMRP' AND        "note 3449394
                      LT_DOCUMENTS-LOGVO NE 'MAIN' ).     "note 3449394
              ACCIT_FI-BLART = LV_BLART_SUBSEQ.
            ENDIF.
            MODIFY ACCIT_FI FROM ACCIT_FI INDEX LV_SAVE_TABIX
              TRANSPORTING BELNR BLART XSPLIT.
            IF ACCIT_FI-BUKRS <> X001-BUKRS.
              PERFORM GET_CURRENCY USING    ACCIT_FI-BUKRS
                                   CHANGING X001.
            ENDIF.
            MOVE-CORRESPONDING ACCIT_FI TO LT_SPLIT_CLEARING.
            MOVE-CORRESPONDING ACCIT_FI TO ACCIT_KEY.
            READ TABLE ACCCR_FI BINARY SEARCH WITH KEY ACCIT_KEY.
            TABIX = SY-TABIX.
            LV_SKIP_ITEM = 'X'.
            LOOP AT ACCCR_FI FROM TABIX.
              IF ACCCR_FI-AWTYP <> ACCIT_FI-AWTYP
              OR ACCCR_FI-AWREF <> ACCIT_FI-AWREF
              OR ACCCR_FI-AWORG <> ACCIT_FI-AWORG
              OR ACCCR_FI-POSNR <> ACCIT_FI-POSNR.
                EXIT.
              ENDIF.
              MOVE-CORRESPONDING ACCCR_FI TO LT_SPLIT_CLEARING.
              LT_SPLIT_CLEARING-BELNR = ACCIT_FI-BELNR.
              COLLECT LT_SPLIT_CLEARING.
              IF NOT LV_MAIN_BELNR IS INITIAL.
                LT_SPLIT_CLEARING-BELNR = LV_MAIN_BELNR.
                LT_SPLIT_CLEARING-WRBTR
                  = - LT_SPLIT_CLEARING-WRBTR.
                COLLECT LT_SPLIT_CLEARING.
              ENDIF.
              IF ( ACCCR_FI-CURTP = '00'
              OR   ACCCR_FI-CURTP = '10'
              OR   ACCCR_FI-CURTP = X001-CURT2
              OR   ACCCR_FI-CURTP = X001-CURT3 )
              AND  ACCCR_FI-WRBTR <> 0.
                CLEAR: LV_SKIP_ITEM.
              ENDIF.
            ENDLOOP.
            IF NOT LV_SKIP_ITEM IS INITIAL.
              LV_BUZEI = LV_BUZEI - 1.
            ENDIF.
            MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
            COLLECT LOGDN.
          ENDLOOP.

        ELSE.

*put the vendor/customer line item(s) into the first split
*FI document
          LOOP AT ACCIT_FI FROM LT_DOCUMENTS-TABIX_FROM
                             TO LT_DOCUMENTS-TABIX_TO
            WHERE ( KOART CA 'DVK' AND KTOSL NE 'BUV'
                                   AND KTOSL NE 'EGX' ) OR
*put down payment clearing(s) into the first split FI document
                  DOCCAT EQ 'DPC_SD' OR                   "note 3074492
*put withholding tax line items into the first split FI document
                  KTOSL EQ 'WIT' OR
                  KTOSL EQ 'OFF' OR
                  KTOSL EQ 'GRU' OR
*put all automatic summary line items into the first
*split FI document as well
*('KDM' and 'PRD' line items are always on item level)
                  KTOSL EQ 'UPF' OR                       "note 3476873
                  KTOSL EQ 'KDT' OR                       "note 3476873
                  KTOSL EQ 'KDF' OR                       "note 3476873
                  KTOSL EQ 'RDF' OR                       "note 3476873
*put cash discount clearing line items into the first split FI document
                  KTOSL EQ 'SKV' OR
*and the clearing line item in case of prepayment processing as well
                  KTOSL EQ 'PPX' OR                       "note 3477973
*and the credit card line items as well
                  CCINS NE SPACE OR
*and the line item with the cash account in case of cash sale as well
                ( POSNR EQ '0000000001' AND               "note 2978879
                  KOART EQ 'S' AND                        "note 2978879
                  AWTYP EQ 'VBRK' ) OR                    "note 2978879
                ( POSNR_SD IS INITIAL AND                 "note 3060320
                  MWART IS INITIAL AND                    "note 3070418
                  KOART EQ 'S' AND                        "note 3060320
                  AWTYP EQ 'WBRK' ).                      "note 3060320
            IF ACCIT_FI-POSNR_SD IS INITIAL AND NOT       "note 3060320
             ( ACCIT_FI-KTOSL EQ 'WIT' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'OFF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'GRU' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'UPF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'KDT' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'KDF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'RDF' OR                 "note 3476873
               ACCIT_FI-KTOSL EQ 'SKV' OR                 "note 3476873
               ACCIT_FI-CCINS NE SPACE ) AND              "note 3476873
               ACCIT_FI-KOART EQ 'S' AND                  "note 3060320
               ACCIT_FI-AWTYP EQ 'WBRK'.                  "note 3060320
              CHECK ACCHD_FI-GLVOR = 'SD00'               "note 3060320
              AND ( LV_NO_CUST_VEND = CHAR_X OR           "note 3060320
                  ( ACCIT_FI-XBILK IS NOT INITIAL AND     "note 3077047
                    ACCIT_FI-COCO_NUM IS NOT INITIAL ) ). "note 3077047
            ENDIF.                                        "note 3060320
            LX_MAIN_DOC = CHAR_X.                         "note 3269293
            IF LV_EXTERNIND IS INITIAL
            OR ( ACCIT_FI-AWTYP EQ 'RMRP' AND             "note 3440208
                 ACCIT_FI-LOGVO NE 'MAIN' ).              "note 3314435
              ACCIT_FI-BELNR    = LV_BELNR.
              ACCIT_FI-BELNR(2) = '$$'.
            ENDIF.
            ACCIT_FI-XSPLIT = 'X'.
            MODIFY ACCIT_FI TRANSPORTING BELNR XSPLIT.
            MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
            COLLECT LOGDN.
          ENDLOOP.
          IF LX_MAIN_DOC = CHAR_X.                        "note 3269293
            IF LV_EXTERNIND IS INITIAL
            OR ( ACCIT_FI-AWTYP EQ 'RMRP' AND             "note 3440208
                 ACCIT_FI-LOGVO NE 'MAIN' ).              "note 3314435
              LV_BELNR = LV_BELNR + 1.
            ENDIF.
          ENDIF.
*put the G/L line items and tax line items into the subsequent
*split FI documents
          LOOP AT ACCIT_FI FROM LT_DOCUMENTS-TABIX_FROM
                             TO LT_DOCUMENTS-TABIX_TO
            WHERE ( NOT KOART CA 'DVK' OR KTOSL EQ 'BUV'
                                       OR KTOSL EQ 'EGX' ) AND
*put down payment clearing(s) into the first split FI document
              NOT ( DOCCAT EQ 'DPC_SD' OR                 "note 3074492
*put withholding tax line items into the first split FI document
                    KTOSL EQ 'WIT' OR
                    KTOSL EQ 'OFF' OR
                    KTOSL EQ 'GRU' OR
*put all automatic summary line items into the first
*split FI document as well
*('KDM' and 'PRD' line items are always on item level)
                    KTOSL EQ 'UPF' OR                     "note 3476873
                    KTOSL EQ 'KDT' OR                     "note 3476873
                    KTOSL EQ 'KDF' OR                     "note 3476873
                    KTOSL EQ 'RDF' OR                     "note 3476873
*put cash discount clearing line items into the first split FI document
                    KTOSL EQ 'SKV' OR
*and the clearing line item in case of prepayment processing as well
                    KTOSL EQ 'PPX' OR                     "note 3477973
*and the credit card line items as well
                    CCINS NE SPACE OR
*and the line item with the cash account in case of cash sale as well
                  ( POSNR EQ '0000000001' AND             "note 2978879
                    KOART EQ 'S' AND                      "note 2978879
                    AWTYP EQ 'VBRK' ) ).                  "note 2978879
             CHECK NOT ( ACCHD_FI-GLVOR = 'SD00' AND      "note 3060320
                       ( LV_NO_CUST_VEND = CHAR_X OR      "note 3060320
* begin of note 3077047
                       ( ACCIT_FI-XBILK IS NOT INITIAL AND
                         ACCIT_FI-COCO_NUM IS NOT INITIAL ) ) AND
* end of note 3077047
                         ACCIT_FI-POSNR_SD IS INITIAL AND "note 3060320
                         ACCIT_FI-MWART IS INITIAL AND    "note 3070418
                         ACCIT_FI-KOART EQ 'S' AND        "note 3060320
                         ACCIT_FI-AWTYP EQ 'WBRK' ).      "note 3060320
            LV_SAVE_TABIX = SY-TABIX.
            LOOP AT XBSET WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER
                            AND TAXPS >= LV_SAVE_TAXPS
                            AND TAXPS < ACCIT_FI-TAXPS.
              IF XBSET-TAXPS NE LV_SAVE_TAXPS.
                LV_COUNT = LV_COUNT + 1.
                LV_SAVE_TAXPS = XBSET-TAXPS.
                IF LV_AUTOMATIC_SPLIT = CHAR_X.
                  READ TABLE LT_DOCUMENTS-TAX_ITEMS
                    BINARY SEARCH
                    WITH KEY TAX_ITEM = XBSET-TAXPS
                    INTO LS_TAX_ITEM.
                  LV_BUZEI_BSET_SUM = LV_BUZEI_BSET
                  + LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES.
                  IF LV_BUZEI_BSET_SUM >  MAX_LINE_ITEMS.
                    LV_BELNR = LV_BELNR + 1.
                    LV_BUZEI = 0.
                    LV_BUZEI_BSET = 0.
                    LV_COUNT = 1.
                  ENDIF.
                ELSE.
                  IF LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS.
                    LV_BELNR = LV_BELNR + 1.
                    LV_BUZEI = 0.
                    LV_BUZEI_BSET = 0.
                    LV_COUNT = 1.
                  ENDIF.
                ENDIF.
                IF LV_COUNT = 1.
                  GT_BSET_ONLY-BELNR = LV_BELNR.
                  GT_BSET_ONLY-BELNR(2) = '$$'.
                  GT_BSET_ONLY-BUKRS = XBSET-BUKRS.
                  GT_BSET_ONLY-GJAHR = XBSET-GJAHR.
                  COLLECT GT_BSET_ONLY.
                ENDIF.
              ENDIF.
              LV_BUZEI_BSET = LV_BUZEI_BSET + 1.
              XBSET-BELNR = LV_BELNR.
              XBSET-BELNR(2) = '$$'.
              MODIFY XBSET TRANSPORTING BELNR.
            ENDLOOP.
            IF ACCIT_FI-TAXPS NE LV_SAVE_TAXPS AND NOT
             ( ACCIT_FI-TAXPS IS INITIAL OR ACCIT_FI-TAXPS EQ '999999' ).
              LV_COUNT = LV_COUNT + 1.
              LV_SAVE_TAXPS = ACCIT_FI-TAXPS.
              IF LV_AUTOMATIC_SPLIT = CHAR_X.
                READ TABLE LT_DOCUMENTS-INVOICE_ITEMS
                  BINARY SEARCH
                  WITH KEY INVOICE_ITEM = ACCIT_FI-TAXPS
                  INTO LS_INVOICE_ITEM.
                LV_BUZEI_SUM = LV_BUZEI
                + LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS.
                IF LV_BUZEI_SUM > MAX_LINE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_BUZEI_BSET = 0.
                  LV_COUNT = 1.
                ENDIF.
                READ TABLE LT_DOCUMENTS-TAX_ITEMS
                  BINARY SEARCH
                  WITH KEY TAX_ITEM = ACCIT_FI-TAXPS
                  INTO LS_TAX_ITEM.
                LV_BUZEI_BSET_SUM = LV_BUZEI_BSET
                +  LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES.
                IF LV_BUZEI_BSET_SUM > MAX_LINE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_BUZEI_BSET = 0.
                  LV_COUNT = 1.
                ENDIF.
              ELSE.
                IF LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_BUZEI_BSET = 0.
                  LV_COUNT = 1.
                ENDIF.
              ENDIF.
            ENDIF.
            LV_BUZEI = LV_BUZEI + 1.
            ACCIT_FI-BELNR    = LV_BELNR.
            ACCIT_FI-BELNR(2) = '$$'.
            ACCIT_FI-XSPLIT = 'X'.
            IF ( NOT LV_EXTERNIND IS INITIAL OR
                 LV_INTCA EQ 'IT' OR
                 LV_INTCA EQ 'ES' OR
                 LV_INTCA EQ 'PT' OR
             ( ( LV_INTCA EQ 'RO' OR
                 LV_INTCA EQ 'LT' OR
                 LV_INTCA EQ 'EE' ) AND NOT
                 LV_BLART_SUBSEQ IS INITIAL ) )
            AND NOT ( ACCHD_FI-AWTYP EQ 'RMRP' AND
                      LT_DOCUMENTS-LOGVO NE 'MAIN' ).     "note 3314435
              ACCIT_FI-BLART = LV_BLART_SUBSEQ.
            ENDIF.
            MODIFY ACCIT_FI FROM ACCIT_FI INDEX LV_SAVE_TABIX
              TRANSPORTING BELNR BLART XSPLIT.
            MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
            COLLECT LOGDN.
          ENDLOOP.
          LOOP AT XBSET WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER
                          AND TAXPS >= LV_SAVE_TAXPS.
            IF XBSET-TAXPS NE LV_SAVE_TAXPS.
              LV_COUNT = LV_COUNT + 1.
              LV_SAVE_TAXPS = XBSET-TAXPS.
              IF LV_AUTOMATIC_SPLIT = CHAR_X.
                READ TABLE LT_DOCUMENTS-TAX_ITEMS
                  BINARY SEARCH
                  WITH KEY TAX_ITEM = XBSET-TAXPS
                  INTO LS_TAX_ITEM.
                LV_BUZEI_BSET_SUM = LV_BUZEI_BSET
                +  LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES.
                IF LV_BUZEI_BSET_SUM > MAX_LINE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_BUZEI_BSET = 0.
                  LV_COUNT = 1.
                ENDIF.
              ELSE.
                IF LV_COUNT > LV_NUMBER_OF_INVOICE_ITEMS.
                  LV_BELNR = LV_BELNR + 1.
                  LV_BUZEI = 0.
                  LV_BUZEI_BSET = 0.
                  LV_COUNT = 1.
                ENDIF.
              ENDIF.
              IF LV_COUNT = 1.
                GT_BSET_ONLY-BELNR = LV_BELNR.
                GT_BSET_ONLY-BELNR(2) = '$$'.
                GT_BSET_ONLY-BUKRS = XBSET-BUKRS.
                GT_BSET_ONLY-GJAHR = XBSET-GJAHR.
                COLLECT GT_BSET_ONLY.
              ENDIF.
            ENDIF.
            LV_BUZEI_BSET = LV_BUZEI_BSET + 1.
            XBSET-BELNR = LV_BELNR.
            XBSET-BELNR(2) = '$$'.
            MODIFY XBSET TRANSPORTING BELNR.
          ENDLOOP.

        ENDIF.

      ELSE.

        READ TABLE ACCIT_FI
          WITH KEY BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
        ACCIT_FI-BELNR    = LV_BELNR.
        ACCIT_FI-BELNR(2) = '$$'.
        ACCIT_FI-XSPLIT = 'X'.
        MODIFY ACCIT_FI FROM ACCIT_FI TRANSPORTING BELNR XSPLIT
          WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
        XBSET-BELNR    = LV_BELNR.
        XBSET-BELNR(2) = '$$'.
        MODIFY XBSET FROM XBSET TRANSPORTING BELNR
          WHERE BELNR = LT_DOCUMENTS-DOCUMENT_NUMBER.
        MOVE-CORRESPONDING ACCIT_FI TO LOGDN.
        COLLECT LOGDN.

      ENDIF.

      LV_BELNR = LV_BELNR + 1.
      LV_COUNT = 0.
      LV_BUZEI = 0.
      LV_BUZEI_BSET = 0.
      CLEAR LV_SAVE_ZEILE.
      CLEAR LV_SAVE_POSNR_SD.
      CLEAR LV_SAVE_TAXPS.
      CLEAR LX_MAIN_DOC.                                  "note 3269293
      CLEAR LV_MAIN_BELNR.

    ENDLOOP.

    IF NOT LV_XTXIT IS INITIAL.
      PERFORM INIT_TAXPS.
*determine the balances of all split FI documents on the basis
*of which the split clearing line items are created
      PERFORM DETER_BALANCE USING FALSE.
      LOOP AT IBALTAB.
        MOVE-CORRESPONDING IBALTAB TO LT_SPLIT_CLEARING.
        COLLECT LT_SPLIT_CLEARING.
      ENDLOOP.
    ENDIF.

    PERFORM GET_SPLIT_CLEARING TABLES   LT_ACCIT_FI
                                        LT_ACCCR_FI
                               USING    LV_XTXIT
                               CHANGING LT_SPLIT_CLEARING[].

*write back changed FI document type to T_ACCIT
    IF NOT LV_EXTERNIND IS INITIAL
    OR LV_INTCA EQ 'IT'
    OR LV_INTCA EQ 'ES'
    OR LV_INTCA EQ 'PT'
    OR ( ( LV_INTCA EQ 'RO' OR
           LV_INTCA EQ 'LT' OR
           LV_INTCA EQ 'EE' ) AND NOT
         LV_BLART_SUBSEQ IS INITIAL ).
      SORT T_ACCIT BY AWREF AWORG POSNR.
      LOOP AT ACCIT_FI.
        MOVE-CORRESPONDING ACCIT_FI TO ACCIT_KEY.
        READ TABLE T_ACCIT WITH KEY ACCIT_KEY BINARY SEARCH.
        IF SY-SUBRC IS INITIAL.
          T_ACCIT-BLART = ACCIT_FI-BLART.
          MODIFY T_ACCIT INDEX SY-TABIX TRANSPORTING BLART.
        ENDIF.
      ENDLOOP.
    ENDIF.

*insert split clearing line items into ACCIT_FI and ACCCR_FI
    PERFORM INSERT_NEW_ITEMS TABLES LT_ACCIT_FI
                                    LT_ACCCR_FI
                                    ACCIT_FI
                                    ACCCR_FI.

*transfer split clearing line items back to T_ACCIT and T_ACCCR
*for newGL / General Ledger view and for Central Finance (CFIN)
    LOOP AT LT_ACCIT_FI INTO LS_ACCIT_FI.
      MOVE-CORRESPONDING LS_ACCIT_FI TO T_ACCIT.
      APPEND T_ACCIT.
    ENDLOOP.
    LOOP AT LT_ACCCR_FI INTO LS_ACCCR_FI.
      MOVE-CORRESPONDING LS_ACCCR_FI TO T_ACCCR.
      APPEND T_ACCCR.
    ENDLOOP.

    SORT T_ACCIT BY AWTYP AWREF AWORG POSNR.
    SORT T_ACCCR BY AWTYP AWREF AWORG POSNR CURTP.

    CHECK ACCHD_FI-STATUS_NEW NE '2'.

*check the balance of all to be created FI documents
    CLEAR PRUEF.
    PERFORM CHECK_BALANCE.

  ENDFORM.                    " SPLIT_INVOICE

*&---------------------------------------------------------------------*
*&      Form  IS_SUMMARIZATION_ACTIVE
*&---------------------------------------------------------------------*
*       Check whether summarization in FI is active
*----------------------------------------------------------------------*
  FORM IS_SUMMARIZATION_ACTIVE CHANGING C_ACTIVE TYPE XFELD.

    PERFORM READ_TTYPV.
    DESCRIBE TABLE LOGDN LINES SY-TFILL.

    IF SY-TFILL <> 0.
      LOOP AT LOGDN.
        LOOP AT GT_TTYPV WHERE AWTYP = ACCHD_FI-AWTYP
                         AND ( BUKRS = LOGDN-BUKRS
                         OR    BUKRS = SPACE )
                         AND ( BLART = LOGDN-BLART
                         OR    BLART = SPACE ).
          EXIT.
        ENDLOOP.
        IF SY-SUBRC IS INITIAL.
          C_ACTIVE = 'X'.
          EXIT.
        ENDIF.
      ENDLOOP.
    ELSE.
      CLEAR: SAVE.
      LOOP AT ACCIT_FI.
        IF ACCIT_FI-BUKRS <> SAVE-BUKRS
        OR ACCIT_FI-BLART <> SAVE-BLART.
          SAVE-BUKRS = ACCIT_FI-BUKRS.
          SAVE-BLART = ACCIT_FI-BLART.
          LOOP AT GT_TTYPV WHERE AWTYP = ACCHD_FI-AWTYP
                           AND ( BUKRS = SAVE-BUKRS
                           OR    BUKRS = SPACE )
                           AND ( BLART = SAVE-BLART
                           OR    BLART = SPACE ).
            EXIT.
          ENDLOOP.
          IF SY-SUBRC IS INITIAL.
            C_ACTIVE = 'X'.
            EXIT.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDIF.

  ENDFORM.                    " is_summarization_active

*&---------------------------------------------------------------------*
*&      Form  GET_XTXIT
*&---------------------------------------------------------------------*
*       Check whether line by line tax calculation is active
*----------------------------------------------------------------------*
  FORM IS_XTXIT_ACTIVE USING    I_ACCHD_FI TYPE ACCHD_FI
                                I_ACCIT_FI TYPE ACCIT_FI
                       CHANGING C_XTXIT    TYPE XTXIT_TXD.

    DATA: LV_EXTERNAL TYPE XFELD,
          LV_GST_RELE TYPE XFELD,
          LS_BKPF     TYPE BKPF,
          L_TAX_ABROAD_IS_ACTIVE TYPE ABAP_BOOL.

    L_TAX_ABROAD_IS_ACTIVE = CL_FOT_TXA_UTILITIES=>AGENT->IS_TAX_ABROAD_ACTIVE( I_ACCIT_FI-BUKRS ).

    CALL FUNCTION 'CHECK_JURISDICTION_ACTIVE'
      EXPORTING
        I_BUKRS    = I_ACCIT_FI-BUKRS
        I_TAX_ABROAD_ACTIVE = L_TAX_ABROAD_IS_ACTIVE
        I_LAND     = ACCIT_FI-TAX_COUNTRY
      IMPORTING
        E_EXTERNAL = LV_EXTERNAL
        E_XTXIT    = C_XTXIT.

    MOVE-CORRESPONDING I_ACCHD_FI TO LS_BKPF.
    MOVE-CORRESPONDING I_ACCIT_FI TO LS_BKPF.
    CALL FUNCTION 'J_1I4_CALCULATE_TAX_DOCUMENT'
      EXPORTING
        T_BKPF      = LS_BKPF
      CHANGING
        C_FLG_XTXIT = C_XTXIT.
    IF I_ACCHD_FI-AWTYP EQ 'VBRK'.
      IF NOT LV_EXTERNAL IS INITIAL.
        C_XTXIT = LV_EXTERNAL.
      ELSE.
        CALL FUNCTION 'J_1IG_DATE_CHECK'
          IMPORTING
            EX_GST_RELE = LV_GST_RELE.
        IF LV_GST_RELE IS INITIAL.
          CLEAR C_XTXIT.
        ENDIF.
      ENDIF.
    ENDIF.

  ENDFORM.                    " is_xtxit_active

*&---------------------------------------------------------------------*
*&      Form IS_COUNTRY_ALLOWED
*&---------------------------------------------------------------------*
*&   Countries added with notes:
*&   2706392 - Greece (only non-tax invoices)
*&   2733097 - Korea
*&   2754418 - Bulgaria
*&   2928346 - Oman
*&   3001499 - Pakistan
*&   3030401 - Uzbekistan
*&   3030401 - Egypt
*&   3030401 - Jordan
*&   3058944 - Serbia
*&   3133874 - American Virgin Islands
*&   3143475 - Colombia
*&   3163594 - Lichtenstein
*&   3404935 - Dominican Republic
*&   3424602 - Slovakia
*&---------------------------------------------------------------------*
  FORM IS_COUNTRY_ALLOWED USING    IV_INTCA   TYPE INTCA
                          CHANGING C_NO_SPLIT TYPE XFELD.

    IF IV_INTCA NE 'AE' AND IV_INTCA NE 'AO' AND
       IV_INTCA NE 'AR' AND IV_INTCA NE 'AT' AND
       IV_INTCA NE 'AU' AND IV_INTCA NE 'BE' AND
       IV_INTCA NE 'BG' AND IV_INTCA NE 'BM' AND
       IV_INTCA NE 'CA' AND IV_INTCA NE 'CH' AND
       IV_INTCA NE 'CL' AND IV_INTCA NE 'CN' AND
       IV_INTCA NE 'CO' AND IV_INTCA NE 'DE' AND
       IV_INTCA NE 'DK' AND IV_INTCA NE 'DO' AND
       IV_INTCA NE 'EE' AND IV_INTCA NE 'EG' AND
       IV_INTCA NE 'ES' AND IV_INTCA NE 'FI' AND
       IV_INTCA NE 'FR' AND IV_INTCA NE 'GB' AND
       IV_INTCA NE 'HK' AND IV_INTCA NE 'HR' AND
       IV_INTCA NE 'HU' AND IV_INTCA NE 'ID' AND
       IV_INTCA NE 'IE' AND IV_INTCA NE 'IN' AND
       IV_INTCA NE 'IS' AND IV_INTCA NE 'IT' AND
       IV_INTCA NE 'JO' AND IV_INTCA NE 'JP' AND
       IV_INTCA NE 'KR' AND IV_INTCA NE 'KW' AND
       IV_INTCA NE 'LI' AND IV_INTCA NE 'LT' AND
       IV_INTCA NE 'LU' AND IV_INTCA NE 'LV' AND
       IV_INTCA NE 'MA' AND IV_INTCA NE 'MX' AND
       IV_INTCA NE 'MY' AND IV_INTCA NE 'NL' AND
       IV_INTCA NE 'NO' AND IV_INTCA NE 'NZ' AND
       IV_INTCA NE 'OM' AND IV_INTCA NE 'PH' AND
       IV_INTCA NE 'PK' AND IV_INTCA NE 'PL' AND
       IV_INTCA NE 'PR' AND IV_INTCA NE 'PT' AND
       IV_INTCA NE 'QA' AND IV_INTCA NE 'RO' AND
       IV_INTCA NE 'RS' AND IV_INTCA NE 'SA' AND
       IV_INTCA NE 'SE' AND IV_INTCA NE 'SG' AND
       IV_INTCA NE 'SI' AND IV_INTCA NE 'SK' AND
       IV_INTCA NE 'TH' AND IV_INTCA NE 'TN' AND
       IV_INTCA NE 'TR' AND IV_INTCA NE 'TW' AND
       IV_INTCA NE 'US' AND IV_INTCA NE 'UZ' AND
       IV_INTCA NE 'VI' AND IV_INTCA NE 'VN' AND
       IV_INTCA NE 'ZA'


*
* space reserved for SPC notes
*
*
        .
      C_NO_SPLIT = 'X'.
    ENDIF.

* begin of note 2706392
* Allow document split for Greece only if invoice is not tax relevant
    IF IV_INTCA EQ 'GR'.
      READ TABLE ACCIT_FI TRANSPORTING NO FIELDS WITH KEY TAXIT = 'X'.
      IF SY-SUBRC <> 0.
        CLEAR C_NO_SPLIT.
      ENDIF.
    ENDIF.
* end of note 2706392

  ENDFORM.                    " is_country_allowed

*&---------------------------------------------------------------------*
*&      Form IS_DOC_NUMBER_EXTERN
*&---------------------------------------------------------------------*
  FORM IS_DOC_NUMBER_EXTERN CHANGING C_EXTERN TYPE XFELD.

    DATA: LS_T003 TYPE T003,
          LS_NRIV TYPE NRIV.

    SORT ACCIT_FI BY BLART BUKRS.

    CLEAR: SAVE-BLART, SAVE-BUKRS.

* Only the main FI document is considered for external
* document number assignment. For postings from the
* logistical invoice verification the main FI document
* gets LOGVO = 'MAIN'. For other postings the main FI
* document has LOGVO = initial, but for postings from
* the logistical invoice verification the invoice
* reduction FI document gets LOGVO = initial as well.
    LOOP AT ACCIT_FI WHERE ( AWTYP EQ 'RMRP' AND
                             LOGVO EQ 'MAIN' )
                      OR ( ( AWTYP EQ 'BKPFF' OR
                             AWTYP EQ 'WBRK'  OR
                             AWTYP EQ 'VBRK'  OR
                             AWTYP EQ 'BEBD' ) AND
                             LOGVO IS INITIAL ).
      IF SAVE-BLART NE ACCIT_FI-BLART
      OR SAVE-BUKRS NE ACCIT_FI-BUKRS.
        SAVE-BLART = ACCIT_FI-BLART.
        SAVE-BUKRS = ACCIT_FI-BUKRS.
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
        IF NOT LS_NRIV-EXTERNIND IS INITIAL.
          C_EXTERN = 'X'.
          EXIT.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDFORM.                    " is_doc_number_extern

*&---------------------------------------------------------------------*
*&      Form  CHECK_DOC_TYPE_SUBSEQ
*&---------------------------------------------------------------------*
* call BAdI method in order to determine the FI document
* type for the subsequent G/L FI documents if external
* number assignment is used
*----------------------------------------------------------------------*
  FORM CHECK_DOC_TYPE_SUBSEQ  USING    IR_BADI_INVOICE_RECEIPT_SPLIT TYPE REF TO FI_INVOICE_RECEIPT_SPLIT
                                       IR_BADI_BILL_ISSUE_SPLIT TYPE REF TO FI_BILL_ISSUE_SPLIT
                                       IV_EXTERNIND TYPE XFELD
                                       IV_INTCA TYPE INTCA
                                       IT_BUKRS TYPE FAGL_T_BUKRS
                                       IV_ERR_MODE TYPE XFELD
                              CHANGING CV_BLART_SUBSEQ TYPE BLART
                                       CV_NO_SPLIT TYPE XFELD.

    DATA: LT_ACCIT_FI     TYPE TABLE OF ACCIT_FI,
          LT_ACCCR_FI     TYPE TABLE OF ACCCR_FI,
          LT_BSET         TYPE BSET_TAB,
          LS_T003         TYPE T003,
          LS_T003_SUBSEQ  TYPE T003,
          LS_NRIV         TYPE NRIV,
          LS_T8G10        TYPE T8G10,
          LS_T8G12        TYPE T8G12,
          LS_T8G12_SUBSEQ TYPE T8G12.

    CHECK NOT IV_EXTERNIND IS INITIAL
    OR IV_INTCA EQ 'IT'
    OR IV_INTCA EQ 'ES'
    OR IV_INTCA EQ 'PT'
    OR IV_INTCA EQ 'LT'
    OR IV_INTCA EQ 'EE'
    OR IV_INTCA EQ 'RO'.

* cross-company code postings are not supported
    DESCRIBE TABLE IT_BUKRS LINES SY-TFILL.
    IF SY-TFILL <> 1.
      IF IV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '109'.
      ENDIF.
      CV_NO_SPLIT = 'X'.
      EXIT.
    ENDIF.
    LT_ACCIT_FI[] = ACCIT_FI[].
    LT_ACCCR_FI[] = ACCCR_FI[].
    LT_BSET[] = XBSET[].
    IF ACCHD_FI-GLVOR = 'RMRP'.
      CALL BADI IR_BADI_INVOICE_RECEIPT_SPLIT->SET_DOCUMENT_TYPE_SUBSEQ
        EXPORTING
          I_ACCHD_FI             = ACCHD_FI
          IT_ACCIT_FI            = LT_ACCIT_FI
          IT_ACCCR_FI            = LT_ACCCR_FI
          IT_BSET                = LT_BSET
        IMPORTING
          E_DOCUMENT_TYPE_SUBSEQ = CV_BLART_SUBSEQ.
    ELSEIF ACCHD_FI-GLVOR = 'SD00'.
      CALL BADI IR_BADI_BILL_ISSUE_SPLIT->SET_DOCUMENT_TYPE_SUBSEQ
        EXPORTING
          I_ACCHD_FI             = ACCHD_FI
          IT_ACCIT_FI            = LT_ACCIT_FI
          IT_ACCCR_FI            = LT_ACCCR_FI
          IT_BSET                = LT_BSET
        IMPORTING
          E_DOCUMENT_TYPE_SUBSEQ = CV_BLART_SUBSEQ.
    ENDIF.
    REFRESH: LT_ACCIT_FI, LT_ACCCR_FI, LT_BSET.
    IF CV_BLART_SUBSEQ IS INITIAL
    AND NOT ( ( IV_INTCA EQ 'RO' OR
                IV_INTCA EQ 'LT' OR
                IV_INTCA EQ 'EE' ) AND
              ACCHD_FI-GLVOR EQ 'RMRP' ).
      IF IV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '110'.
      ENDIF.
      CV_NO_SPLIT = 'X'.
      EXIT.
    ENDIF.
    IF NOT CV_BLART_SUBSEQ IS INITIAL.
      CALL FUNCTION 'FI_DOCUMENT_TYPE_DATA'
        EXPORTING
          I_BLART = ACCIT_FI-BLART
        IMPORTING
          E_T003  = LS_T003.
      CALL FUNCTION 'FI_DOCUMENT_TYPE_DATA'
        EXPORTING
          I_BLART = CV_BLART_SUBSEQ
        IMPORTING
          E_T003  = LS_T003_SUBSEQ.
      CALL FUNCTION 'NUMBER_GET_INFO'
        EXPORTING
          NR_RANGE_NR = LS_T003_SUBSEQ-NUMKR
          OBJECT      = 'RF_BELEG'
          SUBOBJECT   = ACCIT_FI-BUKRS
          TOYEAR      = ACCIT_FI-GJAHR
        IMPORTING
          INTERVAL    = LS_NRIV.
      IF NOT LS_NRIV-EXTERNIND IS INITIAL.
        IF IV_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '111' WITH CV_BLART_SUBSEQ.
        ENDIF.
        CV_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

* the FI document types must be identical regarding all
* relevant attributes
      CLEAR: LS_T003_SUBSEQ-BLART, LS_T003_SUBSEQ-NUMKR,
             LS_T003_SUBSEQ-STBLA,
             LS_T003_SUBSEQ-XNMRL, LS_T003_SUBSEQ-XAUSG,  "note 3091689
             LS_T003_SUBSEQ-XDTCH, LS_T003_SUBSEQ-BLKLS,  "note 3091689
             LS_T003-BLART, LS_T003-NUMKR, LS_T003-STBLA,
             LS_T003-XNMRL, LS_T003-XAUSG,                "note 3091689
             LS_T003-XDTCH, LS_T003-BLKLS.                "note 3091689
      IF LS_T003_SUBSEQ NE LS_T003.
        IF IV_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '112' WITH ACCIT_FI-BLART CV_BLART_SUBSEQ.
        ENDIF.
        EXIT.
      ENDIF.
      SELECT SINGLE * FROM T8G10 INTO LS_T8G10
        WHERE TCODE = ACCHD_FI-TCODE.
      IF NOT SY-SUBRC IS INITIAL.
        SELECT SINGLE * FROM T8G12 INTO LS_T8G12
          WHERE BLART = ACCIT_FI-BLART.
        SELECT SINGLE * FROM T8G12 INTO LS_T8G12_SUBSEQ
          WHERE BLART = CV_BLART_SUBSEQ.
        IF LS_T8G12-PROCESS NE LS_T8G12_SUBSEQ-PROCESS
        OR LS_T8G12-VARIANT NE LS_T8G12_SUBSEQ-VARIANT.
          IF IV_ERR_MODE = 'X'.
            MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '113' WITH ACCIT_FI-BLART CV_BLART_SUBSEQ.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.

  ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  CHECK_DOC_STRUCTURE
*&---------------------------------------------------------------------*
*       Check the document structure of the to be split invoice
*----------------------------------------------------------------------*
  FORM CHECK_DOC_STRUCTURE TABLES   IT_BUKRS STRUCTURE FAGL_S_BUKRS
                           USING    I_XTXIT    TYPE XFELD
                                    I_ERR_MODE TYPE XFELD
                           CHANGING C_NO_SPLIT TYPE XFELD
                                    C_NO_CUST_VEND TYPE XFELD. "3060320

    LOOP AT ACCIT_FI TRANSPORTING NO FIELDS               "note 3060320
      WHERE KOART CA 'DK'.                                "note 3060320
      EXIT.                                               "note 3060320
    ENDLOOP.                                              "note 3060320
    IF NOT SY-SUBRC IS INITIAL.                           "note 3060320
      C_NO_CUST_VEND = CHAR_X.                            "note 3060320
    ENDIF.                                                "note 3060320

    IF I_XTXIT IS INITIAL.

      LOOP AT XBSET WHERE NOT TAXPS IS INITIAL.
        EXIT.
      ENDLOOP.
      IF SY-SUBRC IS INITIAL.
        IF I_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '114'.
        ENDIF.
        C_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

* cross-company code postings are not supported
      DESCRIBE TABLE IT_BUKRS LINES SY-TFILL.
      IF SY-TFILL <> 1.
        IF I_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '109'.
        ENDIF.
        C_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

      IF ACCHD_FI-GLVOR = 'RMRP'.
        SORT ACCIT_FI BY AWREF AWORG BELNR ZEILE POSNR.
        LOOP AT ACCIT_FI WHERE ( KOART CA 'MSA' OR
                                 KTOSL EQ 'EGX' ) AND
                               MWART IS INITIAL AND NOT
                             ( KTOSL EQ 'WIT' OR
                               KTOSL EQ 'OFF' OR
                               KTOSL EQ 'GRU' OR
                               KTOSL EQ 'RKA' OR
                               KTOSL EQ 'VVA' OR
                               KTOSL EQ 'MVA' OR
                               KTOSL EQ 'DIF' OR
                               KTOSL EQ 'UPF' OR
                               KTOSL EQ 'KDM' OR
                               KTOSL EQ 'KDF' OR
                               KTOSL EQ 'KDT' OR
                               KTOSL EQ 'RDF' OR
                               KTOSL EQ 'SKV' OR
                               KTOSL EQ 'PPX' ) AND       "note 3477973
                               ZEILE IS INITIAL.
          EXIT.
        ENDLOOP.
      ELSEIF ACCHD_FI-GLVOR = 'SD00'.
        SORT ACCIT_FI BY AWREF AWORG BELNR POSNR_SD POSNR.
        LOOP AT ACCIT_FI WHERE KOART CA 'MSA' AND
                               MWART IS INITIAL AND NOT
                             ( KTOSL EQ 'WIT' OR
                               KTOSL EQ 'OFF' OR
                               KTOSL EQ 'GRU' OR
                               KTOSL EQ 'VVA' OR
                               KTOSL EQ 'MVA' OR
                               KTOSL EQ 'KDF' OR
                               KTOSL EQ 'KDT' OR
                               KTOSL EQ 'RDF' ) AND
                               CCINS IS INITIAL AND NOT   "note 2978879
                             ( POSNR EQ '0000000001' AND  "note 2978879
                               KOART EQ 'S' AND           "note 2978879
                               AWTYP EQ 'VBRK' ) AND NOT  "note 2978879
                             ( XBILK IS NOT INITIAL AND   "note 3077047
                               COCO_NUM IS NOT INITIAL AND   "n 3077047
                               KOART EQ 'S' AND           "note 3070418
                               AWTYP EQ 'WBRK' ) AND      "note 3070418
                               POSNR_SD IS INITIAL.
          EXIT.
        ENDLOOP.
      ENDIF.

      IF SY-SUBRC IS INITIAL AND NOT
      ( ACCHD_FI-GLVOR = 'SD00' AND                       "note 3060320
        ACCHD_FI-AWTYP = 'WBRK' AND                       "note 3060320
        C_NO_CUST_VEND = CHAR_X ).                        "note 3060320
        IF I_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '115'.
        ENDIF.
        C_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

      SORT ACCCR_FI BY AWREF AWORG POSNR CURTP.
      SORT XBSET STABLE BY BELNR.

    ELSE.

      IF ACCHD_FI-GLVOR = 'SD00'.
* cross-company code postings are not supported
* for bill issue postings from SD or CRM Billing
        DESCRIBE TABLE IT_BUKRS LINES SY-TFILL.
        IF SY-TFILL <> 1.
          IF I_ERR_MODE = 'X'.
            MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '109'.
          ENDIF.
          C_NO_SPLIT = 'X'.
          EXIT.
        ENDIF.
      ENDIF.

* TAXPS must be set in all tax and tax relevant line items
      LOOP AT ACCIT_FI WHERE KOART CA 'MSA' AND NOT
                           ( KTOSL EQ 'BUV' OR
                             KTOSL EQ 'RKA' OR
                             KTOSL EQ 'DIF' OR
                             KTOSL EQ 'UPF' OR
                             KTOSL EQ 'SKV' OR
                             KTOSL EQ 'PPX' OR            "note 3477973
                             KTOSL EQ 'VVA' OR
                             KTOSL EQ 'MVA' OR
                             KTOSL EQ 'KDT' OR
                             KTOSL EQ 'KDF' OR
                             KTOSL EQ 'RDF' OR
                             KTOSL EQ 'WIT' OR
                             KTOSL EQ 'OFF' OR
                             KTOSL EQ 'GRU' ) AND NOT
                             MWSKZ IS INITIAL AND
                             CCINS IS INITIAL AND NOT     "note 2978879
                           ( POSNR EQ '0000000001' AND    "note 2978879
                             KOART EQ 'S' AND             "note 2978879
                             AWTYP EQ 'VBRK' ) AND        "note 2978879
                             TAXPS IS INITIAL.
        EXIT.
      ENDLOOP.
      IF SY-SUBRC IS INITIAL.
        IF I_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '116'.
        ENDIF.
        C_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

      LOOP AT XBSET WHERE TAXPS IS INITIAL.
        EXIT.
      ENDLOOP.
      IF SY-SUBRC IS INITIAL.
        IF I_ERR_MODE = 'X'.
          MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '116'.
        ENDIF.
        C_NO_SPLIT = 'X'.
        EXIT.
      ENDIF.

* set TAXPS temporarily (-> must be removed again later)
* for G/L line items without tax code so that they are
* not moved to the beginning of the FI document with the
* sort by TAXPS
      LOOP AT ACCIT_FI WHERE ( KOART CA 'MSA' OR
                               KTOSL EQ 'EGX' ) AND
                             TAXIT IS INITIAL AND
                             MWSKZ IS INITIAL AND
                             TAXPS IS INITIAL.
        IF ACCHD_FI-GLVOR = 'RMRP'.
          IF ACCIT_FI-URZEILE IS NOT INITIAL.
            ACCIT_FI-TAXPS = ACCIT_FI-URZEILE.
          ELSEIF ACCIT_FI-ZEILE IS NOT INITIAL.
            ACCIT_FI-TAXPS = ACCIT_FI-ZEILE.
          ENDIF.
        ELSEIF ACCHD_FI-GLVOR = 'SD00'.
          IF ACCIT_FI-POSNR_SD IS NOT INITIAL.
            ACCIT_FI-TAXPS = ACCIT_FI-POSNR_SD.
          ENDIF.
        ENDIF.
        MODIFY ACCIT_FI TRANSPORTING TAXPS.
      ENDLOOP.

* set TAXPS temporarily (-> must be removed again later)
* for automatically created line items so that they are
* not moved from the end of the FI document to its beginning
* with the sort by TAXPS
      LOOP AT ACCIT_FI WHERE ( KTOSL EQ 'BUV' OR
                               KTOSL EQ 'RKA' OR
                               KTOSL EQ 'DIF' OR
                               KTOSL EQ 'UPF' OR
                               KTOSL EQ 'KDT' OR
                               KTOSL EQ 'KDF' OR
                               KTOSL EQ 'RDF' ) AND
                               TAXPS IS INITIAL.
        ACCIT_FI-TAXPS = '999999'.
        MODIFY ACCIT_FI TRANSPORTING TAXPS.
      ENDLOOP.

      SORT ACCIT_FI BY BELNR TAXPS POSNR.
      SORT XBSET STABLE BY BELNR TAXPS.

    ENDIF.

  ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  DET_SPLIT_DOC_STRUCTURE
*&---------------------------------------------------------------------*
* determine the rough structure of the posting by determining the to
* be created FI documents, the numer of their line items, the number
* of their tax line items and the number of their invoice items
*----------------------------------------------------------------------*
  FORM DET_SPLIT_DOC_STRUCTURE USING    IV_XTXIT     TYPE XFELD
                                        IT_BUKRS     TYPE FAGL_T_BUKRS
                                        IV_ERR_MODE  TYPE XFELD
                               CHANGING CT_DOCUMENTS TYPE TT_DOC_STRUCTURE
                                        CV_NO_SPLIT  TYPE XFELD.

    DATA: LS_DOCUMENTS     TYPE ST_DOC_STRUCTURE,
          LV_COUNT         TYPE I,
          LV_SAVE_ZEILE    TYPE CK_ZEILE,
          LV_SAVE_POSNR_SD TYPE POSNR,
          LV_SAVE_TAXPS    TYPE TAX_POSNR,
          LS_INVOICE_ITEM  TYPE ST_INVOICE_ITEMS,
          LS_TAX_ITEM      TYPE ST_TAX_ITEMS,
          LV_SKIP_ITEM     TYPE XFELD,
          LS_PAOBJNR       TYPE RKEOBJNR,
          LT_PAOBJNR       TYPE TABLE OF RKEOBJNR,
          LV_LATEST_BUDAT  TYPE BUDAT.

    FIELD-SYMBOLS: <ACCIT_FI> TYPE ACCIT_FI.

    READ TABLE ACCIT_FI INDEX 1.
    SAVE-BELNR = ACCIT_FI-BELNR.
    LOOP AT ACCIT_FI ASSIGNING <ACCIT_FI>.
      IF <ACCIT_FI>-BELNR NE SAVE-BELNR.
        LS_DOCUMENTS-DOCUMENT_NUMBER = SAVE-BELNR.
        LS_DOCUMENTS-TABIX_FROM = SY-TABIX - LV_COUNT.
        LS_DOCUMENTS-TABIX_TO = SY-TABIX - 1.
        IF ( IV_XTXIT IS INITIAL AND NOT
             LV_SAVE_ZEILE IS INITIAL AND
             ACCHD_FI-GLVOR EQ 'RMRP' )
        OR ( IV_XTXIT IS INITIAL AND NOT
             LV_SAVE_POSNR_SD IS INITIAL AND
             ACCHD_FI-GLVOR EQ 'SD00' )
        OR ( NOT IV_XTXIT IS INITIAL AND NOT
             LV_SAVE_TAXPS IS INITIAL ).
          APPEND LS_INVOICE_ITEM
            TO LS_DOCUMENTS-INVOICE_ITEMS.
        ENDIF.
        CLEAR LV_SAVE_ZEILE.
        CLEAR LV_SAVE_POSNR_SD.
        CLEAR LV_SAVE_TAXPS.
        IF NOT IV_XTXIT IS INITIAL.
          LOOP AT XBSET WHERE BELNR = SAVE-BELNR.
            IF XBSET-TAXPS NE LV_SAVE_TAXPS.
              IF NOT LV_SAVE_TAXPS IS INITIAL.
                APPEND LS_TAX_ITEM
                  TO LS_DOCUMENTS-TAX_ITEMS.
              ENDIF.
              LS_TAX_ITEM-TAX_ITEM = XBSET-TAXPS.
              CLEAR LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES.
              LV_SAVE_TAXPS = XBSET-TAXPS.
              LS_DOCUMENTS-NUMBER_OF_TAX_ITEMS =
                LS_DOCUMENTS-NUMBER_OF_TAX_ITEMS + 1.
            ENDIF.
            LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES =
              LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES + 1.
            LS_DOCUMENTS-NUMBER_OF_BSET_ENTRIES =
              LS_DOCUMENTS-NUMBER_OF_BSET_ENTRIES + 1.
          ENDLOOP.
          IF NOT LV_SAVE_TAXPS IS INITIAL.
            APPEND LS_TAX_ITEM
              TO LS_DOCUMENTS-TAX_ITEMS.
          ENDIF.
        ENDIF.
        SAVE-BELNR = <ACCIT_FI>-BELNR.
        APPEND LS_DOCUMENTS TO CT_DOCUMENTS.
        CLEAR LS_DOCUMENTS.
        CLEAR LV_SAVE_ZEILE.
        CLEAR LV_SAVE_POSNR_SD.
        CLEAR LV_SAVE_TAXPS.
        CLEAR LV_COUNT.
      ENDIF.
      LV_COUNT = LV_COUNT + 1.
      IF IV_XTXIT IS INITIAL.
* Tax items with amount of 0 in all currencies are not copied to
* BSEG, so skip those items
        IF <ACCIT_FI>-BUKRS <> X001-BUKRS.
          PERFORM GET_CURRENCY USING    <ACCIT_FI>-BUKRS
                               CHANGING X001.
        ENDIF.
        MOVE-CORRESPONDING <ACCIT_FI> TO ACCIT_KEY.
        READ TABLE ACCCR_FI BINARY SEARCH
          WITH KEY ACCIT_KEY.
        TABIX = SY-TABIX.
        LV_SKIP_ITEM = 'X'.
        LOOP AT ACCCR_FI FROM TABIX.
          IF ACCCR_FI-AWTYP <> <ACCIT_FI>-AWTYP
          OR ACCCR_FI-AWREF <> <ACCIT_FI>-AWREF
          OR ACCCR_FI-AWORG <> <ACCIT_FI>-AWORG
          OR ACCCR_FI-POSNR <> <ACCIT_FI>-POSNR.
            EXIT.
          ENDIF.
          CHECK ACCCR_FI-CURTP = '00'
          OR    ACCCR_FI-CURTP = '10'
          OR    ACCCR_FI-CURTP = X001-CURT2
          OR    ACCCR_FI-CURTP = X001-CURT3.
          IF ACCCR_FI-WRBTR <> 0.
            CLEAR: LV_SKIP_ITEM.
            EXIT.
          ENDIF.
        ENDLOOP.
        CHECK LV_SKIP_ITEM IS INITIAL.
        IF ACCHD_FI-GLVOR = 'RMRP'.
          IF <ACCIT_FI>-ZEILE NE LV_SAVE_ZEILE AND NOT
           ( <ACCIT_FI>-ZEILE IS INITIAL OR
             <ACCIT_FI>-ZEILE = '999999' ).
            LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS =
              LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS + 1.
            IF NOT LV_SAVE_ZEILE IS INITIAL.
              APPEND LS_INVOICE_ITEM
                TO LS_DOCUMENTS-INVOICE_ITEMS.
            ENDIF.
            LS_INVOICE_ITEM-INVOICE_ITEM = <ACCIT_FI>-ZEILE.
            CLEAR LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS.
            LV_SAVE_ZEILE = <ACCIT_FI>-ZEILE.
          ENDIF.
        ELSEIF ACCHD_FI-GLVOR = 'SD00'.
          IF <ACCIT_FI>-POSNR_SD NE LV_SAVE_POSNR_SD AND NOT
           ( <ACCIT_FI>-POSNR_SD IS INITIAL OR
             <ACCIT_FI>-POSNR_SD = '999999' ).
            LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS =
              LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS + 1.
            IF NOT LV_SAVE_POSNR_SD IS INITIAL.
              APPEND LS_INVOICE_ITEM
                TO LS_DOCUMENTS-INVOICE_ITEMS.
            ENDIF.
            LS_INVOICE_ITEM-INVOICE_ITEM = <ACCIT_FI>-POSNR_SD.
            CLEAR LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS.
            LV_SAVE_POSNR_SD = <ACCIT_FI>-POSNR_SD.
          ENDIF.
        ENDIF.
      ELSE.
        IF <ACCIT_FI>-TAXPS NE LV_SAVE_TAXPS AND NOT
         ( <ACCIT_FI>-TAXPS IS INITIAL OR
           <ACCIT_FI>-TAXPS = '999999' ).
          LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS =
            LS_DOCUMENTS-NUMBER_OF_INVOICE_ITEMS + 1.
          IF NOT LV_SAVE_TAXPS IS INITIAL.
            APPEND LS_INVOICE_ITEM
              TO LS_DOCUMENTS-INVOICE_ITEMS.
          ENDIF.
          LS_INVOICE_ITEM-INVOICE_ITEM = <ACCIT_FI>-TAXPS.
          CLEAR LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS.
          LV_SAVE_TAXPS = <ACCIT_FI>-TAXPS.
        ENDIF.
      ENDIF.
      LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS =
        LS_INVOICE_ITEM-NUMBER_OF_FI_LINE_ITEMS + 1.
      LS_DOCUMENTS-NUMBER_OF_LINE_ITEMS =
        LS_DOCUMENTS-NUMBER_OF_LINE_ITEMS + 1.
      IF <ACCIT_FI>-DOCCAT NE 'DPC_SD'.                   "note 3074492
        LS_DOCUMENTS-LOGVO = <ACCIT_FI>-LOGVO.
      ENDIF.                                              "note 3074492
      IF ( <ACCIT_FI>-PAOBJNR <> IF_FCO_COPA_PAOBJNR=>C_INIT AND <ACCIT_FI>-PAOBJNR <> IF_FCO_COPA_PAOBJNR=>C_ZERO ).
        LS_PAOBJNR = <ACCIT_FI>-PAOBJNR.
        COLLECT LS_PAOBJNR INTO LT_PAOBJNR.
      ENDIF.
    ENDLOOP.
    LS_DOCUMENTS-DOCUMENT_NUMBER = <ACCIT_FI>-BELNR.
    DESCRIBE TABLE ACCIT_FI LINES SY-TFILL.
    LS_DOCUMENTS-TABIX_FROM = SY-TFILL - LV_COUNT + 1.
    LS_DOCUMENTS-TABIX_TO = SY-TFILL.
    IF ( IV_XTXIT IS INITIAL AND NOT
         LV_SAVE_ZEILE IS INITIAL AND
         ACCHD_FI-GLVOR = 'RMRP' )
    OR ( IV_XTXIT IS INITIAL AND NOT
         LV_SAVE_POSNR_SD IS INITIAL AND
         ACCHD_FI-GLVOR = 'SD00' )
    OR ( NOT IV_XTXIT IS INITIAL AND NOT
         LV_SAVE_TAXPS IS INITIAL ).
      APPEND LS_INVOICE_ITEM
        TO LS_DOCUMENTS-INVOICE_ITEMS.
    ENDIF.
    CLEAR LV_SAVE_ZEILE.
    CLEAR LV_SAVE_POSNR_SD.
    CLEAR LV_SAVE_TAXPS.
    CLEAR LV_COUNT.
    IF NOT IV_XTXIT IS INITIAL.
      LOOP AT XBSET WHERE BELNR = <ACCIT_FI>-BELNR.
        IF XBSET-TAXPS NE LV_SAVE_TAXPS.
          IF NOT LV_SAVE_TAXPS IS INITIAL.
            APPEND LS_TAX_ITEM
              TO LS_DOCUMENTS-TAX_ITEMS.
          ENDIF.
          LS_TAX_ITEM-TAX_ITEM = XBSET-TAXPS.
          CLEAR LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES.
          LV_SAVE_TAXPS = XBSET-TAXPS.
          LS_DOCUMENTS-NUMBER_OF_TAX_ITEMS =
            LS_DOCUMENTS-NUMBER_OF_TAX_ITEMS + 1.
        ENDIF.
        LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES =
          LS_TAX_ITEM-NUMBER_OF_BSET_ENTRIES + 1.
        LS_DOCUMENTS-NUMBER_OF_BSET_ENTRIES =
          LS_DOCUMENTS-NUMBER_OF_BSET_ENTRIES + 1.
      ENDLOOP.
      IF NOT LV_SAVE_TAXPS IS INITIAL.
        APPEND LS_TAX_ITEM
          TO LS_DOCUMENTS-TAX_ITEMS.
      ENDIF.
    ENDIF.
    APPEND LS_DOCUMENTS TO CT_DOCUMENTS.
    CLEAR LV_SAVE_ZEILE.
    CLEAR LV_SAVE_POSNR_SD.
    CLEAR LV_SAVE_TAXPS.



* in case of vendor billing documents with AWTYP = 'WBRK'
* do not perform the split of the invoice in FI if the invoice
* contains more than 450/900 different profitability segment numbers
* (field PAOBJNR) and if during later payment cash discount and/or
* exchange rate differences are split according to the profitability
* segment number (field PAOBJNR) as otherwise the maximum number of
* 999 line items in the payment FI document might be exceeded
*
* in case of SD or CRM billing documents do not perform
* the split of the invoice in FI at all if the invoice
* contains profitability segment numbers (field PAOBJNR)
* and if during later payment cash discount and/or
* exchange rate differentces are split according to the
* profitability segment number (field PAOBJNR) as
* otherwise the maximum number of 999 line items in the
* payment FI document might be exceeded

    DATA: LV_MAX_ACCOUNT_ASSIGNMENTS TYPE I VALUE 999999,
          LS_FAGL_S_T8G40_BS         TYPE FAGL_S_T8G40_BS,
          LT_FAGL_TT_T8G40_BS        TYPE FAGL_TT_T8G40_BS.

    CHECK ( ACCHD_FI-AWTYP = 'WBRK' AND ACCHD_FI-GLVOR = 'RMRP' )
    OR    ( ACCHD_FI-AWTYP = 'WBRK' AND ACCHD_FI-GLVOR = 'SD00' )
    OR    ( ACCHD_FI-AWTYP = 'VBRK' AND ACCHD_FI-GLVOR = 'SD00' )
    OR    ( ACCHD_FI-AWTYP = 'BKPFF' AND ACCHD_FI-GLVOR = 'SD00' )  "note 3030853
    OR    ( ACCHD_FI-AWTYP = 'BEBD' AND ACCHD_FI-GLVOR = 'SD00' )
    OR    ( ACCHD_FI-AWTYP = 'BEBD' AND ACCHD_FI-GLVOR = 'RMRP' )   "note 2931261
    OR    ( ACCHD_FI-AWTYP = 'CF3P' AND ACCHD_FI-GLVOR = 'SD00' )
    OR    ( ACCHD_FI-AWTYP = 'CF3PS' AND ACCHD_FI-GLVOR = 'SD00' ).

* determine latest posting date for time-dependent split
    CLEAR LV_LATEST_BUDAT.
    LOOP AT ACCIT_FI ASSIGNING <ACCIT_FI>.
      IF <ACCIT_FI>-BUDAT > LV_LATEST_BUDAT.
        LV_LATEST_BUDAT = <ACCIT_FI>-BUDAT.
      ENDIF.
    ENDLOOP.

    CALL METHOD CL_FAGL_SPLIT_SERVICES=>GET_GL_FIELDS
      EXPORTING
        IT_BUKRS   = IT_BUKRS
        IV_DETAILS = CL_FAGL_SPLIT_SERVICES=>CD_DSC_DETAIL
        IV_DIM     = 'PAOBJNR'
        IV_BUDAT   = LV_LATEST_BUDAT
      IMPORTING
        ET_FIELDS  = LT_FAGL_TT_T8G40_BS.

    IF NOT LT_FAGL_TT_T8G40_BS[] IS INITIAL.
      IF ACCHD_FI-GLVOR = 'SD00'.
        LV_MAX_ACCOUNT_ASSIGNMENTS = 0.
      ELSE.
        LV_MAX_ACCOUNT_ASSIGNMENTS = 900.
      ENDIF.
      REFRESH LT_FAGL_TT_T8G40_BS.
    ENDIF.

    CALL METHOD CL_FAGL_SPLIT_SERVICES=>GET_GL_FIELDS
      EXPORTING
        IT_BUKRS   = IT_BUKRS
        IV_DETAILS = CL_FAGL_SPLIT_SERVICES=>CD_REL_DETAIL
        IV_DIM     = 'PAOBJNR'
        IV_BUDAT   = LV_LATEST_BUDAT
      IMPORTING
        ET_FIELDS  = LT_FAGL_TT_T8G40_BS.

    IF NOT LT_FAGL_TT_T8G40_BS[] IS INITIAL.
      IF ACCHD_FI-GLVOR = 'SD00'.
        LV_MAX_ACCOUNT_ASSIGNMENTS = 0.
      ELSE.
        IF LV_MAX_ACCOUNT_ASSIGNMENTS = 900.
          LV_MAX_ACCOUNT_ASSIGNMENTS = LV_MAX_ACCOUNT_ASSIGNMENTS / 2.
        ELSE.
          LV_MAX_ACCOUNT_ASSIGNMENTS = 900.
        ENDIF.
      ENDIF.
    ENDIF.

    DESCRIBE TABLE LT_PAOBJNR LINES SY-TFILL.
    IF SY-TFILL > LV_MAX_ACCOUNT_ASSIGNMENTS.
      IF NOT IV_XTXIT IS INITIAL.
        PERFORM INIT_TAXPS.
      ENDIF.
      IF IV_ERR_MODE = 'X'.
        MESSAGE ID 'FACI_ANA' TYPE 'E' NUMBER '117'.
      ENDIF.
      CV_NO_SPLIT = 'X'.
      EXIT.
    ENDIF.

  ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  GET_SPLIT_CLEARING
*&---------------------------------------------------------------------*
*       Create the split clearing line items
*----------------------------------------------------------------------*
  FORM GET_SPLIT_CLEARING TABLES CT_ACCIT_FI STRUCTURE ACCIT_FI
                                 CT_ACCCR_FI STRUCTURE ACCCR_FI
                          USING IV_XTXIT TYPE XFELD
                          CHANGING CT_SPLIT_CLEARING TYPE TT_SPLIT_CLEARING.

    DATA: LV_POSNR      TYPE POSNR_ACC,
          LV_BSCHL_H    TYPE BSCHL,
          LV_BSCHL_S    TYPE BSCHL,
          LV_ACC_H      TYPE HKONT,
          LV_ACC_S      TYPE HKONT,
          LV_KTOPL      TYPE KTOPL,
          LV_KOART      TYPE KOART,
          LV_SHKZG      TYPE SHKZG,
          LV_SAKO       TYPE XSAKO,
          LV_SAVE_TABIX TYPE SYTABIX,
          LS_ACCIT_FI   TYPE ACCIT_FI,
          LS_ACCCR_FI   TYPE ACCCR_FI.
    FIELD-SYMBOLS: <SPLIT_CLEARING> TYPE ST_SPLIT_CLEARING.

    LV_POSNR = C_POSNR_SPL.

    CLEAR: SAVE.
    SORT CT_SPLIT_CLEARING STABLE BY BELNR TAX_COUNTRY MWSKZ TXDAT_FROM TXJCD XSKRL CURTP.

    LOOP AT CT_SPLIT_CLEARING ASSIGNING <SPLIT_CLEARING>.
      IF <SPLIT_CLEARING>-BELNR NE SAVE-BELNR
      OR <SPLIT_CLEARING>-TAX_COUNTRY NE SAVE-TAX_COUNTRY
      OR <SPLIT_CLEARING>-MWSKZ NE SAVE-MWSKZ
      OR <SPLIT_CLEARING>-TXDAT_FROM NE SAVE-TXDAT_FROM
      OR <SPLIT_CLEARING>-TXJCD NE SAVE-TXJCD
      OR <SPLIT_CLEARING>-XSKRL NE SAVE-XSKRL.
        SAVE-BELNR = <SPLIT_CLEARING>-BELNR.
        SAVE-TAX_COUNTRY = <SPLIT_CLEARING>-TAX_COUNTRY.
        SAVE-MWSKZ = <SPLIT_CLEARING>-MWSKZ.
        SAVE-TXDAT_FROM = <SPLIT_CLEARING>-TXDAT_FROM.
        SAVE-TXJCD = <SPLIT_CLEARING>-TXJCD.
        SAVE-XSKRL = <SPLIT_CLEARING>-XSKRL.
        READ TABLE ACCIT_FI
          WITH KEY BELNR = <SPLIT_CLEARING>-BELNR.
        CLEAR LS_ACCIT_FI.
        LS_ACCIT_FI-MANDT = ACCHD_FI-MANDT.
        LS_ACCIT_FI-AWTYP = ACCHD_FI-AWTYP.
        LS_ACCIT_FI-AWREF = ACCHD_FI-AWREF.
        LS_ACCIT_FI-AWORG = ACCHD_FI-AWORG.
        LV_POSNR = LV_POSNR + 1.
        LS_ACCIT_FI-POSNR = LV_POSNR.
        LS_ACCIT_FI-LOGVO = ACCIT_FI-LOGVO.
        LS_ACCIT_FI-AWREF_REV = ACCIT_FI-AWREF_REV.
        LS_ACCIT_FI-AWORG_REV = ACCIT_FI-AWORG_REV.
        LS_ACCIT_FI-GJAHR = ACCIT_FI-GJAHR.
        LS_ACCIT_FI-BLDAT = ACCIT_FI-BLDAT.
        LS_ACCIT_FI-BUDAT = ACCIT_FI-BUDAT.
        LS_ACCIT_FI-MONAT = ACCIT_FI-MONAT.
        LS_ACCIT_FI-BLART = ACCIT_FI-BLART.
        LS_ACCIT_FI-XBLNR = ACCIT_FI-XBLNR.
        LS_ACCIT_FI-STBUK = ACCIT_FI-STBUK.
        LS_ACCIT_FI-WWERT = ACCIT_FI-WWERT.
        LS_ACCIT_FI-KTOSL = 'SPL'.
        LS_ACCIT_FI-XSPLIT = CHAR_X.
        LS_ACCIT_FI-KOART = CHAR_S.
        CALL FUNCTION 'FI_STANDARD_ACCOUNT_DETERMINE'
          EXPORTING
            I_BUKRS   = <SPLIT_CLEARING>-BUKRS
            I_KTOSL   = LS_ACCIT_FI-KTOSL
            X_NO_RULE = CHAR_X
          IMPORTING
            E_BSCHH   = LV_BSCHL_H
            E_BSCHS   = LV_BSCHL_S
            E_KONTH   = LV_ACC_H
            E_KONTS   = LV_ACC_S.
* raise error message if account assignment is missing
        IF  LV_ACC_S  IS INITIAL
        OR  LV_ACC_H  IS INITIAL.
          CALL FUNCTION 'FI_CHART_OF_ACCOUNT_DETERMINE'
            EXPORTING
              I_BUKRS = <SPLIT_CLEARING>-BUKRS
            IMPORTING
              E_KTOPL = LV_KTOPL.
          MESSAGE E113 WITH LS_ACCIT_FI-KTOSL ' ' ' ' LV_KTOPL.
        ENDIF.
        IF LV_BSCHL_S IS INITIAL
        OR LV_BSCHL_H IS INITIAL.
          MESSAGE E598 WITH LS_ACCIT_FI-KTOSL.
        ENDIF.
        IF <SPLIT_CLEARING>-WRBTR GT 0.
          LS_ACCIT_FI-SHKZG = CHAR_H.
          LS_ACCIT_FI-HKONT = LV_ACC_H.
          LS_ACCIT_FI-BSCHL = LV_BSCHL_H.
        ELSE.
          LS_ACCIT_FI-SHKZG = CHAR_S.
          LS_ACCIT_FI-HKONT = LV_ACC_S.
          LS_ACCIT_FI-BSCHL = LV_BSCHL_S.
        ENDIF.
        PERFORM SUBST_SINGLE_BSCHL USING LS_ACCIT_FI-BSCHL
                                         ''
                                CHANGING LV_KOART
                                         LV_SHKZG
                                         LS_ACCIT_FI-UMSKS
                                         LS_ACCIT_FI-UMSKZ
                                         LS_ACCIT_FI-XUMSW
                                         LS_ACCIT_FI-XZAHL.
        IF LV_KOART NE LS_ACCIT_FI-KOART.
          MESSAGE E195(F5A) WITH LS_ACCIT_FI-KTOSL.
        ENDIF.
        IF NOT LS_ACCIT_FI-SHKZG IS INITIAL AND
               LS_ACCIT_FI-SHKZG NE LV_SHKZG.
          MESSAGE E846 WITH LS_ACCIT_FI-POSNR
                            LS_ACCIT_FI-SHKZG
                            LS_ACCIT_FI-BSCHL
                            LV_SHKZG.
        ENDIF.
        PERFORM SUBST_SINGLE_HKONT USING <SPLIT_CLEARING>-BUKRS
                                         LS_ACCIT_FI-HKONT
                                CHANGING LV_SAKO.
        LS_ACCIT_FI-LOKKT = LV_SAKO-ALTKT.
        LS_ACCIT_FI-ALTKT = LV_SAKO-BILKT.
        LS_ACCIT_FI-GVTYP = LV_SAKO-GVTYP.
        IF LS_ACCIT_FI-XNCOP IS INITIAL.
          LS_ACCIT_FI-XNCOP = LV_SAKO-XINTB.
        ENDIF.
        IF LS_ACCIT_FI-VBUND IS INITIAL.
          LS_ACCIT_FI-VBUND = LV_SAKO-VBUND.
        ENDIF.
        LS_ACCIT_FI-XBILK = LV_SAKO-XBILK.
        LS_ACCIT_FI-XOPVW = LV_SAKO-XOPVW.
        LS_ACCIT_FI-XKRES = LV_SAKO-XKRES.
        LS_ACCIT_FI-XLGCLR = LV_SAKO-XLGCLR.
        IF LS_ACCIT_FI-XOPVW EQ CHAR_X.
          LS_ACCIT_FI-XKRES = CHAR_X.
        ENDIF.
        LS_ACCIT_FI-XAUTO = CHAR_X.
        MOVE-CORRESPONDING <SPLIT_CLEARING> TO LS_ACCIT_FI.
        LS_ACCIT_FI-ISTAT = '2'.
        LS_ACCIT_FI-VORGN = ACCHD_FI-GLVOR.
        APPEND LS_ACCIT_FI TO CT_ACCIT_FI.
        IF IV_XTXIT IS INITIAL.
          CALL FUNCTION 'FI_TAX_INDICATOR_CHECK'
            EXPORTING
              I_BUKRS = LS_ACCIT_FI-BUKRS
              I_HKONT = LS_ACCIT_FI-HKONT
              I_KOART = LS_ACCIT_FI-KOART
              I_MWSKZ = LS_ACCIT_FI-MWSKZ
              I_STBUK = LS_ACCIT_FI-STBUK
              I_UMSKS = LS_ACCIT_FI-UMSKS
              X_TAXIT = LS_ACCIT_FI-TAXIT
              I_TAX_COUNTRY = LS_ACCIT_FI-TAX_COUNTRY.
        ENDIF.
      ENDIF.
      MOVE-CORRESPONDING <SPLIT_CLEARING> TO LS_ACCCR_FI.
      LS_ACCCR_FI-MANDT = ACCHD_FI-MANDT.
      LS_ACCCR_FI-AWTYP = ACCHD_FI-AWTYP.
      LS_ACCCR_FI-AWREF = ACCHD_FI-AWREF.
      LS_ACCCR_FI-AWORG = ACCHD_FI-AWORG.
      LS_ACCCR_FI-POSNR = LV_POSNR.
      LS_ACCCR_FI-WRBTR = - LS_ACCCR_FI-WRBTR.
      LS_ACCCR_FI-ISTAT = '2'.
      APPEND LS_ACCCR_FI TO CT_ACCCR_FI.
    ENDLOOP.

* split clearing line items might have to be split themselves
* if their amounts have different signs in different currencies
* (improbable, but not impossible, and already happened,
* see below)
    PERFORM CHECK_SHKZG_CONSISTENCY TABLES CT_ACCIT_FI
                                           CT_ACCCR_FI
                                  CHANGING LV_POSNR.

* set PSWSL and PSWBT of the split clearing line items
* after a possible split of split clearing line items as
* split function module 'AC_DOCUMENT_SHKZG_CORRECT' does
* not consider PSWSL and PSWBT
    SORT CT_ACCIT_FI BY AWREF AWORG POSNR.
    SORT CT_ACCCR_FI BY AWREF AWORG POSNR CURTP.
    LOOP AT CT_ACCIT_FI INTO LS_ACCIT_FI.
      LV_SAVE_TABIX = SY-TABIX.
      PERFORM SUBST_SINGLE_HKONT USING LS_ACCIT_FI-BUKRS
                                       LS_ACCIT_FI-HKONT
                              CHANGING LV_SAKO.
      MOVE-CORRESPONDING LS_ACCIT_FI TO ACCCR_KEY.
      IF LV_SAKO-XSALH IS INITIAL.
        ACCCR_KEY-CURTP = '00'.
      ELSEIF NOT LV_SAKO-XSALH IS INITIAL.
        ACCCR_KEY-CURTP = '10'.
      ENDIF.
      READ TABLE CT_ACCCR_FI INTO LS_ACCCR_FI
        WITH KEY ACCCR_KEY BINARY SEARCH.
      IF NOT SY-SUBRC IS INITIAL.
        ACCCR_KEY-CURTP = '10'.
        READ TABLE CT_ACCCR_FI INTO LS_ACCCR_FI
          WITH KEY ACCCR_KEY BINARY SEARCH.
        IF NOT SY-SUBRC IS INITIAL.
          MESSAGE E845 WITH ACCCR_KEY-POSNR ' ' '10'.
        ENDIF.
      ENDIF.
      LS_ACCIT_FI-PSWSL = LS_ACCCR_FI-WAERS.
      LS_ACCIT_FI-PSWBT = LS_ACCCR_FI-WRBTR.
      MODIFY CT_ACCIT_FI INDEX LV_SAVE_TABIX
        FROM LS_ACCIT_FI TRANSPORTING PSWSL PSWBT.
    ENDLOOP.

  ENDFORM.


*&---------------------------------------------------------------------*
*&      Form  CHECK_ALE_INBOUND_N
*&---------------------------------------------------------------------*
  FORM CHECK_ALE_INBOUND_N CHANGING C_NO_SPLIT TYPE XFELD.

    DATA: LS_BUKRS    TYPE FAGL_S_BUKRS,
          LT_BUKRS    TYPE FAGL_T_BUKRS,
          LT_COCD     TYPE FAGL_T_COCD,
          LS_COCD_ACT TYPE FAGL_S_ACT_CC,
          LT_COCD_ACT TYPE FAGL_T_ACT_CC,
          LV_BUDAT    TYPE BUDAT VALUE '99991231'.

    CHECK NOT ALE_FLAG IS INITIAL.
* check whether the posting was split in the sending system
    LOOP AT ACCIT_FI TRANSPORTING NO FIELDS WHERE KTOSL = 'SPL'.
      EXIT.
    ENDLOOP.
* the posting was split in the sending system
    IF SY-SUBRC IS INITIAL.
      SORT ACCIT_FI BY BUKRS.
      LOOP AT ACCIT_FI.
        IF ACCIT_FI-BUKRS NE LS_BUKRS-BUKRS.
          LS_BUKRS-BUKRS = ACCIT_FI-BUKRS.
          APPEND LS_BUKRS TO LT_BUKRS.
          IF ALE_MSG_TYPE = 'FIDCC1'
          OR ALE_MSG_TYPE = 'FIDCC2'.
            APPEND LS_BUKRS TO LT_COCD.
            REFRESH LT_COCD_ACT.
            LV_BUDAT = ACCIT_FI-BUDAT.
            CALL FUNCTION 'FAGL_INFO_GET'
              EXPORTING
                IT_BUKRS    = LT_COCD
                IV_BUDAT    = LV_BUDAT
              IMPORTING
                ET_COCD_ACT = LT_COCD_ACT.
            REFRESH LT_COCD.
            READ TABLE LT_COCD_ACT INDEX 1 INTO LS_COCD_ACT.
            IF LS_COCD_ACT-SPL_ACTIVE = CHAR_X.
* A posting, which has been split in FI in another system, has been
* distributed to this system, in which the document split of newGL
* is active, via ALE message type FIDCC1 or FIDCC2.
* So raise an error and prevent the posting as the newGL view
* is not transferred via IDOC/ALE with message types FIDCC1 and
* FIDCC2, and a correct document split of newGL cannot be executed
* on the basis of single FI documents, which result from such a
* split posting and which separately arrive and thus are separately
* processed in this receiving system.
              MESSAGE E200(GLT0)
                WITH ACCIT_FI-BELNR ACCIT_FI-BUKRS ACCIT_FI-GJAHR.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDLOOP.
* split posting might be further distributed to other systems
      PERFORM CHECK_ALE_OUTBOUND_N USING    LT_BUKRS
                                   CHANGING C_NO_SPLIT.
      IF NOT C_NO_SPLIT IS INITIAL.
        MESSAGE E035(FAGL_ALE).
      ENDIF.
* Indicator XSPLIT, which indicates that the FI document results
* from a posting, which was split in FI, is not transferred via
* IDOC/ALE. So set it in this receiving system again so that
* BKPF-XSPLIT and BSPL are updated there accordingly.
      READ TABLE ACCIT_FI INDEX 1.
      ACCIT_FI-XSPLIT = CHAR_X.
      MODIFY ACCIT_FI TRANSPORTING XSPLIT
        WHERE XSPLIT NE CHAR_X.
* do not execute the split as it was obviously already executed
* in the sending system
      C_NO_SPLIT = CHAR_X.
    ENDIF.

  ENDFORM.                    " CHECK_ALE_INBOUND_N

*&---------------------------------------------------------------------*
*&      Form CHECK_ALE_OUTBOUND_N
*&---------------------------------------------------------------------*
*       do not perform this kind of split if the FI documents are
*       distributed to other systems, where the document split of newGL
*       is active, via IDOC/ALE with message types 'FIDCC1' or 'FIDCC2'
*&---------------------------------------------------------------------*
  FORM CHECK_ALE_OUTBOUND_N USING    IT_BUKRS   TYPE FAGL_T_BUKRS
                            CHANGING C_NO_SPLIT TYPE XFELD.

    DATA: LT_ALE_MODEL_DATA_FIDCC1 TYPE TABLE OF BDI_MODEL
                                   WITH HEADER LINE,
          LT_ALE_MODEL_DATA_FIDCC2 TYPE TABLE OF BDI_MODEL
                                   WITH HEADER LINE,
          LT_ALE_MODEL_DATA_FIDCC* TYPE TABLE OF BDI_MODEL
                                   WITH HEADER LINE,
          LV_OWN_LOGICAL_SYSTEM    TYPE LOGSYS,
          LS_EDK13                 TYPE EDK13,
          LV_RFC_DESTINATION       TYPE RFCDEST,
          LV_RFC_TYPE              TYPE RFCTYPE,
          LV_RELEASE               TYPE SYSAPRL,
          LT_COMPONENT             TYPE TABLE OF OCS_C100
                                   WITH HEADER LINE,
          LS_BUKRS                 TYPE FAGL_S_BUKRS,
          LS_T001                  TYPE T001,
          LS_COCD_ACT              TYPE FAGL_S_ACT_CC,
          LT_COCD_ACT              TYPE FAGL_T_ACT_CC.

    CALL FUNCTION 'ALE_MODEL_INFO_GET'
      EXPORTING
        MESSAGE_TYPE           = 'FIDCC1'
      TABLES
        MODEL_DATA             = LT_ALE_MODEL_DATA_FIDCC1
      EXCEPTIONS
        NO_MODEL_INFO_FOUND    = 1
        OWN_SYSTEM_NOT_DEFINED = 2
        OTHERS                 = 3.
    CALL FUNCTION 'ALE_MODEL_INFO_GET'
      EXPORTING
        MESSAGE_TYPE           = 'FIDCC2'
      TABLES
        MODEL_DATA             = LT_ALE_MODEL_DATA_FIDCC2
      EXCEPTIONS
        NO_MODEL_INFO_FOUND    = 1
        OWN_SYSTEM_NOT_DEFINED = 2
        OTHERS                 = 3.

    APPEND LINES OF LT_ALE_MODEL_DATA_FIDCC1 TO LT_ALE_MODEL_DATA_FIDCC*.
    APPEND LINES OF LT_ALE_MODEL_DATA_FIDCC2 TO LT_ALE_MODEL_DATA_FIDCC*.

    DESCRIBE TABLE LT_ALE_MODEL_DATA_FIDCC* LINES SY-TFILL.
    CHECK SY-TFILL > 0.

    CALL FUNCTION 'OWN_LOGICAL_SYSTEM_GET'
      IMPORTING
        OWN_LOGICAL_SYSTEM             = LV_OWN_LOGICAL_SYSTEM
      EXCEPTIONS
        OWN_LOGICAL_SYSTEM_NOT_DEFINED = 1.

    LOOP AT IT_BUKRS INTO LS_BUKRS.

      CALL FUNCTION 'FI_COMPANY_CODE_DATA'
        EXPORTING
          I_BUKRS = LS_BUKRS-BUKRS
        IMPORTING
          E_T001  = LS_T001.

      CHECK NOT LS_T001-BUKRS_GLOB IS INITIAL.

      LOOP AT LT_ALE_MODEL_DATA_FIDCC*
        WHERE RCVSYSTEM NE LV_OWN_LOGICAL_SYSTEM
        AND ( OBJTYPE EQ 'BUKRS' AND
              OBJVALUE EQ LS_T001-BUKRS_GLOB OR
              OBJTYPE IS INITIAL ).

        LS_EDK13-RCVPRN = LT_ALE_MODEL_DATA_FIDCC*-RCVSYSTEM.
        LS_EDK13-RCVPRT = 'LS'.
        LS_EDK13-MESTYP = LT_ALE_MODEL_DATA_FIDCC*-MESTYP.
* read partner agreement / distribution to a SAP system ?
        CALL FUNCTION 'EDI_AGREE_OUT_MESSTYPE_READ'
          EXPORTING
            REC_EDK13       = LS_EDK13
          EXCEPTIONS
            ENTRY_NOT_EXIST = 1.
        IF SY-SUBRC = 1.
          CONTINUE.
        ENDIF.
* determine the RFC destination from the logical system
        CALL FUNCTION 'RFCDEST_GET_FOR_LOGICAL_SYSTEM'
          EXPORTING
            MESSAGE_TYPE            = LT_ALE_MODEL_DATA_FIDCC*-MESTYP
            LOGICAL_SYSTEM          = LT_ALE_MODEL_DATA_FIDCC*-RCVSYSTEM
          IMPORTING
            RFC_DESTINATION         = LV_RFC_DESTINATION
          EXCEPTIONS
            NO_RFCDESTINATION_FOUND = 1.
        IF SY-SUBRC NE 0.
          C_NO_SPLIT = CHAR_X.
          EXIT.
        ENDIF.
* check whether the RFC destination exists
        CALL FUNCTION 'RFC_READ_DESTINATION_TYPE'
          EXPORTING
            DESTINATION           = LV_RFC_DESTINATION
            AUTHORITY_CHECK       = ' '
            BYPASS_BUF            = ' '
          IMPORTING
            RFCTYPE               = LV_RFC_TYPE
          EXCEPTIONS
            DESTINATION_NOT_EXIST = 1
            INFORMATION_FAILURE   = 2
            INTERNAL_FAILURE      = 3.
        IF SY-SUBRC NE 0.
          C_NO_SPLIT = CHAR_X.
          EXIT.
        ENDIF.
* check whether the RFC destination is a R/3 / ERP system
        CHECK LV_RFC_TYPE-RFCTYPE EQ '3'.
* check whether release of destination R/3 / ERP system is >= ERP 6.0
* or = 4.70
        CLEAR LT_COMPONENT.
        REFRESH LT_COMPONENT.
        CALL FUNCTION 'OCS_GET_SYSTEM_INFO'
          DESTINATION LV_RFC_DESTINATION
          IMPORTING
            EV_SAPRL              = LV_RELEASE
          TABLES
            TT_COMPONENT          = LT_COMPONENT
          EXCEPTIONS
            SYSTEM_FAILURE        = 1
            COMMUNICATION_FAILURE = 2.
        IF SY-SUBRC NE 0 OR LV_RELEASE(2) < '40'.
          C_NO_SPLIT = CHAR_X.
          EXIT.
        ENDIF.
        LOOP AT LT_COMPONENT WHERE LINE(30) = 'SAP_APPL'.
          EXIT.
        ENDLOOP.
        IF SY-SUBRC EQ 0.
          IF LT_COMPONENT-LINE+30(10) < '600' AND
             LT_COMPONENT-LINE+30(10) <> '470'.
            C_NO_SPLIT = CHAR_X.
            EXIT.
          ENDIF.
        ELSE.
          CONTINUE.
        ENDIF.
        IF LT_COMPONENT-LINE+30(10) >= '600'.
          REFRESH LT_COCD_ACT.
* determine whether the document split of newGL is active in the
* system, to which the FI documents will be sent
          CALL FUNCTION 'FAGL_INFO_GET'
            DESTINATION LV_RFC_DESTINATION
            EXPORTING
              IV_GLOB_BUKRS         = LS_T001-BUKRS_GLOB
            IMPORTING
              ET_COCD_ACT           = LT_COCD_ACT
            EXCEPTIONS
              SYSTEM_FAILURE        = 1
              COMMUNICATION_FAILURE = 2.
          IF SY-SUBRC NE 0.
            C_NO_SPLIT = CHAR_X.
            EXIT.
          ENDIF.
          READ TABLE LT_COCD_ACT INDEX 1 INTO LS_COCD_ACT.
          IF LS_COCD_ACT-SPL_ACTIVE = CHAR_X.
            C_NO_SPLIT = CHAR_X.
            EXIT.
          ENDIF.
        ENDIF.

      ENDLOOP.

      IF C_NO_SPLIT = CHAR_X.
        EXIT.
      ENDIF.

    ENDLOOP.

  ENDFORM.                    " CHECK_ALE_OUTBOUND_N

*&---------------------------------------------------------------------*
*&      Form VAT_BREAKDOWN_N
*&---------------------------------------------------------------------*
  FORM VAT_BREAKDOWN_N.

    DATA: LS_T001         TYPE T001,
          LS_T003         TYPE T003,
          LV_MW2TAB_COUNT TYPE I,
          LV_SAVE_HWBAS   TYPE HWBAS.

    LOOP AT LOGDN.
      LOOP AT ACCIT_FI WHERE KOART CA 'DVK' AND KTOSL NE 'BUV'
                                            AND LOGVO EQ LOGDN-LOGVO.
        EXIT.
      ENDLOOP.
      CHECK SY-SUBRC IS INITIAL.
      CALL FUNCTION 'FI_DOCUMENT_TYPE_DATA'
        EXPORTING
          I_BLART = ACCIT_FI-BLART
        IMPORTING
          E_T003  = LS_T003.
      CHECK LS_T003-XNETB IS INITIAL.
      REFRESH MW2TAB.
      LOOP AT ACCIT_FI WHERE KOART CA 'ASM'
                         AND LOGVO EQ LOGDN-LOGVO
                         AND KTOSL NE 'KDT'.
        CLEAR MW2TAB.
        MW2TAB-MWSKZ = ACCIT_FI-MWSKZ.
        MW2TAB-TXDAT_FROM = ACCIT_FI-TXDAT_FROM.
        MW2TAB-TAX_COUNTRY = ACCIT_FI-TAX_COUNTRY.
        MOVE-CORRESPONDING ACCIT_FI TO ACCCR_KEY.
        ACCCR_KEY-CURTP = '10'.
        READ TABLE ACCCR_FI BINARY SEARCH
          WITH KEY ACCCR_KEY.
* VAT tax line item
        IF ACCIT_FI-MWART CA 'AV'.
          IF XEXRT IS INITIAL.
            MW2TAB-DMTAX = ACCCR_FI-WRBTR.
          ELSE.
            LV_SAVE_HWBAS = ACCCR_FI-FWBAS.
            ACCCR_KEY-CURTP = '00'.
            READ TABLE ACCCR_FI BINARY SEARCH
              WITH KEY ACCCR_KEY.
            MW2TAB-DMTAX =
              ACCCR_FI-WRBTR * LV_SAVE_HWBAS / ACCCR_FI-FWBAS.
          ENDIF.
* other line item
        ELSE.
          MW2TAB-DMBTR = ACCCR_FI-WRBTR.
          IF ACCIT_FI-XSKRL IS INITIAL.
            MW2TAB-DMSKT = ACCCR_FI-WRBTR.
          ENDIF.
        ENDIF.
        COLLECT MW2TAB.
      ENDLOOP.
* add proportionate VAT
      CALL FUNCTION 'FI_COMPANY_CODE_DATA'
        EXPORTING
          I_BUKRS = ACCIT_FI-BUKRS
        IMPORTING
          E_T001  = LS_T001.
      IF LS_T001-XSKFN IS INITIAL.
        LOOP AT MW2TAB WHERE DMBTR NE 0.
          MW2TAB-DMSKT = MW2TAB-DMSKT
                       + MW2TAB-DMTAX * MW2TAB-DMSKT / MW2TAB-DMBTR.
          MODIFY MW2TAB.
        ENDLOOP.
      ENDIF.

      LOOP AT ACCIT_FI WHERE KOART CA 'DVK' AND KTOSL NE 'BUV'
                                            AND LOGVO EQ LOGDN-LOGVO.
        CHECK ACCIT_FI-UMSKS NE CHAR_A.
        CHECK ACCIT_FI-REBZT NE CHAR_V.
        CLEAR: ACCIT_FI-MWSK1, ACCIT_FI-DMBT1,
               ACCIT_FI-MWSK2, ACCIT_FI-DMBT2,
               ACCIT_FI-MWSK3, ACCIT_FI-DMBT3,
               ACCIT_FI-TXDAT_FROM1,
               ACCIT_FI-TXDAT_FROM2,
               ACCIT_FI-TXDAT_FROM3,
               ACCIT_FI-TAX_COUNTRY1,
               ACCIT_FI-TAX_COUNTRY2,
               ACCIT_FI-TAX_COUNTRY3.
        CLEAR LV_MW2TAB_COUNT.
        LOOP AT MW2TAB WHERE DMBTR NE 0.
          LV_MW2TAB_COUNT = LV_MW2TAB_COUNT + 1.
        ENDLOOP.
        CHECK LV_MW2TAB_COUNT GT 0.
        IF LV_MW2TAB_COUNT GT 1.
          LOOP AT MW2TAB WHERE DMBTR NE 0.
            IF ACCIT_FI-DMBT1 = 0.
              ACCIT_FI-MWSK1 = MW2TAB-MWSKZ.
              ACCIT_FI-TAX_COUNTRY1 = MW2TAB-TAX_COUNTRY.
              ACCIT_FI-TXDAT_FROM1 = MW2TAB-TXDAT_FROM.
              ACCIT_FI-DMBT1 = MW2TAB-DMSKT.
            ELSE.
              IF ACCIT_FI-DMBT2 = 0.
                ACCIT_FI-MWSK2 = MW2TAB-MWSKZ.
                ACCIT_FI-TAX_COUNTRY2 = MW2TAB-TAX_COUNTRY.
                ACCIT_FI-TXDAT_FROM2 = MW2TAB-TXDAT_FROM.
                ACCIT_FI-DMBT2 = MW2TAB-DMSKT.
              ELSE.
                IF ACCIT_FI-DMBT3 = 0.
                  ACCIT_FI-MWSK3 = MW2TAB-MWSKZ.
                  ACCIT_FI-TAX_COUNTRY3 = MW2TAB-TAX_COUNTRY.
                  ACCIT_FI-TXDAT_FROM3 = MW2TAB-TXDAT_FROM.
                  ACCIT_FI-DMBT3 = MW2TAB-DMSKT.
                ENDIF.
              ENDIF.
            ENDIF.
          ENDLOOP.
* correct sign
          IF ACCIT_FI-SHKZG = CHAR_S.
            ACCIT_FI-DMBT1 = - ACCIT_FI-DMBT1.
            ACCIT_FI-DMBT2 = - ACCIT_FI-DMBT2.
            ACCIT_FI-DMBT3 = - ACCIT_FI-DMBT3.
          ENDIF.
* set unique VAT tax code
          IF  SY-SUBRC = 0
          AND ACCIT_FI-DMBT1 <> 0
          AND ACCIT_FI-DMBT2 = 0
          AND ACCIT_FI-DMBT3 = 0.
* problems with incoming payments with cash discount -
* set cash discount relevant characteristic
            IF ACCIT_FI-MWSKZ EQ '**'.
              ACCIT_FI-MWSKZ = ACCIT_FI-MWSK1.
              ACCIT_FI-TXDAT_FROM = ACCIT_FI-TXDAT_FROM1.
              ACCIT_FI-TAX_COUNTRY = ACCIT_FI-TAX_COUNTRY1.
            ENDIF.
            ACCIT_FI-MWSK1 = SPACE.
            CLEAR ACCIT_FI-TXDAT_FROM1.
            CLEAR ACCIT_FI-TAX_COUNTRY1.
            ACCIT_FI-DMBT1 = 0.
          ENDIF.
        ELSE.
* problems with incoming payments with cash discount -
* set cash discount relevant characteristic
          IF ACCIT_FI-MWSKZ EQ '**'.
            ACCIT_FI-MWSKZ = MW2TAB-MWSKZ.
            ACCIT_FI-TAX_COUNTRY = MW2TAB-TAX_COUNTRY.
            ACCIT_FI-TXDAT_FROM = MW2TAB-TXDAT_FROM.
          ENDIF.
        ENDIF.
        MODIFY ACCIT_FI.
      ENDLOOP.
    ENDLOOP.

  ENDFORM.                    " VAT_BREAKDOWN_N

*&---------------------------------------------------------------------*
*&      Form VAT_BREAKDOWN_N
*&---------------------------------------------------------------------*
*& remove TAXPS again from G/L line items without tax code and
*& automatically created line items in order to prevent problems
*& within the later called FI tax function modules
*&---------------------------------------------------------------------*
  FORM INIT_TAXPS.

    FIELD-SYMBOLS: <ACCIT_FI> TYPE ACCIT_FI.

    LOOP AT ACCIT_FI ASSIGNING <ACCIT_FI>
                     WHERE ( ( KOART CA 'MSA' OR
                               KTOSL EQ 'EGX' ) AND
                             TAXIT IS INITIAL AND
                             MWSKZ IS INITIAL ) OR
                             TAXPS EQ '999999'.
      CLEAR <ACCIT_FI>-TAXPS.
    ENDLOOP.

  ENDFORM.

*&---------------------------------------------------------------------*
*&      Form check_parking_xtxit
*&---------------------------------------------------------------------*
  FORM CHECK_PARKING_XTXIT CHANGING CV_NO_SPLIT TYPE XFELD
                                    CV_XTXIT    TYPE XTXIT_TXD.

    CHECK ( ACCHD_FI-STATUS_NEW EQ '2' OR ACCHD_FI-STATUS_NEW EQ '3' )
      AND NOT CV_XTXIT IS INITIAL.

    IF ACCHD_FI-STATUS_NEW EQ '3'.
      CV_NO_SPLIT = CHAR_X.
      EXIT.
    ENDIF.

    DESCRIBE TABLE XBSET LINES SY-TFILL.
    IF SY-TFILL > 0.
      CV_NO_SPLIT = CHAR_X.
      EXIT.
    ENDIF.

    LOOP AT ACCIT_FI TRANSPORTING NO FIELDS WHERE TAXPS IS NOT INITIAL.
      CV_NO_SPLIT = CHAR_X.
      EXIT.
    ENDLOOP.

    CHECK CV_NO_SPLIT IS INITIAL.

    CLEAR CV_XTXIT.

  ENDFORM.
