*
*---------------------------------------------------------------------*
*       FORM FI_DOCUMENT_COLLECT                                      *
*---------------------------------------------------------------------*
*       FI-Beleg wird aus RW-Beleg projeziert                         *
*---------------------------------------------------------------------*
FORM FI_DOCUMENT_COLLECT  USING I_AWTYP LIKE ACCIT-AWTYP
                                I_AWREF LIKE ACCIT-AWREF
                                I_AWORG LIKE ACCIT-AWORG
                                I_POSNR LIKE ACCIT-POSNR
                                I_ACTIV LIKE BKPF-XUSVR
                                I_XTXIT TYPE XFELD        "note 2410917
                                I_EG_VAT TYPE XFELD.      "2538979 :VAT EG: Line item wise tax calculation

  DATA: FIELD(11).                                          "N2110892
  FIELD-SYMBOLS: <VALUE>.                                   "N2110892

  DATA: H_DMBTR LIKE ACCCR-WRBTR,
        L_XMWST TYPE XMWST.                               "note 2305549

* SFIN2.0: remove CO account assignment for FI-ALE
  DATA: LS_ACDOC_ITEM_CO         TYPE FINS_ACDOC_ITEM_CO,
* 1905: Collect postings from API automatically
        LS_ACDOC_ITEM_ATTR_CO_GL TYPE FINS_ACDOC_ITEM_ATTR_CO_GL.
*  STATICS: lt_ttyps TYPE STANDARD TABLE OF ttyps.
*           lo_fqm   TYPE REF TO cl_fqm_derive_bseg.
*
  CLEAR P_ACC.
*
* Wird überhaupt komprimiert?
*
  LOOP AT GT_TTYPV WHERE AWTYP EQ I_AWTYP                 "note 1779136
                   AND ( BUKRS EQ ACCIT_FI-BUKRS          "note 1779136
                   OR    BUKRS EQ SPACE )                 "note 1779136
                   AND ( BLART EQ ACCIT_FI-BLART          "note 1779136
                   OR    BLART EQ SPACE ).                "note 1779136
    EXIT.                                                 "note 1369659
  ENDLOOP.                                                "note 1369659
  IF SY-SUBRC IS INITIAL                                  "note 1369659
     AND ( ( I_XTXIT IS INITIAL OR GX_GST_USED IS INITIAL )  "n 2675869       "India GST changes
     AND GX_GST_MY IS INITIAL ).
    COMPRES = CHAR_X.                                     "note 1369659
  ELSE.                                                   "note 1369659
    COMPRES = SPACE.                                      "note 1369659
  ENDIF.                                                  "note 1369659
*
  ACCIT_KEY-MANDT = SY-MANDT.
  ACCIT_KEY-AWTYP = I_AWTYP.
  ACCIT_KEY-AWREF = I_AWREF.
  ACCIT_KEY-AWORG = I_AWORG.
  ACCIT_KEY-POSNR = I_POSNR.
*
  CLEAR T_ACCFI.
  IF NOT ACCIT_FI-XCPDD IS INITIAL.
*
* Kennzeichen XCPDD ist für CPD und abweichenden Zahlungsempfänger gef.
*
    READ TABLE T_ACCFI WITH KEY ACCIT_KEY.
  ENDIF.
*
* Währungstypen des Buchungskreises lesen
*
  PERFORM DETERMINE_CURTP TABLES   ICURTP
                          USING    ACCIT_FI-BUKRS.
*
* Normalisierten Währungsvektor in flache FI-Belegstrukturen füllen
*
  PERFORM FI_ITEM_FILL.

  READ TABLE ACCCR_FI BINARY SEARCH WITH KEY ACCIT_KEY.
  TABIX = SY-TABIX.
  LOOP AT ACCCR_FI FROM TABIX.
    IF ACCCR_FI-AWTYP <> I_AWTYP OR
       ACCCR_FI-AWREF <> I_AWREF OR
       ACCCR_FI-AWORG <> I_AWORG OR
       ACCCR_FI-POSNR <> I_POSNR.
      EXIT.
    ENDIF.
*
    LOOP AT ICURTP WHERE BUKRS EQ ACCIT_FI-BUKRS
                   AND   CURTP EQ ACCCR_FI-CURTP.
      PERFORM FI_PROJECTION_FILL USING  ICURTP-TYPE.
      EXIT.
    ENDLOOP.
* amount in company currency requested
* nur falls Betrag von Außen mitgegen
    IF ACCCR_FI-ISTAT EQ 0
    AND ACCCR_FI-CURTP EQ '10'
    AND ACCCR_FI-WRBTR EQ 0.
      IF P_ACC-WRBTR NE SPACE
      AND P_ACC-MWART IS INITIAL
      AND P_ACC-AWTYP NE 'BKPFF'                          "note 1040798
      AND P_ACC-AWTYP NE 'BKPFE'                      "ECS Direct Input
      AND P_ACC-TCODE NE 'FBB1'
* position not marked for foreign currency valuation
      AND ACCIT_FI-FCVAL EQ SPACE
* No check required for TR Transactions.
      AND P_ACC-AWTYP NE 'TR-TM'                           "note 385815
* no check for loans
      AND P_ACC-AWTYP NE 'LOANS'                          "note 1061510
* no check for real estate (RE)
      AND P_ACC-AWTYP NE 'REACI'                          "note 1315492
* TA Währung gleich Null ist in Rechnungsprüfung möglich
      AND P_ACC-GLVOR NE 'RMRP'
* Keine Prüfung bei Vertragskontokorrent
      AND P_ACC-AWTYP NE 'FKKSU'                           "note 370188
      AND P_ACC-AWTYP NE 'IDOC'                            "note 370188
      AND P_ACC-AWTYP(1) NE 'Z'                            "note 370188
      AND P_ACC-AWTYP(1) NE 'Y'                           "note 3406879
      AND P_ACC-AWTYP NE 'IBKPF'                          "note 1578618
* Auftragsabrechnung mit HW = 0 ebenfalls möglich
      AND P_ACC-GLVOR NE 'KOAE'
      AND P_ACC-GLVOR NE 'RFRK'        "Hinweis 164841
**** HW 944416
      AND P_ACC-GLVOR NE 'COFI'
* Keine Prüfung bei ktosl <> initial - es treten zu viele Fälle auf
* in denen diese Situation zulässig ist.
      AND P_ACC-KTOSL IS INITIAL                           "note 209649
* VPRS aus SD-BIL
      AND P_ACC-AWTYP NE 'VBRK'                            "note 932849
* FBB1 mit ALE über IDOC FIDCMT
      AND ALE_FLAG EQ SPACE                                "note 312699
      AND P_ACC-STBLG EQ SPACE                             "note 312699
* Central Finance
      AND ( P_ACC-BUKRS_SENDER IS INITIAL OR P_ACC-BELNR_SENDER IS INITIAL OR P_ACC-GJAHR_SENDER IS INITIAL )
*NewGL-Allocations
      AND P_ACC-AWTYP NE 'GLAL0'
* report FAGL_YEC_POSTINGS
      AND P_ACC-AWTYP NE 'GLYEC'                          "note 1307096
* no check for reorganization (e.g. profit center reorg: FI-GL-REO)
      AND P_ACC-AWTYP NE 'BKPFR'                          "note 1671509
      AND P_ACC-AWTYP NE 'BKPWF'                          "note 3017681
* no check for asset postings                             "note 1431968
      AND P_ACC-AWTYP NE 'AMBU'                           "note 1431968
      AND P_ACC-AWTYP NE 'AMDP'                           "note 1431968
      AND GV_CASH_LEDGER NE 'X'                           "note 3298370
      AND P_ACC-AWTYP NE 'FS-BA'                          "note 1800680
      AND P_ACC-AWTYP NE 'FARR'                           "note 2153391
      AND P_ACC-AWTYP NE 'FARA'                           "note 2153391
      AND P_ACC-AWTYP NE 'BKPFO'                          "note 3264628 : Alignment of valuation object
      AND P_ACC-GLVOR NE 'TBCS'                           "note 2877622
      AND P_ACC-AWTYP NE 'CF3P'                           "note 3044075
      AND P_ACC-GLVOR NE CL_FINS_TRR_RULE_VAL=>GC_GLVOR_TRB_REV_REC        "event based revenue recognition
      AND P_ACC-AWTYP NE CL_FINS_TRR_RULE_VAL=>GC_AWTYP_TRB_REC_REC_CLOSE  "period end closing event based revenue recognition
      AND P_ACC-AWTYP NE IF_ACE_MDO_TYPES=>CV_AWTYP_S4.
* ggf. Kleinbeträge akzeptieren & Nachrechnen
        CALL FUNCTION 'CONVERT_TO_LOCAL_CURRENCY'
          EXPORTING
            LOCAL_CURRENCY   = P_ACC-HWAER
            FOREIGN_CURRENCY = P_ACC-WAERS
            FOREIGN_AMOUNT   = P_ACC-WRBTR
            DATE             = SY-DATUM
            RATE             = P_ACC-KURSF
*           TYPE_OF_RATE     = P_ACC-KURST
          IMPORTING
            LOCAL_AMOUNT     = H_DMBTR.
        IF H_DMBTR NE SPACE.
          MESSAGE E800 WITH P_ACC-BUZEI P_ACC-BUKRS '10' P_ACC-HWAER.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDLOOP.
  IF P_ACC-MWART NE SPACE.                                 "note 601139
ENHANCEMENT-SECTION     FI_DOCUMENT_COLLECT_01 SPOTS ES_SAPLFACI.
    IF ACCHD_FI-GLVOR EQ 'SD00'                            "note 601139
* all SD (i.e. glvor = SD00) tax items must be summarized     N2203186
* as SD does not know 'direct' or 'manual' tax items in terms N2203186
* of accounting, the konw only 'tax only' postings            N2203186
** summarize only automatic tax items               "N2109349 N2203186
**   AND NOT p_acc-xauto IS INITIAL.                "N2109349 N2203186
      AND ( ( I_XTXIT IS INITIAL OR GX_GST_USED IS INITIAL )  "2675869       "India GST changes
      AND     I_EG_VAT IS INITIAL AND GX_GST_MY IS INITIAL ).                          "2538979 :VAT EG: Line item wise tax calculation
* clear material specific information on tax items
      CLEAR: P_ACC-WERKS,
             P_ACC-MATNR,
             P_ACC-MENGE,
             P_ACC-MEINS,
             P_ACC-BPMNG,
             P_ACC-PPRCT,
             P_ACC-PSEGMENT,                              "note 2944769
             P_ACC-PAOBJNR,
             P_ACC-POSN2,
             P_ACC-VBEL2,
             P_ACC-BWTAR,
             P_ACC-EBELN,
             P_ACC-EBELP,                                  "note 333602
             P_ACC-TAXPS.                                  "note 333602
    ENDIF.                                                 "note 601139

END-ENHANCEMENT-SECTION.
* clear TAXPS as otherwise the summarization of            "note 601139
* tax line items in table BSEG is prevented                "note 601139
*    IF ACCHD_FI-AWTYP EQ 'RMRP'.                 "note 601139 N1385914
*      CLEAR P_ACC-TAXPS.                         "note 601139 N1385914
*    ENDIF.                                       "note 601139 N1385914
  ENDIF.
* Clear TAXPS only when compres activ or as definded for GLVORs"N1223773
  IF ( COMPRES = CHAR_X                                     "N1223773
  OR NOT G_BADI_TAX_COMPRESS IS INITIAL )                   "N1575217
  AND ( ( I_XTXIT IS INITIAL OR GX_GST_USED IS INITIAL )    "n2675869     "India GST changes
  AND     I_EG_VAT IS INITIAL AND GX_GST_MY IS INITIAL )                            "2538979 :VAT EG: Line item wise tax calculation
  AND  P_ACC-TAXPS NE 0 AND P_ACC-TAXPS NE SPACE.           "N2185172
* store that taxps was cleared in order to prevent FF 818    N2165871
    CALL FUNCTION 'FI_TAX_STORE_TAXPS_CLEARED'              "N2165871
      IMPORTING
        E_XMWST = L_XMWST.                                  "N2305549
*                                                    "N1070206 "N1223773
    IF L_XMWST IS INITIAL.                                  "N2305549
      CLEAR P_ACC-TAXPS.                             "N1070206 "N1223773
      IF NOT ACCIT_FI-TXGRP IS INITIAL.                     "N2110892
* txgrp must also be cleared - if it exists and is filled;   N2110892
* accit_fi-txgrp exists always if accit has been enhanced;   N2110892
* but p_acc-txgrp exists only if SAPFACCGEN has run too;     N2110892
* in order to avoid F5843 check and clear p_acc-txgrp        N2110892
* only if accit_fi-txgrp is really filled                    N2110892
        FIELD = 'P_ACC-TXGRP'.                              "N2110892
        ASSIGN (FIELD) TO <VALUE>.                          "N2110892
        IF SY-SUBRC IS INITIAL.                             "N2110892
* clear only if field exists and was filled                  N2110892
          CLEAR <VALUE>.                                    "N2110892
        ENDIF.                                              "N2110892
      ENDIF.                                                "N2110892
    ENDIF.                                                  "N2305549
  ENDIF.                                                    "N1223773
  FI_TAX_SERVICES6=>GET_INSTANCE( )->SET_REV_REAL(          "2819789
                    I_REV_REAL = GX_REV_REAL ).             "2819789
* save tax items for external interface before compression
  IF I_ACTIV NE SPACE
  AND P_ACC-MWART NE SPACE.
    IF XRWBSET_HEAD IS INITIAL.
* Cancelation of billing document transfer AWREF_REV/AWORG_REV
      MOVE-CORRESPONDING ACCHD_FI TO XRWBSET_HEAD.
      MOVE-CORRESPONDING ACCIT_FI TO XRWBSET_HEAD.
    ENDIF.
* only in case of uncompressed tax data from SD
    IF P_ACC-GLVOR EQ 'SD00'.
      MOVE-CORRESPONDING P_ACC TO XRWBSET.
      MOVE-CORRESPONDING ACCIT_KEY TO XRWBSET.
* Pointer for link with external tax table ttxy
      XRWBSET-TAXPS = ACCIT_FI-TAXPS.
      XRWBSET-OLD_DOC_NUMBER = ACCIT_FI-OLD_DOC_NUMBER.
* KBERTC is only necessary if collect xrwbset occurs
      APPEND XRWBSET.
    ENDIF.
* Baustelle: Welche Info nach BSET übertragen ?
  ENDIF.
*
* save only one KNUMH per VAT indicator - KNUMH is changed here!
* KHE: for US Taxes TXJCD
  IF NOT P_ACC-KNUMH IS INITIAL.
    LOOP AT SKNUMH WHERE BELNR EQ P_ACC-BELNR
                   AND   TAX_COUNTRY = P_ACC-TAX_COUNTRY
                   AND   MWSKZ EQ P_ACC-MWSKZ
                   AND   TXDAT_FROM EQ P_ACC-TXDAT_FROM
                   AND   TXJCD EQ P_ACC-TXJCD
                   AND   KSCHL EQ P_ACC-KSCHL.
    ENDLOOP.
    IF SY-SUBRC IS INITIAL.
      P_ACC-KNUMH = SKNUMH-KNUMH.
    ELSE.
      MOVE-CORRESPONDING P_ACC TO SKNUMH.
      APPEND SKNUMH.
    ENDIF.
  ENDIF.
* begin of note 437983
* prepare tax line items from SD for summarization accross SHKZG
  IF P_ACC-MWART NE SPACE
* all SD tax items must be summarized                        N2203186
* summarize only automatic tax items                        "N2109349
* AND NOT p_acc-xauto IS INITIAL                   "N2109349 N2203186
  AND ( P_ACC-AWTYP EQ 'VBRK' OR
        P_ACC-AWTYP EQ 'JHTFK' OR
        P_ACC-AWTYP EQ 'JFRK' )                            "note 761574
  AND ALE_FLAG IS INITIAL.                                "note 1118145
* begin of note 544549
* If the external tax interface is active, tax line items from
* SD can occur with tax amount = 0 and KBETR <> 0 (-> exemption).
* Such tax line items should not be compressed across the debit/
* credit indicator
    IF I_ACTIV NE SPACE
    AND P_ACC-WRBTR EQ 0 AND P_ACC-DMBTR EQ 0
    AND P_ACC-DMBE2 EQ 0 AND P_ACC-DMBE3 EQ 0
    AND ACCIT_FI-KBETR NE 0.
      IF P_ACC-XAUTO EQ CHAR_X.
        P_ACC-XAUTO = 'E'.
      ELSEIF P_ACC-XAUTO EQ SPACE.
        P_ACC-XAUTO = 'e'.
      ENDIF.
    ENDIF.
* end of note 544549
    IF P_ACC-SHKZG EQ CHAR_H.
      P_ACC-WRBTR = - P_ACC-WRBTR.
      P_ACC-FWSTE = - P_ACC-FWSTE.
      P_ACC-FWBAS = - P_ACC-FWBAS.
      P_ACC-PSWBT = - P_ACC-PSWBT.
      P_ACC-DMBTR = - P_ACC-DMBTR.
      P_ACC-HWSTE = - P_ACC-HWSTE.
      P_ACC-HWBAS = - P_ACC-HWBAS.
      P_ACC-DMBE2 = - P_ACC-DMBE2.
      P_ACC-H2STE = - P_ACC-H2STE.
      P_ACC-H2BAS = - P_ACC-H2BAS.
      P_ACC-TXBH2 = - P_ACC-TXBH2.
      P_ACC-DMBE3 = - P_ACC-DMBE3.
      P_ACC-H3STE = - P_ACC-H3STE.
      P_ACC-H3BAS = - P_ACC-H3BAS.
      P_ACC-TXBH3 = - P_ACC-TXBH3.
    ENDIF.
*   Cloud summarize taxes for each shkzg separately
    IF CL_FINS=>IS_CLOUD_ACTIVE( ) = ABAP_FALSE.
      CLEAR: P_ACC-BSCHL,
             P_ACC-SHKZG.
    ENDIF.
  ENDIF.
* end of note 437983
* delete rates and translation date for items with KRUEK filled
* in order to avoid that BKPF will be filled from those items
  IF ACCIT_FI-KRUEK = CHAR_X                               "note 795865
  AND ( ACCIT_FI-AWTYP = 'VBRK'                            "note 795865
  OR ACCIT_FI-AWTYP = 'VBRR' )                             "note 795865
  AND X_ONLY_HIST_RATE IS INITIAL.                        "note 1255920
    CLEAR: P_ACC-KURSF, P_ACC-KURS2, P_ACC-KURS3, P_ACC-WWERT. "n795865
  ENDIF.                                                   "note 795865

* sFIN2.0: ALE FIDCC1/FIDCC2 remove CO account assignment to avoid redundent update of acdoca
* in case CO will be updated in parallel with its own ALE scenario
  IF ALE_FLAG = ABAP_TRUE
  AND ( ALE_MSG_TYPE = 'FIDCC1'
  OR    ALE_MSG_TYPE = 'FIDCC2' ).
* check if CO is included in subset due to customer enhancement
* only if sv_subset_wo_co = true clear of co account assignment is required
* clear co account assignment within FI ALE scenario
    CL_FINS_ACDOC_POST_UTIL=>CLEAR_CO_ACCOUNT_ASSGMT_FI_ALE(
      EXPORTING
        IV_GJAHR                    = ACCIT_FI-GJAHR
        IV_SUBSET                   = ACCHD_FI-SUBSET
        IV_ALE_MSG_TYPE             = ALE_MSG_TYPE
      RECEIVING
        RV_CLEAR_CO_ACCNT_ASSGNMENT = DATA(LV_CLEAR_CO_ACCNT_ASSGNMENT)
    ).

    IF LV_CLEAR_CO_ACCNT_ASSGNMENT = ABAP_TRUE.
* clear co relevant fields within p_acc
      MOVE-CORRESPONDING LS_ACDOC_ITEM_CO TO P_ACC.
    ENDIF.
  ENDIF.
* 1905 - remove account assignment of GL items from API/BAPI to allow a compression automatically
  IF ( ACCHD_FI-EV_POSTNG_CTRL = IF_FINS_ACDOC_CONSTANTS=>GC_ENTRY_VIEW_POST_CONTROL-MAX_COMPRESS
  OR   ACCHD_FI-EV_POSTNG_CTRL = IF_FINS_ACDOC_CONSTANTS=>GC_ENTRY_VIEW_POST_CONTROL-LEAVE_OUT_MAX_COMPRESS )
  AND CL_FINS=>IS_CLOUD_ACTIVE( ) = ABAP_TRUE
  AND ACCIT_FI-MWART = SPACE
  AND ACCIT_FI-XOPVW = ABAP_FALSE
  AND ACCIT_FI-XLGCLR = ABAP_FALSE
* Exclude BUV items from summarization
  AND ACCIT_FI-KTOSL  <> IF_FINS_ACDOC_CONSTANTS=>GC_ACCOUNT_DET_GROUP-CROSS_COMPANY.
*    IF acchd_fi-tcode(6) <> cl_fdc_posting_impl=>co_fdc_s_accdoc_hdr_gen_tcode-clearing.
    MOVE-CORRESPONDING  LS_ACDOC_ITEM_ATTR_CO_GL TO P_ACC.
    P_ACC-BUZID = IF_FINS_ACDOC_CONSTANTS=>GC_BUZID-COMPRESSED.
**   CE2308: For payment transaction compress items created based on document splitting
*    ELSE.
*      IF p_acc-ktosl = if_fins_acdoc_constants=>gc_account_det_group-exchange_rate_difference "KDF
*      OR p_acc-ktosl = if_fins_acdoc_constants=>gc_account_det_group-cash_discount_expenses   "SKT
*      OR p_acc-ktosl = if_fins_acdoc_constants=>gc_account_det_group-cash_discount_received   "SKE
*      OR p_acc-ktosl = if_fins_acdoc_constants=>gc_account_det_group-lost_cash_discount.      "VSK
*        MOVE-CORRESPONDING  ls_acdoc_item_attr_co_gl TO p_acc.
*        p_acc-buzid = if_fins_acdoc_constants=>gc_buzid-compressed.
*      ENDIF.
*    ENDIF.
  ENDIF.
* CE2208 - remove all G/L account assignment for all items of tax payable posting
  IF (  ACCHD_FI-AWTYP = IF_FINS_ACDOC_CONSTANTS=>GC_AWTYP-TXPY
     OR P_ACC-TCODE = CL_FDC_POSTING_IMPL=>CO_FDC_S_ACCDOC_HDR_TCODE-GL_TAX_PAYABLE )
  AND ACCHD_FI-EV_POSTNG_CTRL = IF_FINS_ACDOC_CONSTANTS=>GC_ENTRY_VIEW_POST_CONTROL-MAX_COMPRESS.
*   For Tax payable posting exclude BUV items
    IF NOT ( P_ACC-TCODE = CL_FDC_POSTING_IMPL=>CO_FDC_S_ACCDOC_HDR_TCODE-GL_TAX_PAYABLE
         AND ACCIT_FI-KTOSL  = IF_FINS_ACDOC_CONSTANTS=>GC_ACCOUNT_DET_GROUP-CROSS_COMPANY ).
      MOVE-CORRESPONDING  LS_ACDOC_ITEM_ATTR_CO_GL TO P_ACC.
      P_ACC-BUZID = IF_FINS_ACDOC_CONSTANTS=>GC_BUZID-COMPRESSED.
    ENDIF.
  ENDIF.
*
* Projektionsstruktur füllen - Kontokorrentzeilen nicht verdichten!
* Anzahlungsanforderungen/Steuerzeilen wegen Zuordnung VBEL2/POSN2
* nicht verdichten
* Anlagenzeilen nicht verdichten (Ausnahme AUAK )
  IF ACCHD_FI-AWTYP EQ 'CTE' AND ACCHD_FI-GLVOR EQ 'RFT1'
  AND P_ACC-KOART EQ CHAR_K AND P_ACC-KTOSL NE 'BUV'
  AND GV_SUMMARIZE_VENDOR_ITEMS CA '12'.
    IF GV_SUMMARIZE_VENDOR_ITEMS = '2'.
      CLEAR P_ACC-MWSKZ.
      CLEAR P_ACC-TAX_COUNTRY.
      P_ACC-TXDAT_FROM = '00000000'.  "initialize TXDAT_FROM
    ENDIF.
    COLLECT P_ACC.
    IF GV_SUMMARIZE_VENDOR_ITEMS = '2'.
      MOVE-CORRESPONDING ACCIT_FI TO ACCPOSNR.
      CLEAR ACCPOSNR-BUZEI.
      ACCPOSNR-BUZEI_INT = SY-TABIX.
      APPEND ACCPOSNR.
      IF P_ACC-SHKZG EQ CHAR_H.
        P_ACC-WRBTR = - P_ACC-WRBTR.
        P_ACC-PSWBT = - P_ACC-PSWBT.
        P_ACC-DMBTR = - P_ACC-DMBTR.
        P_ACC-DMBE2 = - P_ACC-DMBE2.
        P_ACC-DMBE3 = - P_ACC-DMBE3.
      ENDIF.
      CLEAR: P_ACC-BSCHL, P_ACC-SHKZG.
      COLLECT P_ACC.
    ENDIF.
* CE2208: Collect all items of tax payable posting
  ELSEIF ACCHD_FI-AWTYP = IF_FINS_ACDOC_CONSTANTS=>GC_AWTYP-TXPY
      OR P_ACC-TCODE = CL_FDC_POSTING_IMPL=>CO_FDC_S_ACCDOC_HDR_TCODE-GL_TAX_PAYABLE.
    COLLECT P_ACC.
*   Trigger further compression ignoring SHKZG
    IF ACCHD_FI-AWTYP = IF_FINS_ACDOC_CONSTANTS=>GC_AWTYP-TXPY
    AND ( P_ACC-KOART = 'K' OR P_ACC-KTOSL = 'BUV'
*     Include fees xbilk = Space
      OR ( P_ACC-KOART = 'S' AND P_ACC-XBILK = SPACE AND P_ACC-KTOSL = SPACE ) ).
      MOVE-CORRESPONDING ACCIT_FI TO ACCPOSNR.
      CLEAR ACCPOSNR-BUZEI.
      ACCPOSNR-BUZEI_INT = SY-TABIX.
      APPEND ACCPOSNR.
      IF P_ACC-SHKZG EQ CHAR_H.
        P_ACC-WRBTR = - P_ACC-WRBTR.
        P_ACC-PSWBT = - P_ACC-PSWBT.
        P_ACC-DMBTR = - P_ACC-DMBTR.
        P_ACC-DMBE2 = - P_ACC-DMBE2.
        P_ACC-DMBE3 = - P_ACC-DMBE3.
      ENDIF.
      CLEAR: P_ACC-BSCHL, P_ACC-SHKZG.
      COLLECT P_ACC.
    ENDIF.
  ELSEIF P_ACC-KOART EQ CHAR_D
  OR P_ACC-KOART EQ CHAR_K
* direct tax items must not be summarized                     N2109349
  OR ( NOT P_ACC-MWART IS INITIAL AND P_ACC-XAUTO IS INITIAL "N2109349
* all SD (i.e. glvor = SD00) tax items must be summarized     N2203186
       AND P_ACC-GLVOR NE 'SD00' )                          "N2203186
  OR NOT ( ( I_XTXIT IS INITIAL OR GX_GST_USED IS INITIAL ) "n2675869       "India GST changes
    AND     I_EG_VAT IS INITIAL )                           "2538979 :VAT EG: Line item wise tax calculation
  OR ( P_ACC-KOART EQ CHAR_A                                "note 2104969 "n2311523
  AND  P_ACC-AWTYP <> 'AUAK' )                              "n2311523
  OR P_ACC-AWTYP = 'FKKSU'                                  "n3367977
  OR P_ACC-GLVOR EQ 'RFST'
  OR GX_GST_MY IS NOT INITIAL    " Malaysia  GST changes
* sFIN: disable summarization for postings of FIORI applications using FDC service "n2338679
  OR ( P_ACC-TCODE(5) = CL_FDC_POSTING_IMPL=>CO_FDC_S_ACCDOC_HDR_GEN_TCODE-ALL  "n2338679
  AND  P_ACC-BELNR_SENDER IS INITIAL ).                   "note 3480874
*   AND p_acc-buzid <> if_fins_acdoc_constants=>gc_buzid-compressed ).  "CE2308 compress items from payment transactions
    APPEND P_ACC.
  ELSE.
    PERFORM COMPRESS CHANGING COMPRES.
*   CE2402  Summarization of extended OI managed accounts and non open item managed G/L accounts
    IF P_ACC-X_UJ_CLR = ABAP_TRUE
    OR ( P_ACC-XOPVW = ABAP_FALSE
    AND P_ACC-XLGCLR = ABAP_FALSE ).
      PERFORM COMPRESS_UJ_CLR CHANGING COMPRES.
    ENDIF.
    IF COMPRES IS INITIAL.
      APPEND P_ACC.
    ELSE.
      COLLECT P_ACC.
    ENDIF.
  ENDIF.
*  Verknüpfungsstruktur POSNR-BUZEI füllen
  MOVE-CORRESPONDING ACCIT_FI TO ACCPOSNR.
  CLEAR ACCPOSNR-BUZEI.                      "note 423497
* Index der Tabelle P_ACC merken
  ACCPOSNR-BUZEI_INT = SY-TABIX.
  APPEND ACCPOSNR.

ENDFORM.                    "FI_DOCUMENT_COLLECT