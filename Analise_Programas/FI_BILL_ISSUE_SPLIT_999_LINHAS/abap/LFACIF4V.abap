*
*---------------------------------------------------------------------*
*       FORM COMPRESS                                                 *
*---------------------------------------------------------------------*
*       Verdichtung des FI-Beleges gemäß Einstellung des Kunden       *
*       - Kein Eintrag in TTYPV        -> keine Verdichtung           *
*       - Maskierter Eintrag in TTYPV  -> COLLECT                     *
*       - Spezifizier Eintrag in TTYPV -> Feldinhalt löschen + COLLECT*
*---------------------------------------------------------------------*
FORM COMPRESS CHANGING COMPRES LIKE BOOLE.

  DATA: FIELD(36) TYPE C.
  FIELD-SYMBOLS <F>.
  STATICS: LT_TTYPS           TYPE STANDARD TABLE OF TTYPS.
  DATA: LS_SAKO          TYPE XSAKO,
*        lv_jva_act_mode  TYPE jv_activity_mode,
        LB_FM_NOT_ACTIVE TYPE ABAP_BOOL.
*
  LOOP AT GT_TTYPV WHERE   AWTYP EQ P_ACC-AWTYP           "note 1779136
                   AND   ( BLART EQ P_ACC-BLART           "note 1779136
                   OR      BLART EQ SPACE )               "note 1779136
                   AND   ( BUKRS EQ P_ACC-BUKRS           "note 1779136
                   OR      BUKRS EQ SPACE )               "note 1779136
                   AND     TABNAME NE 'BSET'.             "note 1431727
*
* Feste SAP-Felder (Tabelle TTYPS)
*
    CASE GT_TTYPV-FIELDNAME.                               "note 892315
      WHEN '*    '.
      WHEN 'FKBER'.                                       "note 3425358
        CLEAR: P_ACC-FKBER, P_ACC-FKBER_LONG.             "note 3425358
      WHEN 'WERKS'.
        CLEAR P_ACC-WERKS.
      WHEN 'MATNR'.
        CLEAR P_ACC-MATNR.
      WHEN 'MENGE'.
        CLEAR P_ACC-MENGE.
        CLEAR FIELD.                                      "note 2185343
        FIELD+00(06) = 'P_ACC-'.                          "note 2185343
        FIELD+06(10) = 'SQUAN'.                           "note 2185343
        ASSIGN (FIELD) TO <F>.                            "note 2185343
        IF SY-SUBRC IS INITIAL.                           "note 2185343
          CLEAR <F>.                                      "note 2185343
        ENDIF.                                            "note 2185343
      WHEN 'MEINS'.
        CLEAR P_ACC-MEINS.
      WHEN 'BPMNG'.
        CLEAR P_ACC-BPMNG.
      WHEN 'BPRME'.
        CLEAR P_ACC-BPRME.
      WHEN 'PPRCT'.
        CLEAR P_ACC-PPRCT.
      WHEN OTHERS.
*
* Kundenfelder
*
        CLEAR FIELD.
        FIELD+00(06) = 'P_ACC-'.
        FIELD+06(30) = GT_TTYPV-FIELDNAME.                 "note 892315
        ASSIGN (FIELD) TO <F>.
        IF NOT SY-SUBRC IS INITIAL.
          MESSAGE E843 WITH GT_TTYPV-FIELDNAME 'LFACIGEN' 'SAPFACCG'.
        ENDIF.
        CLEAR <F>.
    ENDCASE.
  ENDLOOP.                                                 "note 892315

*
* Wird überhaupt komprimiert?
*
  IF SY-SUBRC IS INITIAL
* always compress tax items from SD
  OR ( P_ACC-MWART NE SPACE
  AND ( ACCHD_FI-GLVOR EQ 'SD00' OR                           "N1477287
        NOT G_BADI_TAX_COMPRESS IS INITIAL OR                 "N1477287
        ( ACCHD_FI-AWTYP EQ 'CTE' AND ACCHD_FI-GLVOR EQ 'RFT1'
          AND NOT GX_SUMMARIZE_TAX_ITEMS IS INITIAL ) ) ).
    COMPRES = CHAR_X.
  ELSE.
    COMPRES = SPACE.
  ENDIF.

  "2023: badi for max summarization
*  IF lo_fqm IS INITIAL.
*    lo_fqm = NEW #( ).
*  ENDIF.
  CHECK COMPRES = ABAP_TRUE.

  IF ( GB_SUMM_ACCT_ASSIGN = ABAP_TRUE OR GB_SUMM_MAX = ABAP_TRUE )
    "check if line item is tax relevant "or (classically) OI managed
    AND ACCIT_FI-MWART = SPACE "AND accit_fi-xopvw = abap_false AND accit_fi-xlgclr = abap_false
    "check if line item is cash relevant, in this case don't summarize
    AND ( NOT CL_FQMC_ACTIVE=>IS_ACTIVE( IV_ORIGIN_APPLICATION = IF_FQM_ORIGIN_APPLICATION_C=>GC_BSEG-CODE
                                         IV_COMPANY_CODE = ACCIT_FI-BUKRS )
             OR ( ACCIT_FI-FDLEV IS INITIAL AND ACCIT_FI-FDTAG IS INITIAL AND ACCIT_FI-FDGRP IS INITIAL ) )
*                  AND lo_fqm->derive_flow_type( is_input = CORRESPONDING #( p_acc ) ) IS INITIAL
*                  AND lo_fqm->derive_liquidity_item( is_bkpf = CORRESPONDING #( p_acc ) is_bseg = CORRESPONDING #( p_acc ) ) IS INITIAL )
    "exclude BUV items from summarization
    AND ACCIT_FI-KTOSL <> IF_FINS_ACDOC_CONSTANTS=>GC_ACCOUNT_DET_GROUP-CROSS_COMPANY.

    CALL FUNCTION 'FI_GL_ACCOUNT_DATA'
      EXPORTING
        I_BUKRS = ACCIT_FI-BUKRS
        I_SAKNR = ACCIT_FI-HKONT
      IMPORTING
        E_SAKO  = LS_SAKO
      EXCEPTIONS
        OTHERS  = 1.
    IF SY-SUBRC <> 0.
      RETURN.
    ENDIF.
    "for GR/IR with extended OIM, compression is fix (calculated)
    CHECK NOT ( LS_SAKO-X_UJ_CLR = ABAP_TRUE AND
      ( ACCHD_FI-AWTYP = IF_FINS_ACDOC_CONSTANTS=>GC_AWTYP-MKPF OR ACCHD_FI-AWTYP = IF_FINS_ACDOC_CONSTANTS=>GC_AWTYP-RMRP ) ).
*    CALL FUNCTION 'JV_ACTIVE_MODE'
*      EXPORTING
*        iv_bukrs         = accit_fi-bukrs
*      IMPORTING
*        ev_activity_mode = lv_jva_act_mode.

    IF ACCIT_FI-XOPVW = ABAP_FALSE AND ACCIT_FI-XLGCLR = ABAP_FALSE AND LS_SAKO-X_UJ_CLR = ABAP_FALSE.
*      AND lv_jva_act_mode = '0' "no coding block clearing when JVA is active
      "clear all coding block fields, but not for OI managed accounts (like GR/IR account)
      MOVE-CORRESPONDING VALUE FINS_ACDOC_ITEM_ATTR_CO_GL( ) TO P_ACC.
    ENDIF.
    IF GB_SUMM_MAX = ABAP_TRUE.
      "clear all fields from TTYPS
      IF LT_TTYPS IS INITIAL.
        SELECT * FROM TTYPS INTO TABLE LT_TTYPS WHERE TABNAME = 'BSEG'.
*        DATA(lt_comp) = CAST cl_abap_structdescr( cl_abap_typedescr=>describe_by_name( 'FINS_ACDOC_ITEM_ATTR_CO_GL' ) )->components.
*        LOOP AT lt_ttyps REFERENCE INTO DATA(lr_ttyps).
*          IF line_exists( lt_comp[ name = lr_ttyps->fieldname ] ).
*            DELETE lt_ttyps.
*          ENDIF.
*        ENDLOOP.
      ENDIF.
      LOOP AT LT_TTYPS REFERENCE INTO DATA(LR_TTYPS).
*        IF ( p_acc-awtyp = 'MKPF' OR p_acc-awtyp = 'RMRP' ) AND lb_fm_not_active = abap_false.
*          CHECK lr_ttyps->fieldname <> 'XREF3' AND lr_ttyps->fieldname <> 'MENGE' AND lr_ttyps->fieldname <> 'MEINS'
*            AND lr_ttyps->fieldname <> 'EBELN' AND lr_ttyps->fieldname <> 'EBELP' AND lr_ttyps->fieldname <> 'ZEKKN'.
*        ENDIF.
        ASSIGN COMPONENT LR_TTYPS->FIELDNAME OF STRUCTURE P_ACC TO FIELD-SYMBOL(<LV_FIELD>).
        ASSERT SY-SUBRC = 0.
        CLEAR <LV_FIELD>.
      ENDLOOP.
    ENDIF.
    "set BUZID to 'C' for BAPI postings (we need this for potential FB08 reversal later)
    "FIORI postings (which could be reversed by FB08, too), are not running into this subroutine
    IF ACCIT_FI-BAPI_PARAM IS NOT INITIAL.
      P_ACC-BUZID = IF_FINS_ACDOC_CONSTANTS=>GC_BUZID-COMPRESSED.
    ENDIF.

  ENDIF.

ENDFORM.                    "COMPRESS