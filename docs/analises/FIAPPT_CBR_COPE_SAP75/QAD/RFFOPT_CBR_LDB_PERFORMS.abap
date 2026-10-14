*----------------------------------------------------------------------*
***INCLUDE RFFOPT_CBR_LDB_PERFORMS.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Form  SET_CALLBACK
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_LT_CALLBACK  text
*----------------------------------------------------------------------*
FORM SET_CALLBACK  TABLES LT_CALLBACK.

  DATA: LS_CALLBACK TYPE LDBCB.

  LS_CALLBACK-LDBNODE     = 'BOSG'.
  LS_CALLBACK-GET         = 'X'.
  LS_CALLBACK-CB_PROG     = SY-REPID.
  LS_CALLBACK-CB_FORM     = 'CALLBACK'.
  APPEND LS_CALLBACK TO LT_CALLBACK.

  LS_CALLBACK-LDBNODE     = 'BKPF'.
  LS_CALLBACK-GET         = 'X'.
  LS_CALLBACK-CB_PROG     = SY-REPID.
  LS_CALLBACK-CB_FORM     = 'CALLBACK'.
  APPEND LS_CALLBACK TO LT_CALLBACK.

  LS_CALLBACK-LDBNODE     = 'BSEG'.
  APPEND LS_CALLBACK TO LT_CALLBACK.

ENDFORM.                    " SET_CALLBACK


*&---------------------------------------------------------------------*
*&      Form  callback
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->LV_NODE_NAME  text
*      -->LS_WA         text
*      -->LV_MODE       text
*      -->LV_SELECTED   text
*----------------------------------------------------------------------*
FORM CALLBACK USING LV_NODE_NAME  TYPE LDBN-LDBNODE
                          LS_WA
                          LV_MODE TYPE C
                          LV_SELECTED TYPE C.

  DATA: LS_BSEG TYPE BSEG,
        LS_KNB1 TYPE KNB1,
        LS_LFB1 TYPE LFB1.


  CASE LV_NODE_NAME.

    WHEN 'BKPF'.
      BKPF = LS_WA.

    WHEN 'BSEG'.
      LS_BSEG = LS_WA.
      CALL METHOD OBJ_CBR->GET_BSEG
        EXPORTING
          I_BKPF         = BKPF
          I_BSEG         = LS_BSEG
        IMPORTING
          E_BSEG_VC      = GT_BSEG_PAY_VC
          E_BSEG_GL      = GT_BSEG_PAY_GL
          E_BSEG_VC_VEN  = GT_BSEG_PAY_VC_VEN
          E_BSEG_VC_CUS  = GT_BSEG_PAY_VC_CUS
          E_BSEG_VC_GL   = GT_BSEG_PAY_VC_GL
          E_BSEG_REM_ALL = GT_BSEG_REM_ALL.

    WHEN 'KNA1'.
      KNA1 = LS_WA.

    WHEN 'KNB1'.
      LS_KNB1 = LS_WA.
* If customer is not blocked             " CVP Enablement
      IF LS_KNB1-CVP_XBLCK_B IS INITIAL. " CVP Enablement
        CALL METHOD OBJ_CBR->GET_KNA1
          EXPORTING
            IS_KNA1  = KNA1
            IV_BUKRS = LS_KNB1-BUKRS
          IMPORTING
            ET_KNA1  = GT_KNA1.
      ENDIF.                             " CVP Enablement
    WHEN 'LFA1'.
      LFA1 = LS_WA.
    WHEN 'LFB1'.
      LS_LFB1 = LS_WA.
* If vendor is not blocked               " CVP Enablement
      IF LS_LFB1-CVP_XBLCK_B IS INITIAL. " CVP Enablement
        CALL METHOD OBJ_CBR->GET_LFA1
          EXPORTING
            IS_LFA1  = LFA1
            IV_BUKRS = LS_LFB1-BUKRS
          IMPORTING
            ET_LFA1  = GT_LFA1.
      ENDIF.                              " CVP Enablement
  ENDCASE.

ENDFORM.                    "callback
*&---------------------------------------------------------------------*
*&      Form  SET_SELECTIONS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_PARAMS  text
*----------------------------------------------------------------------*
FORM SET_SELECTIONS  TABLES   LT_PARAMS STRUCTURE RSPARAMS
                              LT_BUDAT_RANGE STRUCTURE FIAPPT_S_DATE_RANGE .
  "Insert correct name for <...>.

  DATA: LS_PARAMS      TYPE RSPARAMS,
        LS_BUDAT_RANGE LIKE LINE OF LT_BUDAT_RANGE,
        LV_YEAR        TYPE BAPI0002_4-FISCAL_YEAR,
        LV_PERIOD      TYPE BAPI0002_4-FISCAL_PERIOD.

  READ TABLE LT_BUDAT_RANGE INTO LS_BUDAT_RANGE INDEX 1.

  CALL FUNCTION 'BAPI_COMPANYCODE_GET_PERIOD'
    EXPORTING
      COMPANYCODEID = P_BUKRS
      POSTING_DATE  = LS_BUDAT_RANGE-LOW
    IMPORTING
      FISCAL_YEAR   = LV_YEAR
      FISCAL_PERIOD = LV_PERIOD.


  LS_PARAMS-SELNAME = 'BR_BUKRS'.
  LS_PARAMS-SIGN = 'I'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = P_BUKRS.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LS_PARAMS-SELNAME = 'BR_GJAHR'.
  LS_PARAMS-SIGN = 'I'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = LV_YEAR.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LS_PARAMS-SELNAME = 'BR_BUDAT'.
  LS_PARAMS-SIGN = 'I'.
  LS_PARAMS-OPTION = LS_BUDAT_RANGE-OPTION.
  LS_PARAMS-LOW = LS_BUDAT_RANGE-LOW.
  LS_PARAMS-HIGH = LS_BUDAT_RANGE-HIGH.
  APPEND LS_PARAMS TO LT_PARAMS.

ENDFORM.                    " SET_SELECTIONS
*&---------------------------------------------------------------------*
*&      Form  SET_CALLBACK_KDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_CALLBACK_KDF  text
*----------------------------------------------------------------------*
FORM SET_CALLBACK_KDF  TABLES   LT_CALLBACK_KDF.
  "Insert correct name for <...>.
  DATA: LS_CALLBACK TYPE LDBCB.

  LS_CALLBACK-LDBNODE     = 'LFA1'.
  LS_CALLBACK-GET         = 'X'.
  LS_CALLBACK-CB_PROG     = SY-REPID.
  LS_CALLBACK-CB_FORM     = 'CALLBACK'.
  APPEND LS_CALLBACK TO LT_CALLBACK_KDF.

  LS_CALLBACK-LDBNODE     = 'LFB1'.
  APPEND LS_CALLBACK TO LT_CALLBACK_KDF.

ENDFORM.                    " SET_CALLBACK_KDF
*&---------------------------------------------------------------------*
*&      Form  SET_SELECTIONS_KDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_SELECTIONS_KDF  text
*----------------------------------------------------------------------*
FORM SET_SELECTIONS_KDF  TABLES LT_PARAMS STRUCTURE RSPARAMS.
  "Insert correct name for <...>.
  DATA: LS_PARAMS TYPE RSPARAMS.

  LS_PARAMS-SELNAME = 'KD_LAND1'.
  LS_PARAMS-SIGN = 'E'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = 'PT'.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LS_PARAMS-SELNAME = 'KD_BUKRS'.
  LS_PARAMS-SIGN = 'I'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = P_BUKRS.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LOOP AT P_VEND.
    LS_PARAMS-SELNAME = 'KD_LIFNR'.
    LS_PARAMS-SIGN = 'I'.
    MOVE-CORRESPONDING P_VEND TO LS_PARAMS .
    APPEND LS_PARAMS  TO LT_PARAMS.
    CLEAR LS_PARAMS.
  ENDLOOP.

ENDFORM.                    " SET_SELECTIONS_KDF
*&---------------------------------------------------------------------*
*&      Form  SET_SELECTIONS_DDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_SELECTIONS  text
*----------------------------------------------------------------------*
FORM SET_SELECTIONS_DDF  TABLES  LT_PARAMS STRUCTURE RSPARAMS.

  DATA: LS_PARAMS TYPE RSPARAMS.

  LS_PARAMS-SELNAME = 'DD_LAND1'.
  LS_PARAMS-SIGN = 'E'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = 'PT'.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LS_PARAMS-SELNAME = 'DD_BUKRS'.
  LS_PARAMS-SIGN = 'I'.
  LS_PARAMS-OPTION = 'EQ'.
  LS_PARAMS-LOW = P_BUKRS.
  APPEND LS_PARAMS TO LT_PARAMS.
  CLEAR LS_PARAMS.

  LOOP AT P_CUST.
    LS_PARAMS-SELNAME = 'DD_KUNNR'.
    LS_PARAMS-SIGN = 'I'.
    MOVE-CORRESPONDING P_CUST TO LS_PARAMS .
    APPEND LS_PARAMS  TO LT_PARAMS.
    CLEAR LS_PARAMS.
  ENDLOOP.


ENDFORM.                    " SET_SELECTIONS_DDF
*&---------------------------------------------------------------------*
*&      Form  SET_CALLBACK_DDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LT_CALLBACK  text
*----------------------------------------------------------------------*
FORM SET_CALLBACK_DDF  TABLES   LT_CALLBACK_DDF.

  DATA: LS_CALLBACK TYPE LDBCB.

  LS_CALLBACK-LDBNODE     = 'KNA1'.
  LS_CALLBACK-GET         = 'X'.
  LS_CALLBACK-CB_PROG     = SY-REPID.
  LS_CALLBACK-CB_FORM     = 'CALLBACK'.
  APPEND LS_CALLBACK TO LT_CALLBACK_DDF.

  LS_CALLBACK-LDBNODE     = 'KNB1'.
  APPEND LS_CALLBACK TO LT_CALLBACK_DDF.

ENDFORM.                    " SET_CALLBACK_DDF
*&---------------------------------------------------------------------*
*&      Form  CALL_LDBS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CALL_LDB_BRF TABLES LT_BUDAT_RANGE STRUCTURE FIAPPT_S_DATE_RANGE.

  DATA: LT_CALLBACK TYPE TABLE OF LDBCB.
  DATA: LT_SELECTIONS TYPE TABLE OF RSPARAMS.

  PERFORM SET_SELECTIONS TABLES LT_SELECTIONS LT_BUDAT_RANGE.
  PERFORM SET_CALLBACK TABLES LT_CALLBACK.

  CALL FUNCTION 'LDB_PROCESS'
    EXPORTING
      LDBNAME           = 'BRF'
*     FIELD_SELECTION   =
    TABLES
      CALLBACK          = LT_CALLBACK
      SELECTIONS        = LT_SELECTIONS
    EXCEPTIONS
      LDB_NOT_REENTRANT = 1
      LDB_INCORRECT     = 2
      OTHERS            = 3.
  IF SY-SUBRC <> 0.
    MESSAGE E008(FIAPPT_CBR).
  ENDIF.

*** BEGIN OF CVP Enablement ***
* This method needs to be called in case the LDB BRF is used
*    CALL METHOD obj_cbr->filter_cvp
*      EXPORTING
*        it_lfa1          = gt_lfa1
*        it_kna1          = gt_kna1
*      CHANGING
*        ct_bseg_vc      = gt_bseg_pay_vc
*        ct_bseg_gl      = gt_bseg_pay_gl
*        ct_bseg_vc_ven  = gt_bseg_pay_vc_ven
*        ct_bseg_vc_cus  = gt_bseg_pay_vc_cus
*        ct_bseg_vc_gl   = gt_bseg_pay_vc_gl
*        ct_bseg_rem_all = gt_bseg_rem_all.
*** END OF CVP Enablement ***

ENDFORM.                    " CALL_LDBS
*&---------------------------------------------------------------------*
*&      Form  CALL_LDB_KDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CALL_LDB_KDF .

  DATA: LT_CALLBACK TYPE TABLE OF LDBCB.
  DATA: LT_SELECTIONS TYPE TABLE OF RSPARAMS.

  PERFORM SET_SELECTIONS_KDF TABLES LT_SELECTIONS.
  PERFORM SET_CALLBACK_KDF TABLES LT_CALLBACK.

  CALL FUNCTION 'LDB_PROCESS'
    EXPORTING
      LDBNAME           = 'KDF'
*     FIELD_SELECTION   =
    TABLES
      CALLBACK          = LT_CALLBACK
      SELECTIONS        = LT_SELECTIONS
    EXCEPTIONS
      LDB_NOT_REENTRANT = 1
      LDB_INCORRECT     = 2
      OTHERS            = 3.
  IF SY-SUBRC <> 0.
    MESSAGE E008(FIAPPT_CBR).
  ENDIF.

ENDFORM.                    " CALL_LDB_KDF
*&---------------------------------------------------------------------*
*&      Form  CALL_LDB_DDF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CALL_LDB_DDF .

  DATA: LT_CALLBACK TYPE TABLE OF LDBCB.
  DATA: LT_SELECTIONS TYPE TABLE OF RSPARAMS.

  PERFORM SET_SELECTIONS_DDF TABLES LT_SELECTIONS.
  PERFORM SET_CALLBACK_DDF TABLES LT_CALLBACK.

  CALL FUNCTION 'LDB_PROCESS'
    EXPORTING
      LDBNAME           = 'DDF'
*     FIELD_SELECTION   =
    TABLES
      CALLBACK          = LT_CALLBACK
      SELECTIONS        = LT_SELECTIONS
    EXCEPTIONS
      LDB_NOT_REENTRANT = 1
      LDB_INCORRECT     = 2
      OTHERS            = 3.
  IF SY-SUBRC <> 0.
    MESSAGE E008(FIAPPT_CBR).
  ENDIF.

ENDFORM.                    " CALL_LDB_DDF