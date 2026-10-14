*&---------------------------------------------------------------------*
*&  Include           RFFOPT_CBR_CHECKS
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Form  CHECK_VAT_REG_NUMBER_COUNTRIES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CHECK_VAT_REG_NUMBER_COUNTRIES .
* checks whether the foreign vendor accounts and customer accounts determined
* do not have an additional VAT registration number in Portugal.
* In such a case the vendor or customer, even physically located in a foreign country
* (foreign: not in Portugal) is considered as a Portuguese business partner and payments for
* those business partners must not be reported.

  DATA: LS_LFA1 TYPE LFA1,
        LS_KNA1 TYPE KNA1,
        LS_KNAS TYPE KNAS,
        LT_KNAS TYPE TABLE OF KNAS,
        LS_LFAS TYPE LFAS,
        LT_LFAS TYPE TABLE OF LFAS,
        LS_ENT  TYPE N.

  LOOP AT GT_LFA1 INTO LS_LFA1.
    IF LS_LFA1-STCEG IS INITIAL .
      SELECT  MANDT LIFNR LAND1 FROM LFAS INTO TABLE LT_LFAS WHERE LIFNR = LS_LFA1-LIFNR.
      DESCRIBE TABLE LT_LFAS LINES LS_ENT.
      IF LS_ENT = 1 .
        READ TABLE LT_LFAS INTO LS_LFAS WITH  KEY  LIFNR = LS_LFA1-LIFNR LAND1 = 'PT'.
        IF SY-SUBRC = 0.
* the vendor is to be considered as a Portuguese vendor even if he is located outside Portugal
* i.e. lfa1-land1 ne 'PT'.
          DELETE TABLE GT_LFA1 FROM LS_LFA1.
        ENDIF.
      ENDIF.
    ENDIF.
    REFRESH :LT_LFAS.
    CLEAR: LS_LFA1,LS_LFA1,LS_ENT.
  ENDLOOP.

  LOOP AT GT_KNA1 INTO LS_KNA1.
    IF LS_KNA1-STCEG IS INITIAL.
      SELECT MANDT KUNNR LAND1 FROM KNAS INTO TABLE LT_KNAS WHERE KUNNR = LS_KNA1-KUNNR.
      DESCRIBE TABLE LT_KNAS LINES LS_ENT.
      IF LS_ENT = 1 .
        READ TABLE LT_KNAS INTO LS_KNAS WITH  KEY KUNNR = LS_KNA1-KUNNR LAND1 = 'PT'.
        IF SY-SUBRC = 0.
* the customer is to be considered as a Portuguese customer even if he is located outside Portugal
* i.e. kna1-land1 ne 'PT'.
          DELETE TABLE GT_KNA1 FROM LS_KNA1.
        ENDIF.
      ENDIF.
    ENDIF.
    REFRESH :LT_KNAS.
    CLEAR: LS_KNAS,LS_KNA1,LS_ENT.
  ENDLOOP.

ENDFORM.