***INCLUDE RV60AFZZ .

************************************************************************
*                                                                      *
* This include is reserved for user modifications                      *
* Forms for invoicing                                                  *
* The name of modification modules should begin with 'ZZ'.             *
*                                                                      *
************************************************************************
*       USEREXIT_ACCOUNT_PREP_KOMKCV                                   *
*       USEREXIT_ACCOUNT_PREP_KOMPCV                                   *
*       USEREXIT_NUMBER_RANGE                                          *
*       USEREXIT_PRICING_PREPARE_TKOMK                                 *
*       USEREXIT_PRICING_PREPARE_TKOMP                                 *
************************************************************************

************************************************************************
*       FORM ZZEXAMPLE                                                *
*---------------------------------------------------------------------*
*       text......................................                    *
*---------------------------------------------------------------------*
*FORM ZZEXAMPLE.

*  ...

*ENDFORM.

*eject

*---------------------------------------------------------------------*
*       FORM USEREXIT_ACCOUNT_PREP_KOMKCV                             *
*---------------------------------------------------------------------*
*       This userexit can be used to move additional fields into the  *
*       communication table which is used for account allocation:     *
*       KOMKCV for header fields                                      *
*       This form is called from form KONTENFINDUNG                   *
*---------------------------------------------------------------------*
FORM USEREXIT_ACCOUNT_PREP_KOMKCV.

*  KOMKCV-zzfield = xxxx-zzfield2.

ENDFORM.
*eject

*---------------------------------------------------------------------*
*       FORM USEREXIT_ACCOUNT_PREP_KOMPCV                             *
*---------------------------------------------------------------------*
*       This userexit can be used to move additional fields into the  *
*       communication table which is used for account allocation:     *
*       KOMPCV for item fields                                        *
*       This form is called from form KONTENFINDUNG                   *
*---------------------------------------------------------------------*
FORM USEREXIT_ACCOUNT_PREP_KOMPCV.
  DATA: LV_VBELN TYPE VBAP-VBELN,
        LV_WERKS TYPE VBAP-WERKS.

  KOMPCV-ZZVBTYP = VBRK-VBTYP.

  SELECT SINGLE BKLAS
    INTO KOMPCV-ZZBKLAS
    FROM MBEW
   WHERE BWKEY = VBRP-WERKS
     AND MATNR = VBRP-MATNR.



* WILL - P.Castro, 23.05.2022 new partner for IDOC INBOUND BEGIN:
  DATA : G_VAR(80) VALUE '(SAPLBD20)IDOC_DATA[]'.
  FIELD-SYMBOLS: <FS_TAB> TYPE ANY TABLE.
  ASSIGN (G_VAR) TO <FS_TAB>.


* WILL - P.Castro, 23.05.2022 new partner for IDOC INBOUND END.

ENDFORM.
*eject

*---------------------------------------------------------------------*
*       FORM USEREXIT_NUMBER_RANGE                                    *
*---------------------------------------------------------------------*
*       This userexit can be used to determine the numberranges for   *
*       the internal document number.                                 *
*       US_RANGE_INTERN - internal number range                       *
*       This form is called from form LV60AU02                        *
*---------------------------------------------------------------------*
FORM USEREXIT_NUMBER_RANGE USING US_RANGE_INTERN.

  DATA: LV_STRING  TYPE STRING,
        LV_STRING2 TYPE STRING,
        LV_RES     TYPE BOOLEAN.

  CALL METHOD ZCLSD_RV60AFZZ=>GET_RANGE(
    EXPORTING
      IS_VBRK         = VBRK
    IMPORTING
      EV_NUMBER_RANGE = US_RANGE_INTERN ).

*** BEG - DEV_292 - Credit note w.r.t. invoice - Return intercompany process; validations on IDoc for IC services
* Requirement 1
  IF XVBRK IS NOT INITIAL AND XVBRP[] IS NOT INITIAL.

    CALL METHOD ZCLSD_RV60AFZZ=>CHANGE_REF_DOC_NO(
      EXPORTING
        IS_VBRK  = XVBRK
        IT_VBRP  = XVBRP[]
      CHANGING
        CV_XBLNR = XVBRK-XBLNR ).

  ENDIF.
*** END - DEV_292 - Credit note w.r.t. invoice - Return intercompany process; validations on IDoc for IC services

* { 21.10.2024 - jmendes - CBR POS Customer NIF (from idoc)
  IF XVBRK IS NOT INITIAL AND XVBRP[] IS NOT INITIAL.
    IF ZCLCA_CBR=>IS_CBR_STORE( IV_WERKS = VBRP-WERKS ).
      ZCLCA_CBR=>CHANGE_INVOICE_VATN( CHANGING CH_VBRK = XVBRK ).
      XVBRK[ 1 ]-STCEG = XVBRK-STCEG.
      VBRK-STCEG = XVBRK-STCEG.
    ENDIF.
  ENDIF.
* } 21.10.2024 - jmendes - CBR Invoice Number

  CONSTANTS: ZTGCA_C_MOD_SD TYPE ZCA_MODUL_E VALUE 'SD'.

  DATA: "l_soci(4), " empresa de Colombia.
    L_CONT          TYPE ZDB_FACT_COL-ZINI_FACT, " Diferencia entre el rango actual y el final.
    L_CONT_C(20),
    LV_RANGO_ACTUAL TYPE I,
    LV_RANGO_FIN    TYPE I,
    L_ZMF2(4), " Tipo de factura.  *1
    L_SICH(4), " botão de guardar en VF01. *1
    L_SAMD(4). " botão de guardar en VF04 facturas colectivamente. *1

*  CALL METHOD zclca_fixedvals=>get_cons_val
*    EXPORTING
*      iv_bukrs = vbrk-bukrs
*      iv_modul = ztgca_c_mod_sd
*      iv_proce = 'ZSD_LOSAN_COLOMBIA'
*      iv_fname = 'AUART'
*      iv_seque = 0
*    IMPORTING
*      ev_const = l_zmf2
*    EXCEPTIONS
*      no_data  = 1
*      OTHERS   = 2.
*
*  IF sy-subrc <> 0.
*    CLEAR: l_zmf2.
*  ENDIF.



  L_SICH(4) = 'SICH'.
  L_SAMD(4) = 'SAMD'.

*  IF sy-subrc EQ 0.
*    SELECT SINGLE auart FROM vbak INTO @DATA(lv_auart)
*      WHERE vbeln = @vbrp-aubel.
*
*  ENDIF.

* IF xvbrk-vbtyp = 'M' AND xvbrk-land1 = 'CO'.
  DATA: LT_FKART TYPE RANGE OF FKART.

  ZCLCA_FIXEDVALS=>GET_CONS_RAN(
    EXPORTING
      IV_BUKRS = VBRK-BUKRS
      IV_MODUL = ZTGCA_C_MOD_SD
      IV_PROCE = 'ZSD_LOSAN_COLOMBIA'
      IV_FNAME = 'FKART'
    IMPORTING
      ET_RANGE = LT_FKART
    EXCEPTIONS
      NO_DATA  = 1
      OTHERS   = 2 ).

  IF SY-SUBRC <> 0.
    CLEAR: LT_FKART[].
  ENDIF.


  IF VBRK-FKART IN LT_FKART AND NOT LT_FKART[] IS INITIAL.

    " Aumento el rango del número de factura(NUM_DIAN) al momento de crear la factura SOLO para Colombia.
    IF ( XVBRK-ZZFACT_COL IS INITIAL AND ( SY-UCOMM = L_SICH OR SY-UCOMM = L_SAMD ) )
      OR ( XVBRK-ZZFACT_COL IS INITIAL AND ( SY-UCOMM = L_SICH OR SY-UCOMM = L_SAMD ) ). " Leperez 25/06/2020
      " Hay ocasiones que los campos de vbak vienen en blanco. Por eso se miran los campos de la xvbrk.
      " Leperez 29/96/2020 - Se añade nuevo codigo de boton para que se genere el rango de facturas.

      " Selecciono todos los registros validos a fecha de hoy.
      SELECT * FROM ZDB_FACT_COL INTO TABLE @DATA(LT_FACT_COL)
          WHERE ZDIAN_EXP <= @SY-DATUM
            AND ZDIAN_VENC >= @SY-DATUM.

      " Recorro la tabla de datos.
      LOOP AT LT_FACT_COL INTO DATA(LE_FACT_COL).

        " Leperez 16/12/2019
        " Se trabaja con numeros, no con texto.
        LV_RANGO_ACTUAL = LE_FACT_COL-ZACT_FACT.
        LV_RANGO_FIN = LE_FACT_COL-ZFIN_FACT.
        " Leperez 16/12/2019

        " Si el registro de la tabla no tiene número de rango inicial, se lo asigno y se modifica la tabla del sistema.
        " A la factura se le asigna también el mismo número.
        IF LE_FACT_COL-ZACT_FACT IS INITIAL.

          LE_FACT_COL-ZACT_FACT = LE_FACT_COL-ZINI_FACT.

          " Leperez 5/4/2019
          " Se eliminan los ceros a la izquierda.
          CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
            EXPORTING
              INPUT  = LE_FACT_COL-ZACT_FACT
            IMPORTING
              OUTPUT = LE_FACT_COL-ZACT_FACT.
          " Leperez 5/4/2019

*          MODIFY zdb_fact_col FROM le_fact_col.
*          xvbrk-zzfact_col = le_fact_col-zact_fact.
*          MOVE le_fact_col-zdian_num TO xvbrk-zdian_num. " Lepereez 9/6/2020 - Se asigna numero de DIAN

          " Si el número actual de rango es inferior al final, incremento en uno el rango inicial. Modifico la tabla del sistema y
          " asigno el número a la factura.
        ELSEIF LV_RANGO_ACTUAL < LV_RANGO_FIN." Leperez 16/12/2019

          LE_FACT_COL-ZACT_FACT = LV_RANGO_ACTUAL + 1." Leperez 16/12/2019

          " Leperez 5/4/2019
          " Se eliminan los ceros a la izquierda.
          CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
            EXPORTING
              INPUT  = LE_FACT_COL-ZACT_FACT
            IMPORTING
              OUTPUT = LE_FACT_COL-ZACT_FACT.
          " Leperez 5/4/2019

*          MODIFY zdb_fact_col FROM le_fact_col.
*          xvbrk-zzfact_col = le_fact_col-zact_fact.
*          MOVE le_fact_col-zdian_num TO xvbrk-zdian_num. " Lepereez 9/6/2020 - Se asigna numero de DIAN

        ENDIF.

        CLEAR: LV_STRING,
               LV_STRING2.
        IF XVBRK-BUKRS IS NOT INITIAL.
          LV_STRING = XVBRK-BUKRS.
          CONCATENATE 'Empresa' LV_STRING INTO LV_STRING SEPARATED BY SPACE.
          CONCATENATE LV_STRING '.' INTO LV_STRING.
        ENDIF.
        LV_STRING2 = LV_STRING.

        " Calculo la diferencia entre el número de rango inicial y final.
        " Si es igual o menor a 200 y distinto de 0, informo de los números de rango que quedan por pantalla.
        L_CONT = LV_RANGO_FIN - LV_RANGO_ACTUAL.
        IF  L_CONT =< 200 AND L_CONT <> 0.
          " Aviso - Quedan menos de 200 números de factura disponibles.
          IF L_CONT < 0.
            L_CONT = L_CONT * -1.
          ENDIF.
          L_CONT_C = L_CONT.
          CONDENSE L_CONT_C NO-GAPS.

          IF LV_STRING IS NOT INITIAL.
            CONCATENATE LV_STRING 'Restam' INTO LV_STRING SEPARATED BY SPACE.
          ELSE.
            LV_STRING = 'Restam'.
          ENDIF.

          MESSAGE I398(00) WITH
           LV_STRING L_CONT_C ' números de factura disponiveis. '
           'Verifique os parâmetros de Autorizacão Facturação.'.
        ENDIF.

        " Si queda menos de un mes para el vencimiento, se informa de ello por pantalla.
        IF  LE_FACT_COL-ZDIAN_VENC - SY-DATUM < 30.
          " Aviso - Queda menos de un mes para que cumpla el vencimiento.
          MESSAGE I398(00) WITH
           LV_STRING2
           'Resta menos de um mes para que se cumpra '
           'o vencimento do Nº DIAN. '
           'Verifique os parâmetros de Autorização Facturação.'.
        ENDIF.

        CLEAR: LV_RANGO_ACTUAL, LV_RANGO_FIN, L_CONT.

      ENDLOOP.

    ENDIF.
  ELSE.
    CLEAR: XVBRK-ZZFACT_COL, XVBRK-ZZDIAN_NUM, VBRK-ZZFACT_COL, VBRK-ZZDIAN_NUM.
  ENDIF.



  " **********************************************************************
  " USER: NSO - Nuno Santos Oliveira - Abaco Consulting
  " DATE: 19.04.2023 10:46:09
  " DESCRIPTION: Alerta de validade para o México
  " **********************************************************************
  " {

  IF XVBRK IS NOT INITIAL.

    CALL FUNCTION 'ZSD_MEXICO_DATE_PERIOD_VAL'
      EXPORTING
        I_VKORG                  = XVBRK-VKORG
        I_DISPLAY_POPUP_MESSAGE  = ABAP_TRUE
        I_IGNORE_CONFIG_USERS    = ABAP_TRUE
        I_CREATE_INBOX_MESSAGE   = ABAP_FALSE
        I_VALIDATE_COLOMBIA_DIAN = ABAP_FALSE.

  ENDIF.
  " }
  " **********************************************************************


  " MPereira 09.07.2024 16:21:53
  CALL FUNCTION 'ZEI_VALIDATE_REVERSE'
    EXPORTING
      IV_VBRK   = XVBRK
    IMPORTING
      EV_RESULT = LV_RES.

  IF LV_RES = ABAP_FALSE.
    LEAVE PROGRAM.
  ENDIF.
  " END MPereira 09.07.2024 16:21:53

ENDFORM.
*eject

*---------------------------------------------------------------------*
*       FORM USEREXIT_PRICING_PREPARE_TKOMK                           *
*---------------------------------------------------------------------*
*       This userexit can be used to move additional fields into the  *
*       communication table which is used for pricing:                *
*       TKOMK for header fields                                       *
*       This form is called from form PREISFINDUNG_VORBEREITEN.       *
*---------------------------------------------------------------------*
FORM USEREXIT_PRICING_PREPARE_TKOMK.

* TKOMK-zzfield = xxxx-zzfield2.

* TKOMK-KUNRE = XVBPA_RE-KUNNR.
* TKOMK-KUNWE = XVBPA_WE-KUNNR.
* TKOMK-KNRZE = XVBPA_RG-KUNNR.

* PERFORM XVBPA_SELECT USING 'VE'.
* TKOMK-VRTNR = XVBPA-PERNR.

* PERFORM XVBPA_SELECT USING 'SP'.
* TKOMK-SPDNR = XVBPA-LIFNR.

* PERFORM XVBPA_SELECT USING 'AP'.
* TKOMK-PARNR = XVBPA-PARNR.


  " **********************************************************************
  " USER: NSO - Nuno Santos Oliveira - Abaco Consulting
  " DATE: 27.10.2023 10:15:53
  " DESCRIPTION: Adjust the WE parter in the following conditions:
  " - VBRK-FKART - Constant table - Process 'ZALANDO_IC_WE' Field 'FKART';
  " - VBRK-VKORG - Constant table - Process 'ZALANDO_IC_WE' Field 'VKORG';
  " - VBRK-KUNWE - Constant table - Process 'ZALANDO_IC_WE' Field 'KUNWE';
  " **********************************************************************
  " {

  CONSTANTS: LC_PROCESS_ZALANDO_IC_WE TYPE ZCA_PROCE_E VALUE 'ZALANDO_IC_WE',
             LC_FIELD_FKART           TYPE NAME_FELD VALUE 'FKART',
             LC_FIELD_VKORG           TYPE NAME_FELD VALUE 'VKORG',
             LC_FIELD_KUNWE           TYPE NAME_FELD VALUE 'KUNWE'.

  DATA: LR_ZALANDO_IC_WE_FKART TYPE RANGE OF FKART,
        LR_ZALANDO_IC_WE_VKORG TYPE RANGE OF VKORG,
        LR_ZALANDO_IC_WE_KUNNR TYPE RANGE OF KUNNR.

  CLEAR LR_ZALANDO_IC_WE_FKART.
  ZCLCA_FIXEDVALS=>GET_CONS_RAN( EXPORTING
                                  IV_BUKRS = TKOMK-BUKRS
                                  IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_SD
                                  IV_PROCE = LC_PROCESS_ZALANDO_IC_WE
                                  IV_FNAME = LC_FIELD_FKART
                                 IMPORTING
                                  ET_RANGE = LR_ZALANDO_IC_WE_FKART
                                 EXCEPTIONS
                                  NO_DATA  = 1
                                  OTHERS   = 2 ).

  CLEAR LR_ZALANDO_IC_WE_VKORG.
  ZCLCA_FIXEDVALS=>GET_CONS_RAN( EXPORTING
                                  IV_BUKRS = TKOMK-BUKRS
                                  IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_SD
                                  IV_PROCE = LC_PROCESS_ZALANDO_IC_WE
                                  IV_FNAME = LC_FIELD_VKORG
                                 IMPORTING
                                  ET_RANGE = LR_ZALANDO_IC_WE_VKORG
                                 EXCEPTIONS
                                  NO_DATA  = 1
                                  OTHERS   = 2 ).

  CLEAR LR_ZALANDO_IC_WE_KUNNR.
  ZCLCA_FIXEDVALS=>GET_CONS_RAN( EXPORTING
                                  IV_BUKRS = TKOMK-BUKRS
                                  IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_SD
                                  IV_PROCE = LC_PROCESS_ZALANDO_IC_WE
                                  IV_FNAME = LC_FIELD_KUNWE
                                 IMPORTING
                                  ET_RANGE = LR_ZALANDO_IC_WE_KUNNR
                                 EXCEPTIONS
                                  NO_DATA  = 1
                                  OTHERS   = 2 ).

  IF ( LR_ZALANDO_IC_WE_FKART[] IS NOT INITIAL AND TKOMK-FKART IN LR_ZALANDO_IC_WE_FKART )
     AND ( LR_ZALANDO_IC_WE_VKORG[] IS NOT INITIAL AND TKOMK-VKORG IN LR_ZALANDO_IC_WE_VKORG )
     AND ( LR_ZALANDO_IC_WE_KUNNR[] IS NOT INITIAL AND XVBPA_WE_AUFT-PARVW = 'WE' AND XVBPA_WE_AUFT-KUNNR IN LR_ZALANDO_IC_WE_KUNNR )
     AND ( TKOMK-KUNWE IS NOT INITIAL AND TKOMK-KUNWE <>  XVBPA_WE_AUFT-KUNNR )
     AND ( VBRK-VBTYP IS NOT INITIAL AND CL_SD_DOC_CATEGORY_UTIL=>IS_ANY_INTERCOMPANY( VBRK-VBTYP ) = ABAP_TRUE ).

    TKOMK-KUNWE = XVBPA_WE_AUFT-KUNNR.

  ENDIF.

  " }
  " **********************************************************************


  CALL METHOD ZCLSD_RV60AFZZ=>PRICING_PREPARE_TKOMK(
    EXPORTING
      CS_VBAK  = VBAK
      IT_VBPA  = XVBPA[]
    CHANGING
      CS_TKOMK = TKOMK ).

*  CONSTANTS: ztgca_c_mod_sd TYPE zca_modul_e VALUE 'SD'.
*
*  DATA: "l_soci(4), " empresa de Colombia.
*    l_cont          TYPE zdb_fact_col-zini_fact, " Diferencia entre el rango actual y el final.
*    l_cont_c(20),
*    lv_rango_actual TYPE i,
*    lv_rango_fin    TYPE i,
*    l_zmf2(4), " Tipo de factura.  *1
*    l_sich(4), " botão de guardar en VF01. *1
*    l_samd(4). " botão de guardar en VF04 facturas colectivamente. *1
*
**  CALL METHOD zclca_fixedvals=>get_cons_val
**    EXPORTING
**      iv_bukrs = vbrk-bukrs
**      iv_modul = ztgca_c_mod_sd
**      iv_proce = 'ZSD_LOSAN_COLOMBIA'
**      iv_fname = 'AUART'
**      iv_seque = 0
**    IMPORTING
**      ev_const = l_zmf2
**    EXCEPTIONS
**      no_data  = 1
**      OTHERS   = 2.
**
**  IF sy-subrc <> 0.
**    CLEAR: l_zmf2.
**  ENDIF.
*
*
*
*  l_sich(4) = 'SICH'.
*  l_samd(4) = 'SAMD'.
*
**  IF sy-subrc EQ 0.
**    SELECT SINGLE auart FROM vbak INTO @DATA(lv_auart)
**      WHERE vbeln = @vbrp-aubel.
**
**  ENDIF.
*
** IF xvbrk-vbtyp = 'M' AND xvbrk-land1 = 'CO'.
*  DATA: lt_fkart TYPE RANGE OF fkart.
*
*  zclca_fixedvals=>get_cons_ran(
*    EXPORTING
*      iv_bukrs = vbrk-bukrs
*      iv_modul = ztgca_c_mod_sd
*      iv_proce = 'ZSD_LOSAN_COLOMBIA'
*      iv_fname = 'FKART'
*    IMPORTING
*      et_range = lt_fkart
*    EXCEPTIONS
*      no_data  = 1
*      OTHERS   = 2 ).
*
*  IF sy-subrc <> 0.
*    CLEAR: lt_fkart[].
*  ENDIF.
*
*
*  IF vbrk-fkart IN lt_fkart AND NOT lt_fkart[] IS INITIAL.
*
*    " Aumento el rango del número de factura(NUM_DIAN) al momento de crear la factura SOLO para Colombia.
*    IF ( xvbrk-zzfact_col IS INITIAL AND ( sy-ucomm = l_sich OR sy-ucomm = l_samd ) )
*      OR ( xvbrk-zzfact_col IS INITIAL AND ( sy-ucomm = l_sich OR sy-ucomm = l_samd ) ). " Leperez 25/06/2020
*      " Hay ocasiones que los campos de vbak vienen en blanco. Por eso se miran los campos de la xvbrk.
*      " Leperez 29/96/2020 - Se añade nuevo codigo de boton para que se genere el rango de facturas.
*
*      " Selecciono todos los registros validos a fecha de hoy.
*      SELECT * FROM zdb_fact_col INTO TABLE @DATA(lt_fact_col)
*          WHERE zdian_exp <= @sy-datum
*            AND zdian_venc >= @sy-datum.
*
*      " Recorro la tabla de datos.
*      LOOP AT lt_fact_col INTO DATA(le_fact_col).
*
*        " Leperez 16/12/2019
*        " Se trabaja con numeros, no con texto.
*        lv_rango_actual = le_fact_col-zact_fact.
*        lv_rango_fin = le_fact_col-zfin_fact.
*        " Leperez 16/12/2019
*
*        " Si el registro de la tabla no tiene número de rango inicial, se lo asigno y se modifica la tabla del sistema.
*        " A la factura se le asigna también el mismo número.
*        IF le_fact_col-zact_fact IS INITIAL.
*
*          le_fact_col-zact_fact = le_fact_col-zini_fact.
*
*          " Leperez 5/4/2019
*          " Se eliminan los ceros a la izquierda.
*          CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
*            EXPORTING
*              input  = le_fact_col-zact_fact
*            IMPORTING
*              output = le_fact_col-zact_fact.
*          " Leperez 5/4/2019
*
**          MODIFY zdb_fact_col FROM le_fact_col.
**          xvbrk-zzfact_col = le_fact_col-zact_fact.
**          MOVE le_fact_col-zdian_num TO xvbrk-zdian_num. " Lepereez 9/6/2020 - Se asigna numero de DIAN
*
*          " Si el número actual de rango es inferior al final, incremento en uno el rango inicial. Modifico la tabla del sistema y
*          " asigno el número a la factura.
*        ELSEIF lv_rango_actual < lv_rango_fin." Leperez 16/12/2019
*
*          le_fact_col-zact_fact = lv_rango_actual + 1." Leperez 16/12/2019
*
*          " Leperez 5/4/2019
*          " Se eliminan los ceros a la izquierda.
*          CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
*            EXPORTING
*              input  = le_fact_col-zact_fact
*            IMPORTING
*              output = le_fact_col-zact_fact.
*          " Leperez 5/4/2019
*
**          MODIFY zdb_fact_col FROM le_fact_col.
**          xvbrk-zzfact_col = le_fact_col-zact_fact.
**          MOVE le_fact_col-zdian_num TO xvbrk-zdian_num. " Lepereez 9/6/2020 - Se asigna numero de DIAN
*
*        ENDIF.
*
*        " Calculo la diferencia entre el número de rango inicial y final.
*        " Si es igual o menor a 200 y distinto de 0, informo de los números de rango que quedan por pantalla.
*        l_cont = lv_rango_fin - lv_rango_actual.
*        IF  l_cont =< 200 AND l_cont <> 0.
*          " Aviso - Quedan menos de 200 números de factura disponibles.
*          IF l_cont < 0.
*            l_cont = l_cont * -1.
*          ENDIF.
*          l_cont_c = l_cont.
*          CONDENSE l_cont_c NO-GAPS.
*          MESSAGE i398(00) WITH
*           'Restam ' l_cont_c ' números de factura disponiveis. '
*           'Verifique os parâmetros de Autorizacão Facturação.'.
*        ENDIF.
*
*        " Si queda menos de un mes para el vencimiento, se informa de ello por pantalla.
*        IF  le_fact_col-zdian_venc - sy-datum < 30.
*          " Aviso - Queda menos de un mes para que cumpla el vencimiento.
*          MESSAGE i398(00) WITH
*           'Resta menos de um mes para que se cumpra '
*           'o vencimento do Nº DIAN. '
*           'Verifique os parâmetros de Autorização Facturação.'.
*        ENDIF.
*
*        CLEAR: lv_rango_actual, lv_rango_fin, l_cont.
*
*      ENDLOOP.
*
*    ENDIF.
*  ELSE.
*    CLEAR: xvbrk-zzfact_col, xvbrk-Zzdian_num, vbrk-zzfact_col, vbrk-Zzdian_num.
*  ENDIF.

  " **********************************************************************
  " [NOLIVEIRA - Abaco Consulting] [22.12.2022 17:15:55]
  " DESCRIPTION: 2º level consignment
  " **********************************************************************
  " {
  IF VBRK-KUNAG IS NOT INITIAL.

    DATA: LV_2LEVELCONSIG_KUNAG        TYPE VBRK-KUNAG,
          LV_2LEVELCONSIG_WERKS        TYPE WERKS_D,
          LV_2LEVELCONSIG_COMMISSION   TYPE Z_SD_ECOMMISSIONFRANCHISEE,
          LT_2LEVELCONSIG_CONST        TYPE STANDARD TABLE OF ZCA_CONSTANTS_T,
          LS_2LEVELCONSIG_CONST        TYPE ZCA_CONSTANTS_T,
          LV_2LEVELCONSIG_FKART        TYPE FKART,
          LV_2LEVELCONSIG_FKART_EXISTS TYPE BOOLEAN.

    DATA: LC_2_LEVEL_CONSIG_PROCESS  TYPE ZCA_PROCE_E VALUE '2_LEVEL_CONSIGNMENT'.

    LV_2LEVELCONSIG_FKART_EXISTS = ABAP_FALSE.

    CLEAR LT_2LEVELCONSIG_CONST.
    SELECT
          *
    FROM
          ZCA_CONSTANTS_T
          INTO CORRESPONDING FIELDS OF TABLE LT_2LEVELCONSIG_CONST
    WHERE
          MODUL = 'SD'
          AND PROCE = LC_2_LEVEL_CONSIG_PROCESS
          AND FNAME LIKE 'BILL_DOC_TYPE_%'.

    LOOP AT LT_2LEVELCONSIG_CONST INTO LS_2LEVELCONSIG_CONST.

      CLEAR LV_2LEVELCONSIG_FKART.
      MOVE LS_2LEVELCONSIG_CONST-LOW TO LV_2LEVELCONSIG_FKART.

      IF VBRK-FKART = LV_2LEVELCONSIG_FKART.
        LV_2LEVELCONSIG_FKART_EXISTS = ABAP_TRUE.
        EXIT.
      ENDIF.

    ENDLOOP.

    IF  LV_2LEVELCONSIG_FKART_EXISTS = ABAP_TRUE.

      IF TKOMK-VKORG <> VBRK-VKORG.
        TKOMK-VKORG = VBRK-VKORG.
      ENDIF.

      IF TKOMK-VTWEG <> VBRK-VTWEG.
        TKOMK-VTWEG = VBRK-VTWEG.
      ENDIF.

    ENDIF.

  ENDIF.
  " }
  " **********************************************************************

*  " **********************************************************************
*  " USER: NSO - Nuno Santos Oliveira - Abaco Consulting
*  " DATE: 26.07.2023 11:29:47
*  " DESCRIPTION: Free of Charge
*  " Deactivate discount R100
*  " **********************************************************************
*  " {
*
*  CONSTANTS: lc_process_invoice_form    TYPE zca_proce_e VALUE 'INVOICE_FORM',
*             lc_field_free_charge       TYPE name_feld VALUE 'FREE_CHARGE',
*             lc_field_100_discount_cond TYPE name_feld VALUE 'FREE_CHARGE_100_DISCOUNT_KSCHL'.
*
*  DATA: lr_auart_sd           TYPE RANGE OF auart,
*        lr_kschl_100_discount TYPE RANGE OF kschl,
*        lv_sales_doc_type     TYPE auart.
*
*  IF vbrp-aubel IS NOT INITIAL.
*
*    CLEAR lr_auart_sd.
*    zclca_fixedvals=>get_cons_ran(
*    EXPORTING
*      iv_bukrs = space
*      iv_modul = zclca_fixedvals=>gc_module_sd
*      iv_proce = lc_process_invoice_form
*      iv_fname = lc_field_free_charge
*    IMPORTING
*      et_range = lr_auart_sd
*    EXCEPTIONS
*      no_data  = 1  ).
*
*    CLEAR lr_kschl_100_discount.
*    zclca_fixedvals=>get_cons_ran(
*    EXPORTING
*      iv_bukrs = space
*      iv_modul = zclca_fixedvals=>gc_module_sd
*      iv_proce = lc_process_invoice_form
*      iv_fname = lc_field_100_discount_cond
*    IMPORTING
*      et_range = lr_kschl_100_discount
*    EXCEPTIONS
*      no_data  = 1  ).
*
*    IF lr_auart_sd[] IS NOT INITIAL AND lr_kschl_100_discount[] IS NOT INITIAL.
*
*      CLEAR lv_sales_doc_type.
*      SELECT
*            SINGLE
*            auart
*            INTO lv_sales_doc_type
*      FROM
*            vbak
*      WHERE
*            vbeln = vbrp-aubel.
*
*      IF lv_sales_doc_type IS NOT INITIAL AND lv_sales_doc_type IN lr_auart_sd.
*
*        LOOP AT xkomv ASSIGNING FIELD-SYMBOL(<fs_xkomv_100_discount>) WHERE kschl IN lr_kschl_100_discount .
*          <fs_xkomv_100_discount>-kinak = 'X'.
*        ENDLOOP.
*
*      ENDIF.
*
*    ENDIF.
*
*  ENDIF.
*  " **********************************************************************
*
*  " }
*  " **********************************************************************


ENDFORM.
*eject

*---------------------------------------------------------------------*
*       FORM USEREXIT_PRICING_PREPARE_TKOMP                           *
*---------------------------------------------------------------------*
*       This userexit can be used to move additional fields into the  *
*       communication table which is used for pricing:                *
*       TKOMP for item fields                                         *
*       This form is called from form PREISFINDUNG_VORBEREITEN.       *
*---------------------------------------------------------------------*
FORM USEREXIT_PRICING_PREPARE_TKOMP.

*  TKOMP-zzfield = xxxx-zzfield2.

  CALL METHOD ZCLSD_RV60AFZZ=>PRICING_PREPARE_TKOMP(
    EXPORTING
      CS_VBRP  = VBRP
    CHANGING
      CS_TKOMP = TKOMP ).

  " **********************************************************************
  " [NOLIVEIRA - Abaco Consulting] [28.09.2022 14:32:40]
  " DESCRIPTION: Losan - Retention Taxes - Columbia
  " **********************************************************************
  " {

  CALL FUNCTION 'ZSD_WITHHOLDING_TAX_CODE'
    EXPORTING
      I_BUKRS  = VBRK-BUKRS
      I_KUNNR  = VBRK-KUNRG
    CHANGING
      CS_TKOMP = TKOMP.

  " }
  " **********************************************************************

  " **********************************************************************
  " [NOLIVEIRA - Abaco Consulting] [22.12.2022 17:15:55]
  " DESCRIPTION: 2º level consignment
  " **********************************************************************
  " {
  IF VBRK-KUNAG IS NOT INITIAL.

    DATA: LV_2LEVELCONSIG_KUNAG        TYPE VBRK-KUNAG,
          LV_2LEVELCONSIG_WERKS        TYPE WERKS_D,
          LV_2LEVELCONSIG_COMMISSION   TYPE Z_SD_ECOMMISSIONFRANCHISEE,
          LT_2LEVELCONSIG_CONST        TYPE STANDARD TABLE OF ZCA_CONSTANTS_T,
          LS_2LEVELCONSIG_CONST        TYPE ZCA_CONSTANTS_T,
          LV_2LEVELCONSIG_FKART        TYPE FKART,
          LV_2LEVELCONSIG_FKART_EXISTS TYPE BOOLEAN.

    DATA: LC_2_LEVEL_CONSIG_PROCESS  TYPE ZCA_PROCE_E VALUE '2_LEVEL_CONSIGNMENT'.

    LV_2LEVELCONSIG_FKART_EXISTS = ABAP_FALSE.

    CLEAR LT_2LEVELCONSIG_CONST.
    SELECT
          *
    FROM
          ZCA_CONSTANTS_T
          INTO CORRESPONDING FIELDS OF TABLE LT_2LEVELCONSIG_CONST
    WHERE
          MODUL = 'SD'
          AND PROCE = LC_2_LEVEL_CONSIG_PROCESS
          AND FNAME LIKE 'BILL_DOC_TYPE_%'.

    LOOP AT LT_2LEVELCONSIG_CONST INTO LS_2LEVELCONSIG_CONST.

      CLEAR LV_2LEVELCONSIG_FKART.
      MOVE LS_2LEVELCONSIG_CONST-LOW TO LV_2LEVELCONSIG_FKART.

      IF VBRK-FKART = LV_2LEVELCONSIG_FKART.
        LV_2LEVELCONSIG_FKART_EXISTS = ABAP_TRUE.
        EXIT.
      ENDIF.

    ENDLOOP.

    IF  LV_2LEVELCONSIG_FKART_EXISTS = ABAP_TRUE.

      CLEAR LV_2LEVELCONSIG_KUNAG.
      LV_2LEVELCONSIG_KUNAG = VBRK-KUNAG.

      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
        EXPORTING
          INPUT  = LV_2LEVELCONSIG_KUNAG
        IMPORTING
          OUTPUT = LV_2LEVELCONSIG_KUNAG.

      CLEAR LV_2LEVELCONSIG_WERKS.
      MOVE LV_2LEVELCONSIG_KUNAG TO LV_2LEVELCONSIG_WERKS.

      IF  LV_2LEVELCONSIG_WERKS IS NOT INITIAL.

        CLEAR LV_2LEVELCONSIG_COMMISSION.
        SELECT
              SINGLE
              ZZCOMMISSION
              INTO LV_2LEVELCONSIG_COMMISSION
        FROM
              T001W
        WHERE
              WERKS = LV_2LEVELCONSIG_WERKS.

        IF LV_2LEVELCONSIG_COMMISSION IS NOT INITIAL.
          MOVE LV_2LEVELCONSIG_COMMISSION TO TKOMP-ZZCOMMISSION.
        ENDIF.

      ENDIF.

    ENDIF.

  ENDIF.

  " }
  " **********************************************************************

* P.Castro 09.01.2023 - WS-51 - validate if transporting costs have been charged - BEGIN:

*   IF vbrk-bukrs = '2010' AND vbrk-FKART = 'ZI01'.
*
*          MESSAGE i398(00) WITH
*           'Não foi faturado o item de transporte existente na ordem'
*            vbrp-vgbel.
*
*
*   ENDIF.

* P.Castro 09.01.2023 - WS-51 - validate if transporting costs have been charged - END.

ENDFORM.
*eject