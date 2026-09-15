*{   INSERT         S4DK933588                                        1

  DATA: BEGIN OF ZUK_901,
          MODUL(3) VALUE '901',
          VTWEG  LIKE VBAK-VTWEG,
          SPART  LIKE VBAK-SPART,
*          vgbel  LIKE vbrp-vgbel,
          BILLNO LIKE TVKO-MAXBI,
        END OF ZUK_901.


  DATA: BEGIN OF J_1B_SIZE_SPLIT_901 OCCURS 0,
          KUNRG         LIKE VBRK-KUNRG,
          KUNAG         LIKE VBRK-KUNAG,
          ZTERM         LIKE VBRK-ZTERM,
          VGBEL         LIKE VBRP-VGBEL.
          INCLUDE STRUCTURE ZUK_901.
  DATA:   ITEMNO TYPE CACS_DAYS.
  DATA: END OF J_1B_SIZE_SPLIT_901.

  CONSTANTS : GC_PROCE TYPE ZCA_PROCE_E VALUE 'ZSD_IC_GROUP',
              GC_FNAME TYPE NAME_FELD VALUE 'MAX_N_LINES'.

  DATA: J_1B_SIZE_COPY_901 LIKE J_1B_SIZE_SPLIT_901.
*}   INSERT
FORM DATEN_KOPIEREN_901.
*{   INSERT         S4DK900493                                        1


  CONSTANTS :  GC_PROC_PR      TYPE ZCA_PROCE_E VALUE 'ZSD_IC_GROUP',
               GC_NAME         TYPE NAME_FELD VALUE 'EKORG'.

  DATA: LR_EKORG TYPE RANGE OF EKORG.

  DATA: BEGIN OF ZUK,
    MODUL(3) VALUE '001',
    VTWEG LIKE VBAK-VTWEG,
    SPART LIKE VBAK-SPART,
  END OF ZUK.

* "variáveis locais
 DATA: LS_KNVV TYPE KNVV.

* Ticket - WSAP-565 - 10.03.2021 - ini
*  IF vbrk-fkart(3) EQ 'ZP0'. "Own Stores POS invoices

 DATA LT_RANGE TYPE RANGE OF VBRK-FKART.
 DATA LV_MAX_LINES(5).
 DATA AUX_TABIX LIKE SY-TABIX.
 DATA LV_GROUP.

* WILL, P. Castro - 04.10.2022 - Invoice Intercompany Grouping , with limit of 2500 lines BEGIN

  CLEAR LV_GROUP.
  SELECT * FROM LIPS INTO @DATA(LV_LIPS) UP TO 1 ROWS
  WHERE VBELN = @LIKP-VBELN.
  ENDSELECT.

  SELECT SINGLE EKORG FROM EKKO INTO @DATA(LV_EKORG)
  WHERE EBELN = @LV_LIPS-VGBEL.

  IF LV_EKORG IS NOT INITIAL.

    CLEAR : LR_EKORG.
    ZCLCA_FIXEDVALS=>GET_CONS_RAN(
       EXPORTING
           IV_BUKRS = SPACE
           IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_SD
           IV_PROCE = GC_PROC_PR
           IV_FNAME = GC_NAME
       IMPORTING
           ET_RANGE = LR_EKORG
       EXCEPTIONS
           NO_DATA  = 1
           OTHERS   = 2
         ).

     IF LR_EKORG IS NOT INITIAL AND LV_EKORG IN LR_EKORG.
       LV_GROUP = ABAP_TRUE.
     ENDIF.
  ENDIF.

  IF LV_GROUP EQ ABAP_TRUE.

    CLEAR LV_MAX_LINES.
    CALL METHOD ZCLCA_FIXEDVALS=>GET_CONS_VAL
          EXPORTING
            IV_BUKRS = SPACE
            IV_MODUL = ZCLCA_FIXEDVALS=>GC_MODULE_SD
            IV_PROCE = GC_PROCE
            IV_FNAME = GC_FNAME
            IV_SEQUE = 1
          IMPORTING
            EV_CONST = LV_MAX_LINES
          EXCEPTIONS
            NO_DATA  = 1
            OTHERS   = 2.
        IF SY-SUBRC <> 0.
*       Implement suitable error handling here
          RETURN.
        ENDIF.


*    CONSTANTS : li_max TYPE maxbi VALUE '2'.

    IF LV_MAX_LINES IS NOT INITIAL.


       READ TABLE J_1B_SIZE_SPLIT_901 WITH KEY KUNRG = VBRK-KUNRG
                                               KUNAG = VBRK-KUNAG
                                               ZTERM = VBRK-ZTERM.
*                                               vgbel = vbrp-vgbel.
        IF SY-SUBRC <> 0.
          CLEAR J_1B_SIZE_SPLIT_901.
          MOVE-CORRESPONDING VBRK TO J_1B_SIZE_SPLIT_901.
          MOVE-CORRESPONDING VBRP TO J_1B_SIZE_SPLIT_901.
        ENDIF.

*     Check number of billing items against max. defined by tvko-maxbi
        IF J_1B_SIZE_SPLIT_901-ITEMNO <   LV_MAX_LINES .
           J_1B_SIZE_SPLIT_901-ITEMNO =  J_1B_SIZE_SPLIT_901-ITEMNO + 1.
        ELSE.
          IF J_1B_SIZE_SPLIT_901-VGBEL NE VBRP-VGBEL.
            J_1B_SIZE_SPLIT_901-BILLNO =  J_1B_SIZE_SPLIT_901-BILLNO + 1.
            J_1B_SIZE_SPLIT_901-ITEMNO = 1.
          ELSE.
            J_1B_SIZE_SPLIT_901-ITEMNO =  J_1B_SIZE_SPLIT_901-ITEMNO + 1.
          ENDIF.
        ENDIF.

*     Store actual billing document counter and item counter
        READ TABLE J_1B_SIZE_SPLIT_901 INTO J_1B_SIZE_COPY_901
           WITH KEY KUNRG = VBRK-KUNRG
                    KUNAG = VBRK-KUNAG
                    ZTERM = VBRK-ZTERM.
*                   vgbel = vbrp-vgbel.
        IF SY-SUBRC = 0.
          J_1B_SIZE_SPLIT_901-VGBEL = VBRP-VGBEL.

          MODIFY J_1B_SIZE_SPLIT_901 TRANSPORTING VGBEL ITEMNO BILLNO
             WHERE KUNRG = VBRK-KUNRG
              AND  KUNAG = VBRK-KUNAG
              AND  ZTERM = VBRK-ZTERM.
*             AND  vgbel = vbrp-vgbel.
        ELSE.
          APPEND J_1B_SIZE_SPLIT_901.
        ENDIF.

*     Add billing doc. number to split criteria
        ZUK_901-BILLNO =  J_1B_SIZE_SPLIT_901-BILLNO.

      ENDIF.
*   End of billing document split by number of allowed items

      ZUK_901-VTWEG = VBAK-VTWEG.
      ZUK_901-SPART = VBRP-SPART.
      "zuk_901-vgbel = vbrp-vgbel.
      VBRK-ZUKRI = ZUK_901.

   ELSE.

      ZUK-SPART = VBAK-SPART.
      ZUK-VTWEG = VBAK-VTWEG.
      VBRK-ZUKRI = ZUK.

   ENDIF.
* WILL, P. Castro - 04.10.2022 - Invoice Intercompany Grouping , with limit of 2500 lines END



* { WILL - 21.10.2021 - JMENDES  Do not run Substitution for industry sales area data:
 DATA: GV_VTWEG          TYPE VBAK-VTWEG.
 CHECK ZCL_FI_PROFIT_SUBST=>GET_INDUSTRY_SALES_AREA( EXPORTING  IV_VTWEG = VBAK-VTWEG ).
 CHECK GV_VTWEG NE VBAK-VTWEG.
* WILL - 21.10.2021 }

*NSS - 31.05.2022 - comentar codigo para determinar PAGADOR e substituir por EMISSOR - beg
*  "Get range
*  call method zclca_fixedvals=>get_cons_ran
*    exporting
*      iv_bukrs = space
*      iv_modul = ztgca_c_mod_fin
*      iv_proce = 'PC_LOJAS'
*      iv_fname = 'FKART'
*    importing
*      et_range = lt_range
*    exceptions
*      no_data  = 1
*      others   = 2.
*
*  if ( lt_range is not initial ) and ( vbrk-fkart in lt_range ). "Own Stores POS invoices
** Ticket - WSAP-565 - 10.03.2021 - fim
*NSS - 31.05.2022 - comentar codigo para determinar PAGADOR e substituir por EMISSOR - end
* Profit Center from Sold-to party - Own Stores POS invoices
    SELECT SINGLE ZZPRCTR
    FROM KNVV
    INTO CORRESPONDING FIELDS OF @LS_KNVV
    WHERE KUNNR EQ @VBRK-KUNAG
    AND VKORG EQ @VBRK-VKORG
    AND VTWEG EQ @VBRK-VTWEG
    AND SPART EQ @VBRK-SPART.
*NSS - 31.05.2022 - comentar codigo para determinar PAGADOR e substituir por EMISSOR - beg
*  ELSE.
*
** Profit Center from Payer
*    SELECT SINGLE zzprctr
*    FROM knvv
*    INTO CORRESPONDING FIELDS OF @ls_knvv
*    WHERE kunnr EQ @vbrk-kunrg
*    AND vkorg EQ @vbrk-vkorg
*    AND vtweg EQ @vbrk-vtweg
*    AND spart EQ @vbrk-spart.
*
*  ENDIF.
*NSS - 31.05.2022 - comentar codigo para determinar PAGADOR e substituir por EMISSOR - end
    IF SY-SUBRC EQ 0.
      " Replace Profit Center
      VBRP-PRCTR = LS_KNVV-ZZPRCTR.
    ENDIF.



*}   INSERT
ENDFORM.