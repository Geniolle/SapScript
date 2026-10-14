  METHOD GET_BSEG.

    DATA: LS_BSEG TYPE BSEG, LS_BKPF TYPE BKPF.
    DATA: LS_TRANSACTION TYPE FIAPPT_S_TRANS_CBR.
    DATA :LS_BSEG_VC_VEN TYPE FIAPPT_S_TRANS_CBR.

    CLEAR: LS_BSEG, LS_BKPF, LS_TRANSACTION,LS_BSEG_VC_VEN.

    LS_BKPF = I_BKPF.
    LS_BSEG = I_BSEG.

    READ TABLE E_BSEG_VC_VEN INTO LS_BSEG_VC_VEN WITH KEY BUKRS = LS_BSEG-BUKRS
                                                          BELNR = LS_BSEG-BELNR
                                                          GJAHR = LS_BSEG-GJAHR.
     IF SY-SUBRC = 0 AND LS_BSEG-KTOSL = 'WIT' .
       EXIT.
     ENDIF .

    IF LS_BKPF-XREVERSAL NE '1'.
* xreversal = 1: reversed document will not be processed
      MOVE-CORRESPONDING LS_BSEG TO LS_TRANSACTION.
      LS_TRANSACTION-BLDAT = LS_BKPF-BLDAT.
      LS_TRANSACTION-BUDAT = LS_BKPF-BUDAT.
      LS_TRANSACTION-TCODE = LS_BKPF-TCODE.
      LS_TRANSACTION-WAERS = LS_BKPF-WAERS.
      LS_TRANSACTION-WWERT = LS_BKPF-WWERT.
      LS_TRANSACTION-KURSF = LS_BKPF-KURSF.
*      ls_transaction-kurst = ls_bkpf-kurst.
      LS_TRANSACTION-XREVERSAL = LS_BKPF-XREVERSAL.
      LS_TRANSACTION-STBLG = LS_BKPF-STBLG.
      LS_TRANSACTION-STJAH = LS_BKPF-STJAH.

      IF GT_EXT_ACC_RANGE IS NOT INITIAL
        AND LS_BSEG-HKONT IN GT_EXT_ACC_RANGE
        AND LS_BSEG-KOART = 'S'
        AND LS_BSEG-BELNR NE LS_BSEG-AUGBL.

        LS_TRANSACTION-BANK_ACC_TYPE = 'E'.
        APPEND LS_TRANSACTION TO E_BSEG_GL.

      ELSEIF GT_INT_ACC_RANGE IS NOT INITIAL
        AND LS_BSEG-HKONT IN GT_INT_ACC_RANGE
        AND LS_BSEG-KOART = 'S'
        AND LS_BSEG-BELNR NE LS_BSEG-AUGBL.

        LS_TRANSACTION-BANK_ACC_TYPE = 'I'.
        APPEND LS_TRANSACTION TO E_BSEG_GL.

      ELSEIF GT_OTH_ACC_RANGE IS NOT INITIAL         "note 2288618
        AND LS_BSEG-HKONT IN GT_OTH_ACC_RANGE
        AND LS_BSEG-KOART = 'S'
        AND LS_BSEG-BELNR NE LS_BSEG-AUGBL.

        LS_TRANSACTION-BANK_ACC_TYPE = 'O'.
        IF LS_BSEG-XAUTO IS INITIAL.
          APPEND LS_TRANSACTION TO E_BSEG_GL.
        ELSE.
        IF LS_BSEG-KTOSL = 'ZAH' .
          APPEND LS_TRANSACTION TO E_BSEG_GL.
          ELSE.
           APPEND LS_TRANSACTION TO E_BSEG_REM_ALL.
          ENDIF .
        ENDIF.

      ELSE.
        IF LS_BSEG-BELNR = LS_BSEG-AUGBL.
          IF LS_BSEG-KOART = 'K'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_VEN.
          ELSEIF LS_BSEG-KOART = 'D'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_CUS.
*          ELSEIF ls_bseg-koart = 'S' and ls_bkpf-xreversal = '2'.
*            APPEND ls_transaction TO e_bseg_gl.
          ELSEIF LS_BSEG-KOART = 'S'.
            APPEND LS_TRANSACTION  TO E_BSEG_GL.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_GL.
          ENDIF.
          APPEND LS_TRANSACTION TO E_BSEG_VC.

        ELSEIF LS_BKPF-XREVERSAL = ' '
* 2: reversal document, 1: reversed document, ' ': neither reversal nor reversed document
               AND LS_BKPF-BSTAT = ' '
* bstat = ' ': normal document
* note for down payment request: bstat = 'S' (noted item) and must be excluded from the selection
               AND LS_BSEG-UMSKS = 'A'.
* umsks = 'A': down payments, down payment requests
          IF LS_BSEG-KOART = 'K'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_VEN.
          ELSEIF LS_BSEG-KOART = 'D'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_CUS.
          ENDIF.
          APPEND LS_TRANSACTION TO E_BSEG_VC.

        ELSEIF LS_BKPF-XREVERSAL = '1' AND GV_PERIOD_1ST_RUN NE 'X'.
          IF LS_BSEG-KOART = 'K'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_VEN.
          ELSEIF LS_BSEG-KOART = 'D'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_CUS.
          ELSEIF LS_BSEG-KOART = 'S'.
            APPEND LS_TRANSACTION  TO E_BSEG_VC_GL.
          ENDIF.
          APPEND LS_TRANSACTION TO E_BSEG_VC.

          ELSEIF LS_BKPF-XREVERSAL = '2'.
             APPEND LS_TRANSACTION TO E_BSEG_GL.
        ENDIF.

        APPEND LS_TRANSACTION TO E_BSEG_REM_ALL.
      ENDIF.
    ENDIF.

  ENDMETHOD.