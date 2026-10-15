*&---------------------------------------------------------------------*
*&      Form  SUBST_SINGLE_AWREF_REB
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM SUBST_SINGLE_AWREF_REB TABLES   XBKPF STRUCTURE  BKPF
                                     XBSEG STRUCTURE  BSEG
                            USING    I_AWREF_REB LIKE ACCIT-AWREF_REB
                                     I_AWORG_REB LIKE ACCIT-AWORG_REB
                                     I_AWTYP     LIKE ACCIT-AWTYP
                                     I_AWORG     LIKE ACCIT-AWORG
                                     I_AWSYS     LIKE ACCHD-AWSYS
                                     I_BUKRS     LIKE ACCIT-BUKRS
                                     I_SHKZG     LIKE ACCIT-SHKZG"320615
                                     I_AWSYS_REB LIKE ACCIT-AWSYS_REB "EHP5 ALE Szenario
                            CHANGING E_REBZG     LIKE BSEG-BELNR
                                     E_REBZJ     LIKE BSEG-GJAHR
                                     E_REBZZ     LIKE BSEG-BUZEI
                                     E_SUBRC     TYPE SYSUBRC"427167
                                     SAVE_TABIX  LIKE SY-TABIX. "998689

  DATA: L_SUBRC TYPE SYSUBRC.                               "Note427167
  DATA: L_BELNR LIKE BSEG-BELNR,                            "Note481927
        L_GJAHR LIKE BSEG-GJAHR.                            "Note481927

* Ermittlung des FI-Belegs zu AWREF_REB.
  IF XBKPF-AWTYP NE I_AWTYP
  OR XBKPF-AWKEY(10) NE I_AWREF_REB.
    REFRESH: XBKPF, XBSEG.
    IF I_AWSYS_REB IS INITIAL.
      CALL FUNCTION 'FI_DOCUMENT_READ'
        EXPORTING
          I_AWTYP     = I_AWTYP
          I_AWREF     = I_AWREF_REB
          I_AWORG     = I_AWORG_REB
          I_AWSYS     = I_AWSYS
        TABLES
          T_BKPF      = XBKPF
          T_BSEG      = XBSEG
        EXCEPTIONS
          WRONG_INPUT = 1
          NOT_FOUND   = 2
          OTHERS      = 3.
    ELSE.
      CALL FUNCTION 'FI_DOCUMENT_READ'
        EXPORTING
          I_AWTYP     = I_AWTYP
          I_AWREF     = I_AWREF_REB
          I_AWORG     = I_AWORG_REB
          I_AWSYS     = I_AWSYS_REB
        TABLES
          T_BKPF      = XBKPF
          T_BSEG      = XBSEG
        EXCEPTIONS
          WRONG_INPUT = 1
          NOT_FOUND   = 2
          OTHERS      = 3.
    ENDIF.
    L_SUBRC = SY-SUBRC.                                     "Note427167
* Note 481927: Begin of insertion
    IF L_SUBRC NE 0.
      L_BELNR = I_AWREF_REB.
      L_GJAHR = I_AWORG_REB(4).
      CALL FUNCTION 'FI_DOCUMENT_READ'
        EXPORTING
          I_BUKRS     = I_BUKRS
          I_BELNR     = L_BELNR
          I_GJAHR     = L_GJAHR
        TABLES
          T_BKPF      = XBKPF
          T_BSEG      = XBSEG
        EXCEPTIONS
          WRONG_INPUT = 1
          NOT_FOUND   = 2
          OTHERS      = 3.
      L_SUBRC = SY-SUBRC.
    ENDIF.
* Note 481927: End of insertion
    IF L_SUBRC IS INITIAL.                                "Note 1761056
      SORT XBKPF BY BUKRS GJAHR BELNR.                    "Note 1761056
      SORT XBSEG BY BUKRS GJAHR BELNR BUZEI.              "Note 1761056
    ENDIF.                                                "Note 1761056
  ENDIF.
* (del) if sy-subrc is initial.                             "Note427167
  IF L_SUBRC IS INITIAL.                                    "Note427167
    LOOP AT XBKPF WHERE BUKRS EQ I_BUKRS.
      LOOP AT XBSEG FROM SAVE_TABIX                        "note 998689
                    WHERE BELNR = XBKPF-BELNR
                    AND   GJAHR = XBKPF-GJAHR
                    AND   BUKRS = XBKPF-BUKRS
                    AND   KTOSL NE 'BUV'                    "Note436951
                    AND   VORGN NE 'AZUM'                 "note 1749187
                    AND   KOART = ACCIT_FI-KOART.
        CHECK NOT ( XBKPF-AWTYP = 'RMRP'                  "note 2366639
              AND   XBSEG-KTOSL = 'EGX' ).                "note 2366639
        E_REBZG = XBSEG-BELNR.
        E_REBZJ = XBSEG-GJAHR.
        E_REBZZ = XBSEG-BUZEI.
        SAVE_TABIX = SY-TABIX + 1.                         "note 998689
        EXIT.
      ENDLOOP.
*      if xbseg-shkzg eq i_shkzg.             "note 320615 "note 517513
      IF SY-SUBRC IS INITIAL.                             "note 1141926
        DESCRIBE TABLE XBKPF LINES SY-TFILL.               "note 320615
        IF SY-TFILL GE 2.                                  "note 320615
          IF XBSEG-SHKZG EQ I_SHKZG.                       "note 517513
            IF NOT ( ACCIT_FI-AWREF_REV IS INITIAL         "note 505998
            AND ( XBSEG-KOART EQ CHAR_K                    "note 505998
            AND XBSEG-SHKZG EQ CHAR_H )                    "note 505998
            AND ALE_FLAG IS INITIAL ).                     "note 505998
              CLEAR SAVE_TABIX.                            "note 998689
              CONTINUE.                                    "note 320615
            ENDIF.                                         "note 505998
          ELSE.                                            "note 517513
            IF ACCIT_FI-AWREF_REV IS INITIAL               "note 517513
            AND ( XBSEG-KOART EQ CHAR_K                    "note 517513
            AND XBSEG-SHKZG NE CHAR_H )                    "note 517513
            AND ALE_FLAG IS INITIAL.                       "note 517513
              CLEAR SAVE_TABIX.                            "note 998689
              CONTINUE.                                    "note 517513
            ENDIF.                                         "note 517513
          ENDIF.                                           "note 320615
        ENDIF.                                             "note 320615
      ENDIF.                                              "note 1141926
* Document with invoice reference found
      IF SY-SUBRC EQ SPACE.
        EXIT.
      ENDIF.
    ENDLOOP.
* multiple documents eg. invoices with additional ML document
* invoice reference found ?
    IF  E_REBZG IS INITIAL
    AND E_REBZJ IS INITIAL
    AND E_REBZZ IS INITIAL.
* (del) message e346.                                       "Note427167
      E_SUBRC = 4.                                          "Note427167
    ENDIF.
  ELSE.
* (del) message e346.                                       "Note427167
    E_SUBRC = 8.                                            "Note427167
  ENDIF.

ENDFORM.                               " SUBST_SINGLE_AWREF_REB