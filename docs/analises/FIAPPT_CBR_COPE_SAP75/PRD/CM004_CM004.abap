  METHOD CHECK_OLD_REGISTERS.

    DATA: LS_EX_FIAPPTD_TRANS LIKE IS_FIAPPTD_TRANS.

    CLEAR: EV_REGISTER_ID, EV_REGISTER_EXISTS,
           EV_ORIGIN_RUN_ID.

    SELECT RUN_ID NPC2 SCB_IND PAIS_ATIVO
      DUE_DATE REGISTER_ID ORIGIN_RUN_ID
    INTO CORRESPONDING FIELDS OF LS_EX_FIAPPTD_TRANS
    FROM FIAPPTD_TRANS WHERE BUKRS = GV_BUKRS
      AND BELNR = IS_FIAPPTD_TRANS-BELNR
      AND DOC_YEAR = IS_FIAPPTD_TRANS-DOC_YEAR
      AND BUZEI = IS_FIAPPTD_TRANS-BUZEI
      AND PERIOD = GV_PERIOD
      AND PAY_DOC = IS_FIAPPTD_TRANS-PAY_DOC
      AND PAY_DOC_YEAR = IS_FIAPPTD_TRANS-PAY_DOC_YEAR.
      IF LS_EX_FIAPPTD_TRANS-RUN_ID EQ GV_PREV_RUNID.
        EXIT.
      ENDIF.
    ENDSELECT.
*      AND run_id = gv_prev_runid.

    IF LS_EX_FIAPPTD_TRANS IS NOT INITIAL.   "register found

      IF LS_EX_FIAPPTD_TRANS-NPC2 EQ IS_FIAPPTD_TRANS-NPC2 AND
        LS_EX_FIAPPTD_TRANS-SCB_IND EQ IS_FIAPPTD_TRANS-SCB_IND  AND
        LS_EX_FIAPPTD_TRANS-PAIS_ATIVO EQ IS_FIAPPTD_TRANS-PAIS_ATIVO AND
        LS_EX_FIAPPTD_TRANS-DUE_DATE EQ IS_FIAPPTD_TRANS-DUE_DATE.
*********Register same wrt the 4 fields - npc2, scb, duedate and pais_ativo
        IF IV_XREVERSAL = '1'.        "if the document is reversed.
          IF LS_EX_FIAPPTD_TRANS-REGISTER_ID EQ 'A'.
            EV_REGISTER_EXISTS = 'X'.
          ELSE.               "In the previous reporting,
*      the register was not reported to be reversed
            EV_REGISTER_ID = 'A'.      "So, report the reversal
* in this run run_id of the run when this register was first created.
            EV_ORIGIN_RUN_ID = LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
          ENDIF.
        ELSE.
          EV_REGISTER_EXISTS = 'X'.      "Same register exists
        ENDIF.
      ELSE.                              "Change in register
        EV_REGISTER_ID = 'M'.
        "run_id of the run when this register was first created.
        EV_ORIGIN_RUN_ID = LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
      ENDIF.
    ELSE.                                "register not found
*** BEGIN OF ILM Enablement ***
* Check if register exists in archived data
      READ TABLE GT_AO_FIAPPTD_TRANS WITH KEY BUKRS = GV_BUKRS
                                              BELNR = IS_FIAPPTD_TRANS-BELNR
                                              DOC_YEAR = IS_FIAPPTD_TRANS-DOC_YEAR
                                              BUZEI = IS_FIAPPTD_TRANS-BUZEI
                                              PERIOD = GV_PERIOD
                                              PAY_DOC = IS_FIAPPTD_TRANS-PAY_DOC
                                              PAY_DOC_YEAR = IS_FIAPPTD_TRANS-PAY_DOC_YEAR
                                         INTO LS_EX_FIAPPTD_TRANS.
      IF SY-SUBRC EQ 0. " register found in archived data
        IF LS_EX_FIAPPTD_TRANS-NPC2 EQ IS_FIAPPTD_TRANS-NPC2 AND
          LS_EX_FIAPPTD_TRANS-SCB_IND EQ IS_FIAPPTD_TRANS-SCB_IND  AND
          LS_EX_FIAPPTD_TRANS-PAIS_ATIVO EQ IS_FIAPPTD_TRANS-PAIS_ATIVO AND
          LS_EX_FIAPPTD_TRANS-DUE_DATE EQ IS_FIAPPTD_TRANS-DUE_DATE.
* ********Register same wrt the 4 fields - npc2, scb, duedate and pais_ativo
          IF IV_XREVERSAL = '1'.        "if the document is reversed.
            IF LS_EX_FIAPPTD_TRANS-REGISTER_ID EQ 'A'.
              EV_REGISTER_EXISTS = 'X'.
            ELSE.               "In the previous reporting,
*       the register was not reported to be reversed
              EV_REGISTER_ID = 'A'.      "So, report the reversal
*  in this run run_id of the run when this register was first created.
              EV_ORIGIN_RUN_ID = LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
            ENDIF.
          ELSE.
            EV_REGISTER_EXISTS = 'X'.      "Same register exists
          ENDIF.
        ELSE.                              "Change in register
          EV_REGISTER_ID = 'M'.
          "run_id of the run when this register was first created.
          EV_ORIGIN_RUN_ID = LS_EX_FIAPPTD_TRANS-ORIGIN_RUN_ID.
        ENDIF.
      ELSE. " register not found in archived data
*** END OF ILM Enablement ***
        IF IV_XREVERSAL = '1'.             "payment document reversed
*     and payment has not been reported previously =>dont pick
*     specifying that this register should not be considered.
          EV_REGISTER_EXISTS = 'X'.
        ELSE.
          EV_REGISTER_ID = 'C'.
        ENDIF.
      ENDIF. " ILM Enablement
    ENDIF.

  ENDMETHOD.