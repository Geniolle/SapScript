  METHOD CHECK_OLD_BAL_REGISTERS.

    DATA: LS_BALANCE TYPE FIAPPTD_BALANCE.
    CLEAR: EV_REG_EXISTS, EV_REG_CHANGED, EV_ORIGIN_RUNID.

    SELECT * FROM FIAPPTD_BALANCE INTO
    CORRESPONDING FIELDS OF LS_BALANCE WHERE
    BUKRS = IS_FIAPPTD_BALANCE-BUKRS
    AND REPORTING_YEAR = IS_FIAPPTD_BALANCE-REPORTING_YEAR
    AND PERIOD = IS_FIAPPTD_BALANCE-PERIOD
*      AND run_id = gv_prev_runid
    AND ACCOUNT = IS_FIAPPTD_BALANCE-ACCOUNT
    AND KOART = IS_FIAPPTD_BALANCE-KOART
    AND WAERS = IS_FIAPPTD_BALANCE-WAERS.
      IF LS_BALANCE-RUN_ID EQ GV_PREV_RUNID.
        EXIT.
      ENDIF.
    ENDSELECT.

    IF LS_BALANCE IS INITIAL.
*** BEGIN OF ILM Enablement ***
* Check if register exists in archived data
      READ TABLE GT_AO_FIAPPTD_BALANCE WITH KEY BUKRS = IS_FIAPPTD_BALANCE-BUKRS
                                                REPORTING_YEAR = IS_FIAPPTD_BALANCE-REPORTING_YEAR
                                                PERIOD = IS_FIAPPTD_BALANCE-PERIOD
                                                ACCOUNT = IS_FIAPPTD_BALANCE-ACCOUNT
                                                KOART = IS_FIAPPTD_BALANCE-KOART
                                                WAERS = IS_FIAPPTD_BALANCE-WAERS
                                           INTO LS_BALANCE.
      IF SY-SUBRC NE 0.
* Entry not found in archived data, then it's a new register
      ELSE.
* if there is a difference in the amount then its a change
* register = 'M' otherwise skip the entry.
        IF LS_BALANCE-WRBTR = IS_FIAPPTD_BALANCE-WRBTR AND
           LS_BALANCE-SCB_IND = IS_FIAPPTD_BALANCE-SCB_IND.
          EV_REG_EXISTS = 'X'.
        ELSE.
          EV_REG_EXISTS   = 'X'.
          EV_REG_CHANGED  = 'X'.
          EV_ORIGIN_RUNID = LS_BALANCE-ORIGIN_RUN_ID.
        ENDIF.
      ENDIF.
*** END OF ILM Enablement ***
*        Entry does not exist, then it's a new register
    ELSE.
*     if there is a difference in the amount then its a change
*     register = 'M' otherwise skip the entry.
      IF LS_BALANCE-WRBTR = IS_FIAPPTD_BALANCE-WRBTR AND
         LS_BALANCE-SCB_IND = IS_FIAPPTD_BALANCE-SCB_IND.
        EV_REG_EXISTS = 'X'.
      ELSE.
        EV_REG_EXISTS   = 'X'.
        EV_REG_CHANGED  = 'X'.
        EV_ORIGIN_RUNID = LS_BALANCE-ORIGIN_RUN_ID.
      ENDIF.
    ENDIF.

  ENDMETHOD.