  METHOD INIT_FIAPPTD_BAL_STR.

    DATA: LS_FIAPPTD_BALANCE LIKE ES_FIAPPTD_BALANCE.
    DATA: LS_DATE_RANGE LIKE LINE OF IT_DATE_RANGE,
          LS_FIAPPTD_RUNID TYPE FIAPPTD_RUNID.

    LS_FIAPPTD_BALANCE-BUKRS = GV_BUKRS.
    LS_FIAPPTD_BALANCE-REPORTING_YEAR = GV_YEAR.
    LS_FIAPPTD_BALANCE-PERIOD = GV_PERIOD.

    LS_FIAPPTD_BALANCE-RUN_ID = GV_RUNID.
*   will be overwritten for changed registers.
*    ls_fiapptd_balance-origin_run_id = gv_runid.

    READ TABLE IT_DATE_RANGE INTO LS_DATE_RANGE INDEX 1.
    IF SY-DATUM IN IT_DATE_RANGE.
      LS_FIAPPTD_BALANCE-REFERENCE_DATE = SY-DATUM.
    ELSEIF SY-DATUM GT LS_DATE_RANGE-HIGH.
      LS_FIAPPTD_BALANCE-REFERENCE_DATE = LS_DATE_RANGE-HIGH.
    ENDIF.

    LS_FIAPPTD_BALANCE-VALUE_TYPE = 'P'.
    LS_FIAPPTD_BALANCE-BANK_ACC_TYPE = 'X'.

*    ls_fiapptd_balance-pais_ativo = ''. "blank for all balances.
*    ls_fiapptd_balance-due_date = ''. "blank for all balances.
*    ls_fiapptd_balance-observations = ''. "blank for all balances.

    SELECT SINGLE * FROM FIAPPTD_RUNID INTO LS_FIAPPTD_RUNID
      WHERE BUKRS = GV_BUKRS
        AND REPORTING_YEAR = GV_YEAR
        AND PERIOD = GV_PERIOD.
    IF SY-SUBRC EQ 0.  "Not the first run for this period.
      LS_FIAPPTD_BALANCE-REGISTER_ID = 'M'.
    ELSE.              "First run for this period.
      LS_FIAPPTD_BALANCE-REGISTER_ID = 'C'.
    ENDIF.

    ES_FIAPPTD_BALANCE = LS_FIAPPTD_BALANCE.

  ENDMETHOD.