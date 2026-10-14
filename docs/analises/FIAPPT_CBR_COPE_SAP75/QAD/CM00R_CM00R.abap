  METHOD INIT_FIAPPTD_TRANS_STR.

    DATA: LS_FIAPPTD_TRANS LIKE ES_FIAPPTD_TRANS.

    LS_FIAPPTD_TRANS-BUKRS = GV_BUKRS.
    LS_FIAPPTD_TRANS-REPORTING_YEAR = GV_YEAR.
    LS_FIAPPTD_TRANS-PERIOD = GV_PERIOD.
    LS_FIAPPTD_TRANS-RUN_ID = GV_RUNID.
*The Run ID will be overwritten for the cases of changed or
*reversed registers with origin_run_id of the old register.
*    ls_fiapptd_trans-origin_run_id = gv_runid.
    ES_FIAPPTD_TRANS = LS_FIAPPTD_TRANS.

  ENDMETHOD.