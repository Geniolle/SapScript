  METHOD CONSTRUCTOR.

    DATA: LT_FIAPPTD_RUNID TYPE TABLE OF FIAPPTD_RUNID,
          LS_FIAPPTD_RUNID TYPE FIAPPTD_RUNID.

    GV_BUKRS = IV_BUKRS.
    GV_YEAR = IV_YEAR.
    GV_PERIOD = IV_PERIOD.
    GV_FSV = IV_FSV.
    GV_RUNID = IV_RUNID.

    SELECT * INTO TABLE LT_FIAPPTD_RUNID FROM FIAPPTD_RUNID
       WHERE BUKRS = GV_BUKRS
         AND REPORTING_YEAR = GV_YEAR
         AND PERIOD = GV_PERIOD.
    IF SY-SUBRC NE 0.
      GV_PERIOD_1ST_RUN = 'X'.
    ELSE.
      SORT LT_FIAPPTD_RUNID BY RUN_DATE RUN_TIME DESCENDING.
      READ TABLE LT_FIAPPTD_RUNID INTO LS_FIAPPTD_RUNID INDEX 1.
    ENDIF.
    GV_PREV_RUNID = LS_FIAPPTD_RUNID-RUN_ID.

*** BEGIN OF ILM Enablement ***
* Check if fetched documents have already been reported
* in another RunID and are already archived
    CLEAR: GT_AO_FIAPPTD_RUNID,
           GT_AO_FIAPPTD_TRANS,
           GT_AO_FIAPPTD_BALANCE.
* Get archived data for company code, year and period
    CALL METHOD CL_FI_CBR_PT_READ_AO=>READ_ARCHIVED_DATA
      EXPORTING
        IV_BUKRS                       = IV_BUKRS
        IV_YEAR                        = IV_YEAR
        IV_PERIOD                      = IV_PERIOD
*       iv_run_id                      =
*       iv_run_date                    =
      IMPORTING
        ET_FIAPPTD_RUNID               = GT_AO_FIAPPTD_RUNID
        ET_FIAPPTD_BALANCE             = GT_AO_FIAPPTD_BALANCE
        ET_FIAPPTD_TRANS               = GT_AO_FIAPPTD_TRANS
      EXCEPTIONS
        NO_ACTIVE_INFOSTRUCT_FOUND     = 1
        MISSING_ENTRIES_FOR_AS_FAST_AC = 2
        OTHERS                         = 3.
    IF SY-SUBRC <> 0.
* Couldn't read archive for some reason, do nothing
    ENDIF.
*** END OF ILM Enablement ***

  ENDMETHOD.