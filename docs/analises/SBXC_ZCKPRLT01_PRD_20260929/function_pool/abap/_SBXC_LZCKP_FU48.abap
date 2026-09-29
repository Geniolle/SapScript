FUNCTION /sbxc/zckp_load_lfbw.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(LIFNR) TYPE  LIFNR
*"     REFERENCE(BUKRS) TYPE  BUKRS
*"  EXPORTING
*"     REFERENCE(IRF) TYPE  CHAR1
*"----------------------------------------------------------------------

  READ TABLE gt_irf WITH KEY lifnr = lifnr
  bukrs = bukrs.
  IF sy-subrc EQ 0.
    irf = gt_irf-irf.
  ELSE.
    SELECT SINGLE lifnr FROM lfbw
      INTO gt_irf-lifnr
       WHERE lifnr = lifnr AND
        bukrs = bukrs.
    IF sy-subrc EQ 0.
      gt_irf-lifnr = lifnr.
      gt_irf-bukrs = bukrs.
      gt_irf-irf = 'X'.
    ELSE.
      gt_irf-lifnr = lifnr.
      gt_irf-bukrs = bukrs.
      gt_irf-irf = ' '.
    ENDIF.
    APPEND gt_irf.
  ENDIF.



ENDFUNCTION.
