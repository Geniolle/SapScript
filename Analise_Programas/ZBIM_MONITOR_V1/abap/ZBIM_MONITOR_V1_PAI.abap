*----------------------------------------------------------------------*
***INCLUDE ZBIM_MONITOR_V1_PAI.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  PAI  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE pai INPUT.

  save_ok = ok_code.
  CLEAR ok_code.

  CASE save_ok.
    WHEN 'SAIR'.
      LEAVE PROGRAM.
    WHEN 'VOLTAR'.
      PERFORM f_save_data USING 'V'.
*      CLEAR g_grid.
      SET SCREEN 0.LEAVE SCREEN.
    WHEN 'SAVE'.
      PERFORM f_save_data USING 'S'.
  ENDCASE.

*>>>NBA
  COMMIT WORK.
  REFRESH gt_log_data.
  PERFORM f_select_data.
  CALL METHOD g_grid->refresh_table_display
    EXPORTING
      is_stable = stable.
*<<<NBA


ENDMODULE.