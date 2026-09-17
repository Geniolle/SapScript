*----------------------------------------------------------------------*
***INCLUDE ZBIM_MONITOR_V1_PBO.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Module PBO OUTPUT
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
MODULE pbo OUTPUT.

  SET PF-STATUS 'MAIN100'.
  SET TITLEBAR 'MAIN100'.

  DATA: go_event TYPE REF TO lcl_alv.
  DATA: stable      TYPE lvc_s_stbl.

  IF g_custom_container IS INITIAL.
    CREATE OBJECT g_custom_container
      EXPORTING
        container_name = g_container.

    CREATE OBJECT g_grid
      EXPORTING
        i_parent = g_custom_container.

    gs_layout-stylefname = 'CELLTAB'.
    gs_layout-cwidth_opt = 'X'.
    gs_layout-zebra = 'X'.

    CREATE OBJECT go_event.
    SET HANDLER go_event->handle_hotspot_click FOR g_grid.

    PERFORM f_init_fieldcat.
    PERFORM f_init_style. "CHANGING lt_celltab.

    IF p_log IS NOT INITIAL.
      CALL METHOD g_grid->set_table_for_first_display
        EXPORTING
          i_structure_name = 'ZBIM_BLK_INVOICE'
          is_layout        = gs_layout
        CHANGING
          it_fieldcatalog  = gt_fieldcat
          it_outtab        = gt_log_data.
    ELSE.
      CALL METHOD g_grid->set_table_for_first_display
        EXPORTING
          i_structure_name = 'ZBIM_BLK_INV_FI'
          is_layout        = gs_layout
        CHANGING
          it_fieldcatalog  = gt_fieldcat
          it_outtab        = gt_fi_data.
    ENDIF.

    CALL METHOD g_grid->set_ready_for_input
      EXPORTING
        i_ready_for_input = 1.

    stable-col = 'X'.
    stable-row = 'X'.

*    CALL METHOD g_grid->refresh_table_display
*      EXPORTING
*        is_stable = stable.
  ELSE.

*>>>NBA
*    CALL METHOD g_grid->refresh_table_display
*      EXPORTING
*        is_stable = stable.
*<<<NBA

  ENDIF.
ENDMODULE.