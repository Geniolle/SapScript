*&---------------------------------------------------------------------*
*& Report ZBIM_MONITOR_V1
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zbim_monitor_v1.


INCLUDE zbim_monitor_v1_top.
INCLUDE zbim_monitor_v1_scr.
include zbim_monitor_v1_lcl.
INCLUDE zbim_monitor_v1_frms.
INCLUDE zbim_monitor_v1_pbo.
INCLUDE zbim_monitor_v1_pai.

START-OF-SELECTION.

  PERFORM f_check_adm_block.

  IF gv_adm_block IS INITIAL.
    PERFORM f_update_before_run.
    PERFORM f_select_data.
    PERFORM f_display_alv.
  ENDIF.