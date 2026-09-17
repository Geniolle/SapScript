*&---------------------------------------------------------------------*
*& Include          ZBIM_MONITOR_TOP
*&---------------------------------------------------------------------*
TABLES: t001,rbkp, ekko, zbim_blk_invoice.

TYPES: BEGIN OF ty_outtab_log.
         INCLUDE STRUCTURE zbim_blk_invoice.
TYPES:   celltab TYPE lvc_t_styl.
TYPES   END OF ty_outtab_log.

TYPES: BEGIN OF ty_outtab_fi.
         INCLUDE STRUCTURE zbim_blk_inv_fi.
TYPES:   celltab TYPE lvc_t_styl.
TYPES   END OF ty_outtab_fi.

DATA: t_alv        TYPE STANDARD TABLE OF zbim_blk_invoice,
      gt_log_data  TYPE STANDARD TABLE OF ty_outtab_log,
      t_alv_fi     TYPE STANDARD TABLE OF zbim_blk_inv_fi,
      gt_fi_data   TYPE STANDARD TABLE OF ty_outtab_fi,
      gv_adm_block TYPE abap_bool,
      save_ok      LIKE sy-ucomm,
      ok_code      LIKE sy-ucomm.

DATA:
  g_custom_container   TYPE REF TO cl_gui_custom_container,
  g_grid               TYPE REF TO cl_gui_alv_grid,
  gs_layout            TYPE lvc_s_layo,
  gt_toolbar_excluding TYPE TABLE OF ui_func,
  g_container          TYPE scrfname VALUE 'C_GRID',
  gs_f4                TYPE lvc_s_f4,
  gt_fieldcat          TYPE lvc_t_fcat,
  gr_column            TYPE REF TO cl_salv_column_table,
  gr_columns           TYPE REF TO cl_salv_columns_table,
  key                  TYPE salv_s_layout_key.

*data teste_sbx(1).