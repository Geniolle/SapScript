*&---------------------------------------------------------------------*
*& Include          ZBIM_MONITOR_SCR
*&---------------------------------------------------------------------*

SELECTION-SCREEN: BEGIN OF BLOCK b01 WITH FRAME TITLE TEXT-001.
  SELECT-OPTIONS: s_bukrs  FOR t001-bukrs NO INTERVALS.
*  PARAMETERS: p_newalv AS CHECKBOX.
  PARAMETERS: p_log RADIOBUTTON GROUP g1 DEFAULT 'X',
              p_fi  RADIOBUTTON GROUP g1.
SELECTION-SCREEN: END OF BLOCK b01.
SELECTION-SCREEN: BEGIN OF BLOCK b02 WITH FRAME TITLE TEXT-002.
  SELECT-OPTIONS: s_belnr FOR rbkp-belnr,
                  s_gjahr FOR rbkp-gjahr,
                  s_ebeln FOR ekko-ebeln,
                  s_lifnr FOR zbim_blk_invoice-lifnr NO INTERVALS.
SELECTION-SCREEN: END OF BLOCK b02.
SELECTION-SCREEN: BEGIN OF BLOCK b03 WITH FRAME TITLE TEXT-003.
  PARAMETERS: p_spgrp AS CHECKBOX,
              p_spgrm AS CHECKBOX,
              p_razao AS CHECKBOX.
SELECTION-SCREEN: END OF BLOCK b03.

SELECTION-SCREEN: BEGIN OF BLOCK b04 WITH FRAME TITLE TEXT-004.
  SELECT-OPTIONS: s_wi_cd FOR zbim_blk_invoice-wi_cd,
                  s_stat FOR zbim_blk_invoice-wi_stat NO INTERVALS,
                  s_wiid FOR zbim_blk_invoice-wi_id NO INTERVALS,
                  s_wiuser FOR zbim_blk_invoice-wi_cruser NO INTERVALS.
SELECTION-SCREEN: END OF BLOCK b04.
SELECTION-SCREEN: BEGIN OF BLOCK b05 WITH FRAME TITLE TEXT-005.
  SELECT-OPTIONS: s_docinv FOR zbim_blk_invoice-zmblnr_inv  NO INTERVALS,
                  s_invdat FOR zbim_blk_invoice-zmjahr_inv  NO-EXTENSION NO INTERVALS,
                  s_docem FOR zbim_blk_invoice-zmblnr_em    NO INTERVALS,
                  s_emdat FOR zbim_blk_invoice-zmjahr_em    NO-EXTENSION NO INTERVALS,
                  s_doc_nc FOR zbim_blk_invoice-zbelnr_nc   NO INTERVALS,
                  s_ncdate FOR zbim_blk_invoice-zgjahr_nc   NO INTERVALS NO-EXTENSION,
                  s_status FOR  zbim_blk_invoice-zstatus    NO-EXTENSION NO INTERVALS,
                  s_zeop FOR zbim_blk_invoice-zeop_date     NO-EXTENSION,
                  s_docest FOR zbim_blk_invoice-zdoc_estorno NO-EXTENSION NO INTERVALS,
                  s_nc_nd  FOR zbim_blk_invoice-zdebit_note NO-EXTENSION NO INTERVALS.
SELECTION-SCREEN: END OF BLOCK b05.
SELECTION-SCREEN: BEGIN OF BLOCK b06 WITH FRAME TITLE TEXT-009.
  PARAMETERS: p_mywid  RADIOBUTTON GROUP g2,
              p_mywidt RADIOBUTTON GROUP g2,
              p_all    RADIOBUTTON GROUP g2,
              p_allwi  RADIOBUTTON GROUP g2.
SELECTION-SCREEN: END OF BLOCK b06.