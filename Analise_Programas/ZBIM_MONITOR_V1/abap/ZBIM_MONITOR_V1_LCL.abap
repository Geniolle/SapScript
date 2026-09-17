*&---------------------------------------------------------------------*
*& Include          ZBIM_MONITOR_LCL
*&---------------------------------------------------------------------*
*----------------------------------------------------------------------*
*       CLASS lcl_alv DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_alv DEFINITION.
  PUBLIC SECTION.

    METHODS:
      handle_hotspot_click
        FOR EVENT hotspot_click OF cl_gui_alv_grid
        IMPORTING e_row_id
                  e_column_id.

ENDCLASS.                    "lcl_alv DEFINITION
*----------------------------------------------------------------------*
*       CLASS lcl_alv IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS lcl_alv IMPLEMENTATION.

  METHOD handle_hotspot_click.

    DATA: stable      TYPE lvc_s_stbl,
          lv_belnr    TYPE belnr_d,
          lv_gjahr    TYPE gjahr,
          lv_wi_id    TYPE sww_wiid,
          lv_bukrs    TYPE bukrs,
          lv_BUZEI    TYPE rblgp,
          ls_workitem TYPE swlwp1,
          lv_wi_stat  TYPE sww_wistat,
          lv_status   TYPE z_status,
          lv_agent    TYPE sww_aagent.

    IF p_log IS NOT INITIAL.
      "RSM@SBX-20220819-ini
      "Correção: tabela T_ALV não ordenada pelas ações da ALV. Utilizar a GT_LOG_DATA
*      READ TABLE t_alv INTO DATA(ls_alv) INDEX e_row_id.
      READ TABLE gt_log_data INTO DATA(ls_alv) INDEX e_row_id.
      "RSM@SBX-20220819-fim
      lv_belnr = ls_alv-belnr.
      lv_gjahr = ls_alv-gjahr.
      lv_bukrs =  ls_alv-bukrs.
      lv_buzei = ls_alv-buzei.
      lv_wi_id = ls_alv-wi_id.
      lv_wi_stat = ls_alv-wi_stat.
      lv_status = ls_alv-zstatus.
    ELSEIF p_fi IS NOT INITIAL.
      "RSM@SBX-20220819-ini
      "Correção: tabela T_ALV_FI não ordenada pelas ações da ALV. Utilizar a GT_FI_DATA
*      READ TABLE t_alv_fi INTO DATA(ls_alv1) INDEX e_row_id.
      READ TABLE gt_fi_data INTO DATA(ls_alv1) INDEX e_row_id.
      "RSM@SBX-20220819-fim
      lv_belnr = ls_alv1-belnr.
      lv_gjahr = ls_alv1-gjahr.
      lv_bukrs =  ls_alv1-bukrs.
      lv_wi_id = ls_alv1-wi_id.
      lv_bukrs = ls_alv1-bukrs.
      lv_buzei = ls_alv1-buzei.
      lv_wi_stat = ls_alv1-wi_stat.
      lv_status = ls_alv1-zstatus.
    ENDIF.

    CASE e_column_id.
      WHEN 'BELNR'.
        IF p_log IS NOT INITIAL.
          SET PARAMETER ID 'RBN' FIELD lv_belnr.
          SET PARAMETER ID 'GJR' FIELD lv_gjahr.
          CALL TRANSACTION 'MIR4' AND SKIP FIRST SCREEN.
        ELSEIF p_fi IS NOT INITIAL.
          SET PARAMETER ID 'BLN' FIELD lv_belnr.
          SET PARAMETER ID 'BUK' FIELD lv_bukrs.
          SET PARAMETER ID 'GJR' FIELD lv_gjahr.
          CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.
        ENDIF.
      WHEN 'EBELN'.
        SET PARAMETER ID 'BES' FIELD ls_alv-ebeln.
        CALL TRANSACTION 'ME23N' AND SKIP FIRST SCREEN.
      WHEN 'ZMBLNR_INV'.
        CALL FUNCTION 'MIGO_DIALOG'
          EXPORTING
            i_action            = 'A04'
            i_refdoc            = 'R02'
            i_notree            = 'X'
            i_deadend           = 'X'
            i_okcode            = 'OK_GO'
            i_mblnr             = ls_alv-zmblnr_inv
            i_mjahr             = ls_alv-zmjahr_inv
          EXCEPTIONS
            illegal_combination = 1
            OTHERS              = 2.
      WHEN 'ZMBLNR_EM'.
        CALL FUNCTION 'MIGO_DIALOG'
          EXPORTING
            i_action            = 'A04'
            i_refdoc            = 'R02'
            i_notree            = 'X'
            i_deadend           = 'X'
            i_okcode            = 'OK_GO'
            i_mblnr             = ls_alv-zmblnr_em
            i_mjahr             = ls_alv-zmjahr_em
          EXCEPTIONS
            illegal_combination = 1
            OTHERS              = 2.

      WHEN 'ZBELNR_NC'.
        SET PARAMETER ID 'RBN' FIELD ls_alv-zbelnr_nc.
        SET PARAMETER ID 'GJR' FIELD ls_alv-zgjahr_nc.
        CALL TRANSACTION 'MIR4' AND SKIP FIRST SCREEN.
      WHEN 'ZEXECUTE'.
        "CCF 22.08.2022 Ticket 130023
        "Guarda utilizador que esta a tratar o processo
        UPDATE zbim_blk_invoice SET user_trata = sy-uname
        WHERE bukrs = lv_bukrs
                AND belnr = lv_belnr
                AND gjahr = lv_gjahr
                AND buzei = lv_buzei.
        "Fim CCF 22.08.2022
        IF lv_status <> 'CLOSED' OR lv_wi_stat <> 'encerrado'.
          CALL FUNCTION 'SWL_WI_DISPATCH'
            EXPORTING
              wi_id                    = lv_wi_id
              wi_first_time            = 'X'
              wi_function              = 'APRO'
            EXCEPTIONS
              function_cancelled       = 1
              function_not_implemented = 2
              function_failed          = 3
              function_disabled        = 4
              OTHERS                   = 5.

          IF sy-subrc = 0.

            WAIT UP TO 4 SECONDS.

            SELECT SINGLE *
              FROM zbim_blk_invoice
              INTO @DATA(ls_blk_invoice)
              WHERE bukrs = @ls_alv-bukrs
                AND belnr = @ls_alv-belnr
                AND gjahr = @ls_alv-gjahr
                AND buzei = @ls_alv-buzei.

            ls_workitem-wi_id = lv_wi_id.

            IF ls_blk_invoice IS NOT INITIAL.
              READ TABLE t_alv ASSIGNING FIELD-SYMBOL(<fs>) WITH KEY bukrs = ls_alv-bukrs
                                                                     belnr = ls_alv-belnr
                                                                     gjahr = ls_alv-gjahr
                                                                     buzei = ls_alv-buzei.

              IF sy-subrc = 0.
                <fs>-zmblnr_em  = ls_blk_invoice-zmblnr_em.
                <fs>-zmjahr_em  = ls_blk_invoice-zmjahr_em.
                <fs>-zmblnr_inv = ls_blk_invoice-zmblnr_inv.
                <fs>-zmjahr_inv = ls_blk_invoice-zmjahr_inv.
                <fs>-zbelnr_nc  = ls_blk_invoice-zbelnr_nc.
                <fs>-zgjahr_nc  = ls_blk_invoice-zgjahr_nc.
                <fs>-zstatus    = ls_blk_invoice-zstatus.
                <fs>-zeop_date  = ls_blk_invoice-zeop_date.

                CALL FUNCTION 'ZBIM_GET_WIID'
                  EXPORTING
                    i_workitem = ls_workitem
                    i_status   = <fs>-zstatus
                  IMPORTING
                    o_wi_id    = <fs>-wi_id
                    o_status   = <fs>-wi_stat
                    o_agent    = lv_agent.

                <fs>-wi_cruser = lv_agent.
              ENDIF.
            ENDIF.

            CALL FUNCTION 'GET_GLOBALS_FROM_SLVC_FULLSCR'
              IMPORTING
                e_grid = g_grid.

            stable-col = 'X'.
            stable-row = 'X'.
          ELSE.
            MESSAGE i003(zbim).
          ENDIF.
        ELSE.
          MESSAGE i002(zbim).
        ENDIF.
*>>>NBA -INI  01.11.2021
      WHEN 'ICON_LOG_NC'.
        CHECK NOT ls_alv-balognr IS INITIAL. " Se existe algum log
        PERFORM exibe_log USING ls_alv-balognr.
      WHEN 'ICON_LOG_MM'.
        CHECK NOT ls_alv-BALOGNR_mm IS INITIAL. " Se existe algum log
        PERFORM exibe_log USING ls_alv-balognr_mm.
      WHEN 'ICON_LOG_DESB'.
        CHECK NOT ls_alv-balognr_desb IS INITIAL. " Se existe algum log
        PERFORM exibe_log USING ls_alv-balognr_desb.
*<<<NBA-FIM

    ENDCASE.
  ENDMETHOD.
ENDCLASS.