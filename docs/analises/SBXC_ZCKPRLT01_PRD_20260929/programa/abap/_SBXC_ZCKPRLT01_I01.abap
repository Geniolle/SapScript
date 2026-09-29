*&---------------------------------------------------------------------*
*&  Include           ZCKPRLT01_I01                                    *
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0100  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE user_command_0100 INPUT.

  REFRESH: linha_cab_sel, linha_cab_sel_ind.
  CALL METHOD go_grid_cab->get_selected_rows
    IMPORTING
      et_index_rows = linha_cab_sel_ind
      et_row_no     = linha_cab_sel.

  READ TABLE linha_cab_sel_ind ASSIGNING FIELD-SYMBOL(<fs_linha>) INDEX 1.
  IF sy-subrc EQ 0.
    READ TABLE <t_cab_table> ASSIGNING <cab> INDEX <fs_linha>-index.
  ENDIF.

  go_grid_lin->check_changed_data( ).
  go_grid_cab->check_changed_data( ).

  data(refresh) = 'X'.

  CALL METHOD go_grid_cab->check_changed_data
    IMPORTING
      e_valid   = data(validar)
    CHANGING
      c_refresh = refresh.

  save_ok = ok_code.
  CLEAR ok_code.

  CASE save_ok.
    WHEN 'BACK'.
      PERFORM exit_program.

    WHEN 'GRAVA'.
      PERFORM grava_alteracoes.

    WHEN 'DOC_ON'.
      PERFORM exibe_documento.

    WHEN 'DOC_OFF'.
      PERFORM esconde_documento.

    WHEN 'TOOG_ON'.
      PERFORM exibe_filtros.

    WHEN 'TOOG_OFF'.
      PERFORM esconde_filtros.
    WHEN OTHERS.

  ENDCASE.

  CALL METHOD go_grid_lin->refresh_table_display( is_stable = gc_stable ).

  CALL METHOD go_grid_cab->set_selected_rows
    EXPORTING
      it_index_rows            = linha_cab_sel_ind
      it_row_no                = linha_cab_sel
      is_keep_other_selections = 'X'.

ENDMODULE.                 " USER_COMMAND_0100  INPUT
