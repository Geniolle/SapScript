*----------------------------------------------------------------------*
***INCLUDE /SBXC/LZCKP_FI01.
*----------------------------------------------------------------------*

*{   INSERT         DEVK939019                                        1
*&---------------------------------------------------------------------*
*&      Module  SC_TC_MODIFY  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE sc_tc_modify INPUT.
DATA: ls_aux TYPE ty_ped.
    READ TABLE lt_ped INTO ls_aux INDEX sc_tc-current_line.
    IF sy-subrc EQ 0.
      MODIFY lt_ped
        FROM ls_ped
        INDEX sc_tc-current_line.
    ELSE.
      APPEND ls_ped TO lt_ped.
    ENDIF.
ENDMODULE.
*}   INSERT

*{   INSERT         DEVK939019                                        2
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0001  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE user_command_0001 INPUT.
CASE sy-ucomm.
  WHEN 'CANCEL'.
    refresh lt_ped.
    set SCREEN 0.
  WHEN 'OK'.
    set SCREEN 0.
  WHEN OTHERS.
ENDCASE.
ENDMODULE.
*}   INSERT
