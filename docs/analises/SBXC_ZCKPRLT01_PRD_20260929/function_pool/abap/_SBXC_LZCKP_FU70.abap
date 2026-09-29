FUNCTION /sbxc/zckp_alt_massa_c.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM)
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------
  FIELD-SYMBOLS: <status>   TYPE any, <processo> TYPE any.
  DATA edit TYPE /sbxc/zckp_edit.
  ASSIGN COMPONENT 'STATUS1'  OF STRUCTURE cab TO <status>.
  ASSIGN COMPONENT 'PROCESSO' OF STRUCTURE cab TO <processo>.
  SELECT SINGLE edit FROM /sbxc/zckp_tab10
                     INTO edit
                    WHERE processo = <processo>
                      AND status   = <status>.
  IF edit = 'X'.
    LOOP AT lt_sval.
      CHECK lt_sval-value IS NOT INITIAL.
      CONDENSE lt_sval-value.
      IF lt_sval-value(4) NE '0.00'.
        CONCATENATE 'CAB' lt_sval-fieldname INTO l_fieldname
        SEPARATED BY '-'.
        UNASSIGN <field>.
        ASSIGN (l_fieldname) TO <field>.
        IF <field> IS ASSIGNED.
          <field> = lt_sval-value.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDIF.
ENDFUNCTION.
