FUNCTION /sbxc/zckp_altera_massa_pre.
*"--------------------------------------------------------------------
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
*"--------------------------------------------------------------------

  TABLES /sbxc/zckp_tab04.
* Estruturas de cabeçalho e linha
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: wa_item TYPE /sbxc/zckp_invi,
        p_idx_lin LIKE sy-index,
        gt_tab05 TYPE TABLE OF /sbxc/zckp_tab05 WITH HEADER LINE,
        l_fieldname TYPE c LENGTH 60.

  FIELD-SYMBOLS: <field> type any.

  MOVE idx_lin TO p_idx_lin.
*
** Preencher estrutura item
  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.



  REFRESH gt_tab05.

  READ TABLE item INTO wa_item INDEX p_idx_lin.
*  IF sy-subrc NE 0.
*    READ TABLE item INTO wa_item INDEX 1.
*  ENDIF.

  SELECT SINGLE est_lin FROM /sbxc/zckp_tab00 INTO /sbxc/zckp_tab00-est_lin
    WHERE processo   = wa_item-processo.

  SELECT * FROM /sbxc/zckp_tab04
              WHERE processo  = wa_item-processo
                AND estrutura = 'LIN'.
    SELECT SINGLE estrutura campo
             INTO (gt_tab05-estrutura, gt_tab05-campo)
             FROM  /sbxc/zckp_tab05
            WHERE  processo  = /sbxc/zckp_tab04-processo
              AND  estrutura = /sbxc/zckp_tab04-estrutura
              AND  campo     = /sbxc/zckp_tab04-campo
              AND  uname     = sy-uname.
    IF sy-subrc = 0. APPEND  gt_tab05. ENDIF.
  ENDSELECT.
  SORT  gt_tab05.
  DELETE ADJACENT DUPLICATES FROM gt_tab05.
  REFRESH: lt_sval.
  LOOP AT gt_tab05.

    SELECT SINGLE fieldname FROM dd03l INTO dd03l-fieldname
         WHERE tabname = /sbxc/zckp_tab00-est_lin AND
               fieldname = gt_tab05-campo.

    IF sy-subrc = 0.

      lt_sval-tabname = /sbxc/zckp_tab00-est_lin.
      lt_sval-fieldname =  gt_tab05-campo.
      APPEND lt_sval.
      IF lt_sval-fieldname = 'ITEM_AMOUNT'.
        lt_sval-tabname = '/SBXC/ZCKP_ST_INVH'.
        lt_sval-fieldname =  'CURRENCY'.
        lt_sval-field_attr = '04'.
        APPEND lt_sval. "CLEAR lt_sval.
      ENDIF.
      CLEAR lt_sval.
    ENDIF.

  ENDLOOP.


  CALL FUNCTION 'POPUP_GET_VALUES'
    EXPORTING
*     NO_VALUE_CHECK  = ' '
      popup_title     = text-049 "'Alteração de Campos'
      start_column    = '10'
      start_row       = '5'
    IMPORTING
      returncode      = l_returncode
    TABLES
      fields          = lt_sval
    EXCEPTIONS
      error_in_fields = 1
      OTHERS          = 2.
  IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.



ENDFUNCTION.
