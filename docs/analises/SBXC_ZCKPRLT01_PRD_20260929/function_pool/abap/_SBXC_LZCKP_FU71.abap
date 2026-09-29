FUNCTION /sbxc/zckp_alt_massa_c_pre.
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
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        gt_tab05 TYPE TABLE OF /sbxc/zckp_tab05 WITH HEADER LINE.
  MOVE-CORRESPONDING cab TO header.

  SELECT SINGLE * FROM /sbxc/zckp_tab00
         WHERE processo = header-processo.

  REFRESH gt_tab05.
  SELECT * FROM /sbxc/zckp_tab04
        WHERE processo  = header-processo
          AND estrutura = 'CAB'.
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
         WHERE tabname = /sbxc/zckp_tab00-est_cab AND
               fieldname = gt_tab05-campo.
    IF sy-subrc = 0.
      lt_sval-tabname = /sbxc/zckp_tab00-est_cab.
      lt_sval-fieldname =  gt_tab05-campo.
      APPEND lt_sval.
      CLEAR lt_sval.
    ENDIF.
  ENDLOOP.
  CALL FUNCTION 'POPUP_GET_VALUES'
    EXPORTING
      popup_title     = text-057 "'Alteração de Campos'
      start_column    = '10'
      start_row       = '5'
    IMPORTING
      returncode      = l_returncode
    TABLES
      fields          = lt_sval
    EXCEPTIONS
      error_in_fields = 1
      OTHERS          = 2.
ENDFUNCTION.
