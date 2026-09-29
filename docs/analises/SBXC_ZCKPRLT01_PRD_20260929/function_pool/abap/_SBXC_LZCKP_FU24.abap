FUNCTION /sbxc/zckp_mm_regista_adnt_ckp.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) LIKE  SY-UCOMM
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_adiantamento WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_adiantamento.

* Preencher estrutura cabecalho
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
*    APPEND header.
*  ENDLOOP.

*  LOOP AT header.

  CLEAR status_ok.
  PERFORM valida_status USING '0125'
                              header-processo
                              header-ano
                              header-seqno
                              ctrl-status1
                    CHANGING status_ok.

  CHECK status_ok = 'X'.

  MOVE-CORRESPONDING header TO wa_header.

  CALL FUNCTION '/SBXC/ZCKP_MM_REGISTA_ADIANT'
    EXPORTING
      headerdata             = header
* IMPORTING
*   INVOICEDOCNUMBER       =
*   FISCALYEAR             =
    TABLES
*        itemdata               = item_f
      return                 = it_return
            .

*  ENDLOOP.

ENDFUNCTION.
