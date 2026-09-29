FUNCTION /sbxc/zckp_associa_ped.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM) LIKE  SY-UCOMM
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

* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.

  DATA: item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: lt_sval LIKE sval OCCURS 0 WITH HEADER LINE.
  FIELD-SYMBOLS <status> TYPE any.

  REFRESH item_aux.
  DATA: status1(20),erro(10).

  status1 = 'CAB-STATUS1'.
  ASSIGN (status1) TO <status>.
* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
    LOOP AT linha.
      MOVE-CORRESPONDING linha TO item.
      APPEND item.
    ENDLOOP.

    IF  <status> = '9'.
      MESSAGE s011(/sbxc/zckp_cockpit)." WITH header-user_tratamento.
*Doc. já está pré-editado, continuar processamento com opção "Via MIRO"
    ELSEIF  <status> = '6'.
      MESSAGE s026(/sbxc/zckp_cockpit)." WITH header-user_tratamento.
      "Status do documento não permite esta operação
    ELSE.
* Insere novo pedido
      REFRESH: lt_sval.

      lt_sval-tabname = 'EKKO'.
      lt_sval-fieldname = 'EBELN'.
      APPEND lt_sval. CLEAR lt_sval.

      CALL FUNCTION 'POPUP_GET_VALUES'
        EXPORTING
*         NO_VALUE_CHECK        = ' '
          popup_title  = text-047 "'Nº Pedido'
          start_column = '10'
          start_row    = '5'
        IMPORTING
          returncode   = l_returncode
        TABLES
          fields       = lt_sval
 EXCEPTIONS
         ERROR_IN_FIELDS       = 1
         OTHERS       = 2
        .
      IF sy-subrc <> 0.
 MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.

      LOOP AT lt_sval.
        item-po_number = lt_sval-value.

        IF item-po_number  IS NOT INITIAL.
* Valida se pedido pertence à empresa
          SELECT SINGLE bukrs lifnr ekorg waers FROM ekko INTO (ekko-bukrs, ekko-lifnr, ekko-ekorg, ekko-waers) WHERE ebeln = item-po_number.
          ""ODC - 19_06_2020
          SELECT * INTO TABLE @DATA(lt_parvw)
            FROM ekpa
            WHERE ebeln EQ @ls_ped-ebeln
            AND ebelp EQ '00000'
            AND ekorg EQ @ekko-ekorg.
          ""Fim ODC - 19_06_2020
          IF ekko-bukrs NE header-comp_code.
            MESSAGE i020(/sbxc/zckp_cockpit) WITH  header-comp_code.
            erro = 'X'.
*   Pedido Selecionado não pertence à empresa &
          ELSEIF ekko-lifnr NE header-vendor.
            "ODC - 19_06_2020
          "Valida se o fornecedor existe em alguma função parceiro
          READ TABLE lt_parvw ASSIGNING FIELD-SYMBOL(<fs1>) WITH KEY lifn2 = ekko-lifnr.
          IF sy-subrc NE 0.
            MESSAGE i049(/sbxc/zckp_cockpit) WITH  header-vendor.
            erro = 'X'.
            ELSE.
              APPEND item.
              header-po_number = item-po_number. "ODC - 17_06_2020
          ENDIF.
*            MESSAGE i049(/sbxc/zckp_cockpit) WITH  header-vendor.
*            APPEND item.
*            header-po_number = item-po_number. "ODC - 17_06_2020
            ""Fim ODC - 19_06_2020
          ""ODC - 31_03_2021
        ELSEIF ekko-waers ne header-currency. "Validar se o pedido a associar está na mesma moeda do documento do cockpit
          "ERRO: Moeda Cockpit & não corresponde à do Doc. &
          MESSAGE i041(/sbxc/zckp_cockpit) WITH header-currency ekko-waers.
           erro = 'X'.
           ""Fim ODC - 31_03_2021
          ELSE.
            APPEND item.
            header-po_number = item-po_number. "ODC - 17_06_2020
          ENDIF.
        ENDIF.
      ENDLOOP.

      IF erro IS INITIAL.
        IF  <status> IS INITIAL OR <status> = '0' OR <status> = '5' OR <status> = '2'.
          PERFORM valida_qtd_fatura_sc TABLES item USING header .

          REFRESH linha.

* valida imputação contabilistica multipla
          PERFORM valida_imp_mult TABLES item_aux USING  header.

          SORT item_aux BY po_number po_item.
          DELETE ADJACENT DUPLICATES FROM item_aux COMPARING po_number po_item ref_doc ref_doc_year ref_doc_item invoice_doc_item.

          DATA: lv_ind TYPE /sbxc/zckp_invi-invoice_doc_item.
          lv_ind = 0.
          LOOP AT item_aux.
            MOVE-CORRESPONDING item_aux TO linha.
            item_aux-invoice_doc_item = lv_ind + 1.
            APPEND linha.
          ENDLOOP.
        ENDIF.
        refresh = 'X'.
      ENDIF.
    ENDIF.
  ENDIF.

  MOVE-CORRESPONDING header TO cab. "ODC - 17_06_2020
ENDFUNCTION.
