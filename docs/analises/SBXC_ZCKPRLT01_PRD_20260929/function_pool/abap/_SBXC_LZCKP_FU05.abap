FUNCTION /sbxc/zckp_cancel.
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


* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_hinvh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_hinvh.

  DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: lv_errortext   TYPE string.

DATA: ls_reverse TYPE ZSCKP_TO_BIM,
ld_aworg  type accdn-aworg,
       lt_accdn type table of accdn.

  REFRESH: msg_cockpit, msg_ckp.

  refresh = 'X'.

* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
    LOOP AT header.

      CLEAR status_ok.
*    PERFORM valida_status USING '34'
*                                header-processo
*                                header-ano
*                                header-seqno
*                                ctrl-status1
*                      CHANGING status_ok.
*
*    CHECK status_ok = 'X'.


      IF ctrl-status1 = '3'.


        CALL FUNCTION '/SBXC/ZCKP_CANCEL_DOC_LO'
          EXPORTING
            e_ucomm = e_ucomm
            idx_lin = idx_lin
          IMPORTING
            refresh = refresh
          TABLES
            linha   = linha
            cor     = cor
          CHANGING
            cab     = header "#EC CI_FLDEXT_OK[2610650]
            ctrl    = ctrl.

        ls_reverse-MM_FI = 'MM'.

      ELSEIF ctrl-status1 = '4'.
        CALL FUNCTION '/SBXC/ZCKP_CANCEL_DOC_FI'
          EXPORTING
            e_ucomm = e_ucomm
            idx_lin = idx_lin
          IMPORTING
            refresh = refresh
          TABLES
            linha   = linha
            cor     = cor
          CHANGING
            cab     = header "#EC CI_FLDEXT_OK[2610650]
            ctrl    = ctrl.
        ls_reverse-MM_FI = 'FI'.
      ELSEIF ctrl-status1 = '9'. "elimina documento pre-editado

        CALL FUNCTION '/SBXC/ZCKP_CANCEL_DOC_PRE_EDIT'
          EXPORTING
            e_ucomm = e_ucomm
            idx_lin = idx_lin
          IMPORTING
            refresh = refresh
          TABLES
            linha   = linha
            cor     = cor
          CHANGING
            cab     = header "#EC CI_FLDEXT_OK[2610650]
            ctrl    = ctrl.

      ls_reverse-MM_FI = 'MM'.
      ENDIF.


 "Envia reversão para workflow
      IF header-doc_estorno IS NOT INITIAL .
        ls_reverse-bukrs = header-comp_code.
        ls_reverse-belnr = header-doc_fi.
        ls_reverse-gjahr = header-ano_lanc.
        ls_reverse-doc_lo = header-doc_lo.

        IF header-doc_fi IS INITIAL. " Se o documento financeiro não estiver preecnhido então determina

          MOVE header-ano TO ld_aworg.
            CLEAR lt_accdn. REFRESH lt_accdn.
            CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
              EXPORTING
                i_awtyp      = 'RMRP'
                i_awref      = header-doc_lo
                i_aworg      = ld_aworg
              TABLES
                e_accdn      = lt_accdn
              EXCEPTIONS
                no_doc_found = 1
                OTHERS       = 2.


            IF sy-subrc = 0.
              LOOP AT lt_accdn ASSIGNING FIELD-SYMBOL(<fs1>) WHERE belnr IS NOT INITIAL
                and LDGRP IS INITIAL. "ODC - 04_11_2020.

                ls_reverse-belnr = ls_accdn-belnr.
                ls_reverse-gjahr = ls_accdn-gjahr.
                endloop.
                endif.

        ENDIF.
            CREATE OBJECT lo_docinf.

            TRY.
                lo_docinf->reverse_bim_process( exporting i_zsckp_to_bim = ls_reverse
                                                          doc_estorno = header-doc_estorno ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*           header-erro = lv_errortext.
*
*           UPDATE /sbxc/zckp_invh SET erro = header-erro
*           WHERE processo EQ header-processo
*            AND ano EQ header-ano
*            AND seqno EQ header-seqno.
*           commit WORK AND WAIT.
            ENDTRY.

      ENDIF.
  "Fim envio de reversão para workflow

      CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
        TABLES
          i_message_tab = msg_ckp.

      CLEAR cab.
      MOVE-CORRESPONDING header TO cab.

    ENDLOOP.
  ENDIF.
ENDFUNCTION.
