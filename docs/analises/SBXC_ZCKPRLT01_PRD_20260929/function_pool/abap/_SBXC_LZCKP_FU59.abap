FUNCTION /sbxc/zckp_mm_liga_ult_varios.
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
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"     REFERENCE(CAB)
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header    TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header  TYPE /sbxc/zckp_invh, wa_headerc TYPE /sbxc/zckp_invh.

  DATA: item      TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.
  DATA: item_f    TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.

  DATA: l_status TYPE /sbxc/zckp_ctrl-status1,
        l_doc    TYPE rbkp-belnr.

  REFRESH msg_cockpit.
  DATA: fieldcat_linha TYPE slis_fieldcat_alv,
        fieldcat_tab   TYPE slis_t_fieldcat_alv,
        grupos         TYPE slis_t_sp_group_alv,
        wa_grupos      TYPE slis_sp_group_alv,
        wa_eventos     TYPE slis_alv_event,
        eventos        TYPE slis_t_event,
        layout         TYPE slis_layout_alv,
        is_variant     TYPE disvariant,
        reprepid       TYPE slis_reprep_id,
        grid_set       TYPE lvc_s_glay.

  DATA:    wa_fieldcat       LIKE LINE OF fieldcat_tab.

  DATA: programa LIKE sy-repid.

  DATA: BEGIN OF itab_doc_alv OCCURS 0,
          bukrs    LIKE bkpf-bukrs,
          belnr    LIKE bkpf-belnr,
          gjahr    LIKE bkpf-gjahr,
          processo LIKE /sbxc/zckp_ctrl-processo,
          ano      LIKE /sbxc/zckp_ctrl-ano,
          seqno    LIKE /sbxc/zckp_ctrl-seqno,
          doc_lo   LIKE /sbxc/zckp_tab14-doc_lo,
          cpudt    LIKE /sbxc/zckp_tab14-cpudt,
        END OF itab_doc_alv.

  DATA: itab_doc LIKE itab_doc_alv OCCURS 0 WITH HEADER LINE,
        itab_doc_aux LIKE itab_doc_alv OCCURS 0 WITH HEADER LINE.

  REFRESH: itab_doc, msg_cockpit, itab_doc_aux.
  DATA object_id LIKE sapb-sapobjid.


  DATA: es_exit_caused_by_user  TYPE  slis_exit_by_user,
        e_exit_caused_by_caller.

DATA: lo_docinf   TYPE REF TO zcl_bim.
  DATA: ls_bim TYPE zsckp_to_bim,
        lv_errortext   TYPE string.
  DATA: lv_return TYPE sy-subrc.
  CONSTANTS: c_fi TYPE zsckp_to_bim-mm_fi VALUE 'FI',
             c_mm TYPE zsckp_to_bim-mm_fi VALUE 'MM'.


* Preencher estrutura cabecalho
*  MOVE-CORRESPONDING cab TO header.
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  MOVE-CORRESPONDING header TO wa_header.
  APPEND header.
*  ENDLOOP.
  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
    LOOP AT linha.
      MOVE-CORRESPONDING linha TO item.
      APPEND item.
    ENDLOOP.

    IF
       ctrl-status1 = '6' OR ctrl-status1 =  '9'
      OR  ctrl-status1 = '3' OR  ctrl-status1 = '4'.
      MESSAGE s027(/sbxc/zckp_cockpit)." WITH ctrl-status1 header-ref_doc_no.

    ELSE.
      IF  header-doc_fi IS NOT INITIAL AND header-doc_fi <> '@B1@' .
* atualiza tab11
        MOVE-CORRESPONDING header TO /sbxc/zckp_tab14.
        MODIFY /sbxc/zckp_tab14.
        COMMIT WORK.
      ENDIF.
      status_ok = 'X'.
      CLEAR layout.
*    layout-colwidth_optimize = 'X'.
      layout-zebra = ' '.
      layout-no_vline = ' '.
      layout-no_hline = ' '.
      layout-def_status = 'A'.
      layout-edit = 'X'.
      layout-edit_mode = 'X'.
      CONCATENATE text-050 header-processo header-seqno header-ano INTO
      layout-window_titlebar SEPARATED BY space .
      is_variant-report = sy-repid.

      programa = sy-repid.
      grid_set-edt_cll_cb = 'X'.

* Eventos a capturar
      REFRESH eventos.
      REFRESH fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'BUKRS'.
      wa_fieldcat-ref_tabname = 'BKPF'.
      wa_fieldcat-col_pos = 1.
      wa_fieldcat-edit = 'X'.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'BELNR'.
      wa_fieldcat-ref_tabname = 'BKPF'.
      wa_fieldcat-col_pos = 2.
      wa_fieldcat-edit = 'X'.
      APPEND wa_fieldcat TO fieldcat_tab.

      CLEAR wa_fieldcat.
      wa_fieldcat-fieldname = 'GJAHR'.
      wa_fieldcat-ref_tabname = 'BKPF'.
      wa_fieldcat-col_pos = 3.
      wa_fieldcat-edit = 'X'.
      APPEND wa_fieldcat TO fieldcat_tab.

*    wa_eventos-name = 'DATA_CHANGED'.
*    wa_eventos-form = 'F_DATA_CHANGED'.
*    APPEND wa_eventos TO eventos.

*    wa_eventos-name = 'DATA_CHANGED'.
*    wa_eventos-form = 'DATA_CHANGED'.
*    APPEND wa_eventos TO eventos.

      DO 30 TIMES.
        itab_doc-bukrs = header-comp_code.
        MOVE-CORRESPONDING header TO itab_doc.
        IF  itab_doc-belnr = '@B1@'.
          CLEAR: itab_doc-belnr, itab_doc-doc_lo.
        ENDIF.
        APPEND itab_doc.
      ENDDO.
*     data: ls_event type slis_alv_event.
*ls_event-name = 'CALLER_EXIT'.
*append ls_event to eventos.

      CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
        EXPORTING
          it_fieldcat              = fieldcat_tab
          i_callback_user_command  = 'USER_COMMAND'
*         i_callback_top_of_page   = 'TOP_OF_PAGE'
*         i_background_id          = 'ALV_BACKGROUND'
          it_events                = eventos
          is_layout                = layout
          i_grid_settings          = grid_set
          i_callback_pf_status_set = 'SET_STATUS2'
          is_variant               = is_variant
          i_callback_program       = programa
          i_save                   = 'X'
          it_special_groups        = grupos
          i_screen_start_column    = 10
          i_screen_start_line      = 5
          i_screen_end_column      = 70
          i_screen_end_line        = 30
        IMPORTING
          e_exit_caused_by_caller  = e_exit_caused_by_caller
          es_exit_caused_by_user   = es_exit_caused_by_user
        TABLES
          t_outtab                 = itab_doc
        EXCEPTIONS
          program_error            = 1
          OTHERS                   = 2.


*    check ES_EXIT_CAUSED_BY_USER ne 'X'. "RSM

* verifica dados
      CHECK sy-ucomm <> '&AC1' AND sy-ucomm <> 'CANCEL'.
      DELETE itab_doc WHERE belnr IS INITIAL.
      DELETE ADJACENT DUPLICATES FROM itab_doc COMPARING bukrs belnr gjahr.

      DESCRIBE TABLE itab_doc LINES indice.
      IF indice = 0.
        status_ok = ' '.
      ENDIF.
      LOOP AT itab_doc.
        indice = sy-tabix.
* Verificar se documento existe
        CLEAR bkpf.
        SELECT SINGLE gjahr belnr bukrs cpudt cputm usnam awtyp awkey bldat stblg xblnr waers
     INTO (bkpf-gjahr, bkpf-belnr, bkpf-bukrs, bkpf-cpudt, bkpf-cputm, bkpf-usnam, bkpf-awtyp, bkpf-awkey, bkpf-bldat, bkpf-stblg, bkpf-xblnr, bkpf-waers )
             FROM bkpf
               WHERE bukrs = itab_doc-bukrs     AND
                     belnr = itab_doc-belnr  AND
                     gjahr = itab_doc-gjahr.


        IF sy-subrc NE 0.
          status_ok = ' '.
* Erro empresa do documento diferente da empresa da linha seleccionada
          wa_msg-msgid = 'F5A'.
          wa_msg-msgno = '397'.
          wa_msg-msgty = 'E'.
          wa_msg-msgv1 = itab_doc-belnr.
          wa_msg-msgv2 = itab_doc-bukrs.
          wa_msg-msgv3 = itab_doc-gjahr.
*    wa_msg-msgv1 = invoicedocnumber  .
          APPEND wa_msg TO msg_cockpit .
          else.
            APPEND itab_doc_aux.
        ENDIF.

* valida campos fornecedor/empresa/atribuiçã/data documento/moeda/bstat/fornecedor
*        SELECT SINGLE lifnr
*        INTO bseg-lifnr FROM bseg
*        WHERE bukrs = bkpf-bukrs
*          AND belnr = bkpf-belnr
*          AND gjahr = bkpf-gjahr
*          AND koart = 'K'
*          AND lifnr = wa_header-vendor.

        DATA: lv_wrbtr TYPE bseg-wrbtr.
        CLEAR lv_wrbtr.
        SELECT lifnr wrbtr
       INTO (bseg-lifnr, bseg-wrbtr) FROM bseg
       WHERE bukrs = bkpf-bukrs
         AND belnr = bkpf-belnr
         AND gjahr = bkpf-gjahr
         AND koart = 'K'
         AND lifnr = wa_header-vendor
          ORDER BY buzei ASCENDING.
          lv_wrbtr = lv_wrbtr + bseg-wrbtr.
        ENDSELECT.
        bseg-wrbtr = lv_wrbtr.

        IF header-comp_code <> itab_doc-bukrs.
          status_ok = ' '.
* Erro empresa do documento diferente da empresa da linha seleccionada
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '002'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = header-comp_code.
          wa_msg-msgv2 = itab_doc-bukrs.
*    wa_msg-msgv1 = invoicedocnumber  .
          APPEND wa_msg TO msg_cockpit .
        ENDIF.
*        IF bkpf-xblnr <> header-ref_doc_no.
**          Associação Impossível: Referência Cockpit & não corresponde à do Doc. &
*          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
*          wa_msg-msgno = '053'.
*          wa_msg-msgty = 'W'.
*          wa_msg-msgv1 = bkpf-xblnr.
*          wa_msg-msgv2 = header-ref_doc_no.
*          APPEND wa_msg TO msg_cockpit .
*
*        ENDIF.

        IF   bkpf-bldat <> header-doc_date.
*          Associação Impossível: DataDoc Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '040'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-bldat.
          wa_msg-msgv2 = header-doc_date.
          APPEND wa_msg TO msg_cockpit .

        ENDIF.

        IF bkpf-waers <> header-currency.
*          Associação Impossível: Moeda Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '041'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-waers.
          wa_msg-msgv2 = header-currency.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF   bkpf-stblg <> ' '.
*          Associação Impossível: Documento & Estornado
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '042'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bkpf-belnr.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

        IF bseg-lifnr <> wa_header-vendor.
*          Associação Impossível: Fornecedor Cockpit & não corresponde à do Doc. &
          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
          wa_msg-msgno = '043'.
          wa_msg-msgty = 'W'.
          wa_msg-msgv1 = bseg-lifnr.
          wa_msg-msgv2 = wa_header-vendor.
          APPEND wa_msg TO msg_cockpit .
        ENDIF.

*        IF bseg-wrbtr <> wa_header-gross_amount.
*          wa_msg-msgid = '/SBXC/ZCKP_COCKPIT'.
*          wa_msg-msgno = '046'.
*          wa_msg-msgty = 'W'.
*          wa_msg-msgv1 = bseg-wrbtr.
*          wa_msg-msgv2 = wa_header-gross_amount.
*          APPEND wa_msg TO msg_cockpit .
*        ENDIF.

        IF bkpf-awtyp EQ 'RMRP'.

          itab_doc-doc_lo = bkpf-awkey(10).
          itab_doc-cpudt = bkpf-cpudt.
          MODIFY itab_doc INDEX indice TRANSPORTING doc_lo cpudt.
        ELSE.
          CLEAR: itab_doc-doc_lo, itab_doc-cpudt.
          MODIFY itab_doc INDEX indice TRANSPORTING doc_lo cpudt.
        ENDIF.
      ENDLOOP.

      IF  msg_cockpit[] IS NOT INITIAL .
        CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
          TABLES
            i_message_tab = msg_cockpit.
        return.
      ELSE.
        CHECK bkpf-awtyp IS NOT INITIAL.
* Status fica atualizado com 3 ou 4 mediante o ultimo doc lido
        IF bkpf-awtyp EQ 'RMRP'.
          l_status = '3'. "Integrado via logistica
        ELSE.
          l_status = '4'. "Integrado via financeira
        ENDIF.

        indice = sy-tabix.
        REFRESH: item_f, it_return.

*      move-corresponding header to wa_header.
        ctrl-status1 = l_status.
*      header-status1 =  l_status.
        CLEAR wa_header-doc_estorno.

        header-doc_fi = 'VARIOS'.
        header-doc_lo =  'VARIOS'.
        UPDATE /sbxc/zckp_invh SET
          doc_fi   = header-doc_fi
          doc_lo =    header-doc_lo
          doc_estorno = '          '
          ano_lanc = ' ' "wa_header-ano_lanc
          data_criacao = ' ' "tab_doc-cpudt'
         WHERE processo = header-processo AND
          ano = header-ano AND
          seqno = header-seqno.

        COMMIT WORK.

        UPDATE /sbxc/zckp_ctrl SET
*         status1 = '3'
           status1 = l_status
           status2 = 'M' "Actualizado manualmente
           data_chg_st1 = sy-datum "bkpf-cpudt
           hora_chg_st1 = sy-uzeit "bkpf-cputm
           user_chg_st1 = sy-uname "bkpf-usnam
          WHERE processo = header-processo AND
           ano = header-ano AND
           seqno = header-seqno.
* Actualizar estruturas
        COMMIT WORK.

        "Envia status para Saphety
        CALL FUNCTION '/SBXC/ZCKP_ENVIA_INF_SAPHETY'
          EXPORTING
            cab    = header
            status = 'ACCOUNTED'.

        IF header-url IS NOT INITIAL.

          LOOP AT itab_doc_aux.

            CALL FUNCTION '/SBXC/ZCKP_ANEXA_URL'
              EXPORTING
                bukrs = header-comp_code
                belnr = itab_doc_aux-belnr
                gjahr = header-pstng_date(4)
                url   = header-url.

          ENDLOOP.
        ENDIF.


*  Se se tratar de uma NC, enviar email para o fornecedor
        DATA: ls_nc.
        CLEAR: ls_nc, wa_headerc.
        LOOP AT itab_doc.

          IF ls_nc IS INITIAL.
            CALL FUNCTION '/SBXC/ZCKP_CHECK_NC'
              EXPORTING
                i_bukrs  = itab_doc-bukrs
                i_doc_fi = itab_doc-belnr
                i_gjahr  = itab_doc-gjahr
              IMPORTING
                e_nc     = ls_nc.
          ELSE.
            IF wa_headerc IS INITIAL.
              wa_headerc-comp_code = itab_doc-bukrs.
              wa_headerc-doc_fi    = itab_doc-belnr.
              wa_headerc-ano_lanc  = itab_doc-gjahr.
            ENDIF.
          ENDIF.
*  Se se tratar de uma NC, enviar email para o fornecedor

          MOVE-CORRESPONDING itab_doc TO /sbxc/zckp_tab14.
          /sbxc/zckp_tab14-doc_fi = itab_doc-belnr.
          /sbxc/zckp_tab14-ano_lanc = itab_doc-gjahr.
          /sbxc/zckp_tab14-comp_code = itab_doc-bukrs.
          /sbxc/zckp_tab14-doc_lo = itab_doc-doc_lo.
          MODIFY /sbxc/zckp_tab14.

        ENDLOOP.

*  Se se tratar de uma NC, enviar email para o fornecedor
        IF ls_nc EQ 'X'.
*         Enviar email para o fornecedor
          CALL FUNCTION '/SBXC/ZCKP_SENDMAIL_FORN'
            EXPORTING
              i_bukrs      = wa_headerc-comp_code
              i_ref_doc_no = wa_headerc-ref_doc_no
              i_gjahr      = wa_headerc-ano_lanc
              i_vendor     = wa_header-vendor
              i_doc_date   = wa_header-doc_date
              i_pstng_date = wa_header-pstng_date
              i_nc         = ls_nc.
        ENDIF.
*  Se se tratar de uma NC, enviar email para o fornecedor


        "Envia para workflow
            ls_bim-bukrs = wa_header-comp_code.
            ls_bim-belnr = wa_header-doc_fi.
            ls_bim-gjahr = wa_header-ano_lanc.
            CASE l_status.
              WHEN '3'.
                ls_bim-doc_lo = wa_header-doc_lo.
                ls_bim-mm_fi = c_mm.
              WHEN '4'.
                ls_bim-mm_fi = c_fi.
            ENDCASE.


            CREATE OBJECT lo_docinf.

            TRY.
                lo_docinf->send_to_bim_process( EXPORTING i_zsckp_to_bim = ls_bim
                                                IMPORTING e_wi_id = wa_header-wi_id
                                                          return_code = lv_return ).

*              CATCH cx_ai_system_fault INTO lo_systemfault.
*                lv_errortext = lo_systemfault->errortext.

*           wa_header-erro = lv_errortext.
            IF lv_return eq '4'.
              wa_header-erro = text-010.
              ELSEIF lv_return eq '1'.
                wa_header-erro = text-011.
            ENDIF.

           UPDATE /sbxc/zckp_invh SET erro = wa_header-erro
           WHERE processo EQ wa_header-processo
            AND ano EQ wa_header-ano
            AND seqno EQ wa_header-seqno.
           commit WORK AND WAIT.
            ENDTRY.

            IF wa_header-wi_id IS NOT INITIAL .
              UPDATE /sbxc/zckp_invh SET wi_id = wa_header-wi_id
                WHERE processo EQ wa_header-processo
                AND ano EQ wa_header-ano
                AND seqno EQ wa_header-seqno.
              COMMIT WORK AND WAIT.
            ENDIF.

            "Fim envio workflow

        IF wa_header-doc_fi <> ''.

          REFRESH: cor_tab[], cor[].
          cor[] = cor_tab[].
        ENDIF.
      ENDIF.
* Preencher estrutura cabecalho
*  move-corresponding cab to header.
      REFRESH:  linha.
      CLEAR cab.
*    LOOP AT header.
*      MOVE-CORRESPONDING header TO cab.
*      APPEND cab.
*    ENDLOOP.

      READ TABLE header INDEX 1.
      MOVE-CORRESPONDING header TO cab.

      LOOP AT item.
        MOVE-CORRESPONDING item TO linha.
        APPEND linha.
      ENDLOOP.

      CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
        TABLES
          i_message_tab = msg_cockpit.


    ENDIF.
  ENDIF.

ENDFUNCTION.
