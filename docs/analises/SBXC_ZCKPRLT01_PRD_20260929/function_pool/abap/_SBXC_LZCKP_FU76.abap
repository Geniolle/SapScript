FUNCTION /sbxc/zckp_send_email_lifnr.
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
**-------------------------------------------------------------------*
*  *  MODIFICAÇÕES
*  ---------------------------------------------------------------------*
*  & Autor        : A. Garrido (SBX)
*  & Data         : 02.06.2023
*  & Referência   : 130047, Envio de email por cockpit
*  & Transporte   : S4DK940140 + S4DK942836 + S4DK943039
*  & Ch. pesquisa : AFG-ID-130047
*  & Objetivo     : Adicionar o anexo no envio de emails a partir do
*                   cockpit SBX para clientes internos ou para fornecedores
*
*  &---------------------------------------------------------------------*
* Estruturas de cabeçalho e linha
  DATA: header       TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        wa_header    TYPE /sbxc/zckp_invh,
        item         TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        item_f       TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        l_returncode,
        lt_sval      LIKE sval OCCURS 0 WITH HEADER LINE,
        lv_pro       TYPE  /sbxc/zckp_processo,
        lv_ano       TYPE  gjahr,
        lv_seq       TYPE  /sbxc/zckp_seqno,
        opcao        TYPE i,
        est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem,
        processo_pos TYPE /sbxc/zckp_processo,
        ano_pos      TYPE gjahr,
        seqno_pos    TYPE /sbxc/zckp_seqno,
        gt_outtab    TYPE TABLE OF /sbxc/zckp_tab07 WITH HEADER LINE,
        gs_private   TYPE slis_data_caller_exit,
        gs_selfield  TYPE slis_selfield,
        g_exit(1)    TYPE c.


  DATA: BEGIN OF mensagens OCCURS 0,
          cod_mes  TYPE /sbxc/zckp_tab06-cod_mes,
          split,
          mensagem TYPE /sbxc/zckp_tab07-mensagem,
        END OF mensagens.

  TYPES: BEGIN OF ty_rej_msg,
           processo TYPE /sbxc/zckp_tab09-processo,
           mensagem TYPE /sbxc/zckp_tab09-mensagem,
           cod_mes  TYPE /sbxc/zckp_tab09-cod_mes,
           resposta TYPE char1,
         END OF ty_rej_msg.

*  DATA: it_rej_msg TYPE STANDARD TABLE OF ty_rej_msg,
*        wa_rej_msg TYPE ty_rej_msg,
*        lv_msg_idx TYPE sy-tabix.

  DATA: l_email     TYPE string,
        error_table LIKE TABLE OF rpbenerr,
        texto(255) ,
        obs         LIKE /sbxc/zckp_invh-url,
        title(100).
  DATA: n_lines TYPE i.
  DATA: l_data_doc(10).
  DATA: lv_send(1) VALUE 'X'.
  DATA: e_mail   TYPE adr6-smtp_addr,
        lv_spras TYPE lfa1-spras.

  DATA: l_subject LIKE thead-tdname,
        l_body    LIKE thead-tdname.
  DATA: text_table    TYPE STANDARD TABLE OF tline,
        wa_text_table LIKE LINE OF text_table.
  DATA: text               TYPE bcsy_text. " Mail body
  DATA: subject            TYPE so_obj_des.

  DATA: send_request       TYPE REF TO cl_bcs.
  DATA: usersender         TYPE user_addr-bname .
  DATA: document           TYPE REF TO cl_document_bcs.
  DATA: sender             TYPE REF TO cl_cam_address_bcs.
  DATA: recipient          TYPE REF TO if_recipient_bcs.
  DATA: recipient_bcc      TYPE REF TO if_recipient_bcs.
  DATA: recipient_cc      TYPE REF TO if_recipient_bcs.
  DATA: internet_address LIKE adr6-smtp_addr, bcc_email1 LIKE adr6-smtp_addr, bcc_email2 LIKE adr6-smtp_addr, bcc_email3 LIKE adr6-smtp_addr.

  DATA: bcs_exception      TYPE REF TO cx_bcs.

  DATA:  userrecip          TYPE syst-uname.
  DATA: sent_to_all        TYPE os_boolean.
*AFG-ID-130047
  DATA: ls_xml    TYPE /sbxc/st_img,
        comp      TYPE i,
        i_attachx TYPE  solix_tab,
        nomefile  TYPE sood-objdes,
        l_size    TYPE sood-objlen, " Size of Attachment
        c_ext     TYPE soodk-objtp.
****************************************************


  REFRESH msg_cockpit.
* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
  APPEND header.

  IF header-em_tratamento = 'X' AND header-user_tratamento <> sy-uname.
    "Processo em tratamento pelo utilizador &
    MESSAGE s036(/sbxc/zckp_cockpit) WITH header-user_tratamento.
  ELSE.
    IF  ctrl-status1 = '6'.
      "Status do documento não permite esta operação
      MESSAGE s026(/sbxc/zckp_cockpit).
    ELSE.
*AFG-ID-130047
      PERFORM ler_imagem USING ls_xml comp header-uuid i_attachx header-barcode header-ref_doc_no_orig.
      l_size = comp.
**************************************
      LOOP AT linha.
        MOVE-CORRESPONDING linha TO item.
* Ler ref1 na ekko
        IF item-po_number NE space.
          SELECT SINGLE ihrez INTO item-ref_1
            FROM ekko
            WHERE ebeln EQ item-po_number.

          CLEAR:  item-ref_doc,
                  item-ref_doc_year,
                  item-ref_doc_item.
        ENDIF.
        APPEND item.
      ENDLOOP.
      READ TABLE header INDEX 1.

      REFRESH: lt_sval.

      lt_sval-tabname = '/SBXC/ZCKP_INVH'.
      lt_sval-fieldname = 'URL'.
      lt_sval-fieldtext = TEXT-052."'Observações'.
      APPEND lt_sval. CLEAR lt_sval.

      CONCATENATE TEXT-025 header-seqno INTO title SEPARATED BY space.
      CALL FUNCTION 'POPUP_GET_VALUES'
        EXPORTING
          popup_title  = title
          start_column = '10'
          start_row    = '5'
        IMPORTING
          returncode   = l_returncode
        TABLES
          fields       = lt_sval.
      IF sy-subrc <> 0.
      ENDIF.
      LOOP AT lt_sval.

        IF lt_sval-fieldname = 'URL'.
          obs = lt_sval-value.
        ENDIF.

      ENDLOOP.

      IF header-doc_date IS NOT INITIAL.
        CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
          EXPORTING
            date_internal            = header-doc_date
          IMPORTING
            date_external            = l_data_doc
          EXCEPTIONS
            date_internal_is_invalid = 1
            OTHERS                   = 2.
      ENDIF.

* Obter endereço email do fornecedor
      SELECT SINGLE smtp_addr spras INTO (e_mail, lv_spras) FROM lfa1 INNER JOIN adr6 ON lfa1~adrnr = adr6~addrnumber "#EC CI_NOORDER
      WHERE lifnr = header-vendor.

      IF e_mail IS INITIAL.
        MESSAGE i056(/sbxc/zckp_cockpit).
        CLEAR lv_send.
        RETURN.
      ENDIF.

      "Verificar idioma
      IF lv_spras IS INITIAL.
        lv_spras = 'P'.
      ELSEIF lv_spras NE 'P'.
        lv_spras = 'E'.
      ENDIF.

      l_body = 'Z_EMAIL_FORN'.
      IF l_body IS INITIAL.
        MESSAGE s066(/sbxc/zckp_cockpit).
        CLEAR lv_send.
        RETURN.
      ENDIF.

      MOVE header-ref_doc_no TO subject.

      REFRESH text_table.
*      Body
      CALL FUNCTION 'READ_TEXT'
        EXPORTING
          id        = 'ST'
          language  = lv_spras "'P'
          name      = l_body
          object    = 'TEXT'
        TABLES
          lines     = text_table
        EXCEPTIONS
          not_found = 1.

      IF sy-subrc EQ 0.
        CLEAR wa_text_table.
        LOOP AT text_table INTO wa_text_table.
          REPLACE FIRST OCCURRENCE OF '&REF_DOC_NO&' IN wa_text_table-tdline WITH header-ref_doc_no.
          REPLACE FIRST OCCURRENCE OF '&DOC_DATE&' IN wa_text_table-tdline WITH l_data_doc.
          REPLACE FIRST OCCURRENCE OF '&OBS&' IN wa_text_table-tdline WITH obs.
          APPEND wa_text_table-tdline TO text.
        ENDLOOP.
      ENDIF.

*AFG-ID-130047
      MOVE obs TO texto.
      APPEND texto TO text.
      CLEAR texto.
****************************************************
      "Determina o endereço do utilizador
      DATA: smtp_addr TYPE adr6-smtp_addr.

      usersender = sy-uname.
      SELECT SINGLE a~smtp_addr INTO smtp_addr          "#EC CI_NOORDER
                FROM usr21 AS u
                  INNER JOIN adr6 AS a
                     ON  u~addrnumber = a~addrnumber
                    AND u~persnumber = a~persnumber
              WHERE bname EQ usersender.
      IF smtp_addr IS INITIAL.
        MESSAGE s065(/sbxc/zckp_cockpit).
        CLEAR lv_send.
        RETURN.
      ENDIF.

      TRY.
*     -------- create persistent send request ------------------------
          send_request = cl_bcs=>create_persistent( ).

*     -------- create and set document -------------------------------
* Create document for mail body
          document = cl_document_bcs=>create_document(
                          i_type    = 'RAW'
*                      i_type    = 'OTF'
                          i_text    = text
*                      i_length  = '12'
                          i_subject = subject ).

*AFG-ID-130047
*Add attachment
          CONCATENATE header-vendor header-ref_doc_no INTO nomefile.
          CONDENSE nomefile.
          CONCATENATE nomefile TEXT-058 INTO nomefile.
          CALL METHOD document->add_attachment
            EXPORTING
              i_attachment_type    = c_ext
              i_attachment_size    = l_size
              i_attachment_subject = nomefile
              i_att_content_hex    = i_attachx[].


*     add document to send request
          CALL METHOD send_request->set_document( document ).

*     --------- set sender -------------------------------------------
*Sender deixa de ser do usuário para passar a ser fixo: fornecedores@sonaecapital.pt
*  sender = cl_sapuser_bcs=>create( usersender ).
*  sender = cl_cam_address_bcs=>create_internet_address( 'fornecedores@sonaecapital.pt' ).
*Sender deixa de ser do usuário para passar a ser fixo: fornecedores@sonaecapital.pt


          sender = cl_cam_address_bcs=>create_internet_address( smtp_addr ).

          CALL METHOD send_request->set_sender
            EXPORTING
              i_sender = sender.


          IF NOT e_mail IS INITIAL.
*       create recipient: e-mail address from Import
            internet_address = e_mail.
            recipient = cl_cam_address_bcs=>create_internet_address(
                                              internet_address ).

            CALL METHOD send_request->add_recipient
              EXPORTING
                i_recipient = recipient.
*        i_express   = 'X'.

            recipient_cc = cl_cam_address_bcs=>create_internet_address(
                                          smtp_addr ).

            CALL METHOD send_request->add_recipient
              EXPORTING
                i_recipient = recipient_cc
                i_copy      = 'X'.

*AFG-ID-130047
            CLEAR recipient_cc.
            CALL FUNCTION 'HR_FBN_GET_USER_EMAIL_ADDRESS'
              EXPORTING
                user_id       = sy-uname
                reaction      = ''
              IMPORTING
                email_address = l_email
              TABLES
                error_table   = error_table.

            internet_address = l_email.
            recipient_cc = cl_cam_address_bcs=>create_internet_address(
                                              internet_address ).

            CALL METHOD send_request->add_recipient(
                i_recipient = recipient_cc
                i_copy      = 'X' ).
****************************************************

          ENDIF.

* Exception handling
        CATCH cx_bcs INTO bcs_exception.
          RETURN.
      ENDTRY.

      SELECT * FROM  /sbxc/zckp_hmail                   "#EC CI_NOORDER
    WHERE processo = header-processo AND
          ano = header-ano AND
          seqno = header-seqno.

      ENDSELECT.
      IF sy-subrc NE 0.
        /sbxc/zckp_hmail-num_mess = 1.
      ELSE.
        /sbxc/zckp_hmail-num_mess = /sbxc/zckp_hmail-num_mess + 1.
      ENDIF.
      MOVE-CORRESPONDING header TO  /sbxc/zckp_hmail.
      /sbxc/zckp_hmail-mandt = sy-mandt.
      /sbxc/zckp_hmail-sender = usersender.
      /sbxc/zckp_hmail-recip =  userrecip.
      /sbxc/zckp_hmail-smtp_addr =  internet_address.
      /sbxc/zckp_hmail-data_envio = sy-datum.
      /sbxc/zckp_hmail-texto = subject.
      INSERT /sbxc/zckp_hmail.


      CHECK lv_send EQ 'X'.

      TRY.
*     ---------- send document ---------------------------------------
          CALL METHOD send_request->send(
            EXPORTING
              i_with_error_screen = 'X'
            RECEIVING
              result              = sent_to_all ).
          COMMIT WORK.

* -----------------------------------------------------------
* *                     exception handling
* -----------------------------------------------------------
        CATCH cx_bcs INTO bcs_exception.
*      RAISE send_error.
      ENDTRY.

*      refresh = 'X'.
*
*      CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
*        TABLES
*          i_message_tab = msg_cockpit.
**endif.
    ENDIF.
  ENDIF.

ENDFUNCTION.
