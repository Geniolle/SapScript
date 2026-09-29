FUNCTION /sbxc/zckp_send_email_saphety.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(I_ATTACH) TYPE  /SBXC/ZCKP_MAIL_ATTACH_TAB OPTIONAL
*"     REFERENCE(E_UCOMM)
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"  EXCEPTIONS
*"      SEND_ERROR
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.
  refresh = 'X'.
  DATA: l_email     TYPE string,
        error_table LIKE TABLE OF rpbenerr,
        texto(255) ,
        obs         LIKE /sbxc/zckp_invh-url,
        title(100).
*  TABLES:  /sbxc/zckp_hmail.

  DATA: lv_send(1) VALUE 'X'.

  DATA: send_request       TYPE REF TO cl_bcs. "" Send request
  DATA: short_message      TYPE REF TO cl_bcs.
  DATA: document           TYPE REF TO cl_document_bcs.
  DATA: sender             TYPE REF TO cl_sapuser_bcs.
  DATA: recipient          TYPE REF TO if_recipient_bcs.
  DATA: dist_list          TYPE REF TO cl_distributionlist_bcs.
  DATA: bcs_exception      TYPE REF TO cx_bcs.
  DATA: sent_to_all        TYPE os_boolean.
  DATA: text               TYPE bcsy_text. " Mail body
  DATA: l_attach    TYPE bcsy_text, " Attachment
        l_extension TYPE soodk-objtp VALUE 'OTF',  " TXT format
        l_size      TYPE sood-objlen, " Size of Attachment
        l_document  TYPE REF TO cl_document_bcs,   " Mail body
        wa_text     TYPE soli. " Work area for attach
  DATA: internet_address   LIKE adr6-smtp_addr.
  DATA: addr_bcc           LIKE adr6-smtp_addr.
*  DATA: l_attach           TYPE /SBXC/CKP_MAIL_ATTACH.
  DATA: usersender         TYPE user_addr-bname .
  DATA: userrecip          TYPE syst-uname.
  DATA: e_mail             TYPE adr6-smtp_addr.
  DATA: subject            TYPE so_obj_des.
  DATA: lt_sval LIKE sval OCCURS 0 WITH HEADER LINE.
* Preencher estrutura cabecalho
*  move-corresponding cab to header.
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
*  APPEND header.
*  ENDLOOP.



  DATA: ls_xml    TYPE /sbxc/st_img.
  DATA: comp TYPE i.
  DATA: i_attachx    TYPE  solix_tab.
  DATA: nomefile     TYPE sood-objdes.
  DATA: c_ext        TYPE soodk-objtp.
  PERFORM ler_imagem USING ls_xml comp header-uuid i_attachx header-barcode header-ref_doc_no_orig.

  l_size = comp.

  usersender = sy-uname.

*  LOOP AT header.

*    IF sy-tabix = 1.
  REFRESH: lt_sval.

  lt_sval-tabname = 'USER_ADDR'.
  lt_sval-fieldname = 'BNAME'.
  APPEND lt_sval. CLEAR lt_sval.

  lt_sval-tabname = 'ADR6'.
  lt_sval-fieldname = 'SMTP_ADDR'.
  APPEND lt_sval. CLEAR lt_sval.

  lt_sval-tabname = '/SBXC/ZCKP_INVH'.
  lt_sval-fieldname = 'URL'.
  lt_sval-fieldtext = TEXT-052. "'Observações'.
  APPEND lt_sval. CLEAR lt_sval.

  CONCATENATE TEXT-025 header-seqno INTO title SEPARATED BY space.
  CALL FUNCTION 'POPUP_GET_VALUES'
    EXPORTING
*     NO_VALUE_CHECK  = ' '
      popup_title     = title
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
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

  LOOP AT lt_sval.

    IF  lt_sval-fieldname = 'BNAME'.
      userrecip = lt_sval-value.
    ELSEIF  lt_sval-fieldname = 'SMTP_ADDR'.
      e_mail = lt_sval-value.
    ELSEIF lt_sval-fieldname = 'URL'.
      obs = lt_sval-value.
    ENDIF.

  ENDLOOP.

  IF userrecip IS INITIAL AND e_mail IS INITIAL.
    MESSAGE s013(/sbxc/zckp_cockpit).
    CLEAR lv_send.
    RETURN.

*   Preencher Utilizador ou Endereço de email
  ENDIF.
  IF userrecip IS NOT INITIAL.
    CALL FUNCTION 'HR_FBN_GET_USER_EMAIL_ADDRESS'
      EXPORTING
        user_id       = userrecip
        reaction      = ''
      IMPORTING
        email_address = l_email
      TABLES
        error_table   = error_table.

    IF l_email IS INITIAL.
      MESSAGE s014(/sbxc/zckp_cockpit).
      CLEAR lv_send.
      RETURN.

**   Utilizador não tem endereço de email associado
    ELSE.
      e_mail = l_email.
    ENDIF.
  ENDIF.

  IF  e_mail NS '@'.
    MESSAGE s015(/sbxc/zckp_cockpit).

    CLEAR lv_send.
    RETURN.
*   Entrar endereço de email válido

  ENDIF.

*    ENDIF.

*    text[] = i_text[].
  CONCATENATE header-processo header-ano header-seqno INTO texto SEPARATED BY space.
  APPEND texto TO text.
  CLEAR texto.


  MOVE obs TO texto.
  APPEND texto TO text.
  CLEAR texto.

  subject = TEXT-027. "'Cockpit de Faturas: Validar os seguintes Processos'.


*    TRY.
*     -------- create persistent send request ------------------------
  send_request = cl_bcs=>create_persistent( ).

*     -------- create and set document -------------------------------
* Craete document for mail body
  document = cl_document_bcs=>create_document(
                  i_type    = 'RAW'
*                      i_type    = 'OTF'
                  i_text    = text
*                      i_length  = '12'
                  i_subject = subject ).
* Add attchment

  DATA: is_lporb       TYPE sibflporb,
        lt_roles       TYPE obl_t_role,
        ls_roles       TYPE LINE OF obl_t_role,
        gt_links       TYPE TABLE OF obl_s_link,
        gs_links       TYPE          obl_s_link,
        gs_folderid    TYPE          soodk,
        gs_objectid    TYPE          soodk,
        document_data  TYPE sofolenti1,
        object_content TYPE TABLE OF solisti1 WITH HEADER LINE.
  DATA: lv_filesize TYPE sood-objlen,
        document_id LIKE  sofolenti1-doc_id.

*    CLEAR is_lporb.
*    is_lporb-typeid = 'BKPF'.
*    is_lporb-catid = 'BO'.
*    CONCATENATE header-comp_code header-doc_fi header-ano_lanc
*    INTO is_lporb-instid.

*    REFRESH lt_roles.


*    TRY.
*        CALL METHOD cl_binary_relation=>read_links_of_binrel
*          EXPORTING
*            is_object   = is_lporb
*            ip_relation = 'ATTA'
*            ip_role     = 'GOSAPPLOBJ'
*          IMPORTING
*            et_links    = gt_links
*            et_roles    = lt_roles.
*      CATCH cx_obl_parameter_error .
*      CATCH cx_obl_internal_error .
*      CATCH cx_obl_model_error .
*    ENDTRY.

*    LOOP AT gt_links INTO gs_links.
*      CLEAR: document_data, object_content[].
*      document_id = gs_links-instid_b.
*      CALL FUNCTION 'SO_DOCUMENT_READ_API1'
*        EXPORTING
*          document_id                = document_id
*        IMPORTING
*          document_data              = document_data
*        TABLES
*          object_content             = object_content
*        EXCEPTIONS
*          document_id_not_exist      = 1
*          operation_no_authorization = 2
*          x_error                    = 3
*          OTHERS                     = 4.
*      IF sy-subrc <> 0.
** Implement suitable error handling here
*      ENDIF.
*
*      lv_filesize = document_data-doc_size.
*
*      CALL METHOD document->add_attachment
*        EXPORTING
*          i_attachment_type    = document_data-obj_type
*          i_attachment_subject = document_data-obj_descr
*          i_attachment_size    = lv_filesize
*          i_att_content_text   = object_content[].
*
*    ENDLOOP.

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
  sender = cl_sapuser_bcs=>create( usersender ).
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
*          i_express   = i_express.

* CCF Ini 26.01.2023 15:55:32
* inserir email de quem envia o email
    CALL FUNCTION 'HR_FBN_GET_USER_EMAIL_ADDRESS'
      EXPORTING
        user_id       = sy-uname
        reaction      = ''
      IMPORTING
        email_address = l_email
      TABLES
        error_table   = error_table.
    internet_address = e_mail.
    recipient = cl_cam_address_bcs=>create_internet_address(
                                      internet_address ).

    CALL METHOD send_request->add_recipient( i_recipient = recipient
       i_copy      = 'X' ).
* CCF Fim 26.01.2023 15:55:32

*  * Ir buscar o email BCC á tabela de Processos
*      clear addr_bcc.
*      select single smtp_addr_bcc from zckp_tab00
*        into addr_bcc
*        where processo = header-processo.
*
*      if addr_bcc is not initial.
*        internet_address = addr_bcc.
*        recipient = cl_cam_address_bcs=>create_internet_address(
*                                          internet_address ).
*
*        call method send_request->add_recipient
*          exporting
*            i_recipient  = recipient
*            i_blind_copy = 'X'.
*      endif.

  ENDIF.

  SELECT * FROM  /sbxc/zckp_hmail                       "#EC CI_NOORDER
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


  /sbxc/zckp_hmail-sender = usersender.
  /sbxc/zckp_hmail-recip =  userrecip.
  /sbxc/zckp_hmail-smtp_addr =  internet_address.
  /sbxc/zckp_hmail-data_envio = sy-datum.
  /sbxc/zckp_hmail-texto = obs.
  INSERT /sbxc/zckp_hmail.
  MOVE-CORRESPONDING  header TO cab.
*  ENDLOOP.

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

ENDFUNCTION.
