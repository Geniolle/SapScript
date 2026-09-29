FUNCTION /sbxc/zckp_sendmail_forn.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(I_BUKRS) TYPE  BSEG-BUKRS
*"     REFERENCE(I_REF_DOC_NO) TYPE  XBLNR
*"     REFERENCE(I_GJAHR) TYPE  BSEG-GJAHR
*"     REFERENCE(I_VENDOR) TYPE  ELIFN
*"     REFERENCE(I_DOC_DATE) TYPE  /SBXC/ZCKP_INVH-DOC_DATE
*"     REFERENCE(I_PSTNG_DATE) TYPE  /SBXC/ZCKP_INVH-PSTNG_DATE
*"       OPTIONAL
*"     REFERENCE(I_BCC) TYPE  CHAR1 OPTIONAL
*"     REFERENCE(I_REJECT) TYPE  CHAR1 OPTIONAL
*"     REFERENCE(I_EST_MENS) TYPE  /SBXC/ZCKP_TAB06-EST_MENSAGEM
*"       OPTIONAL
*"     REFERENCE(I_COD_MENS) TYPE  /SBXC/ZCKP_TAB06-COD_MES OPTIONAL
*"     REFERENCE(I_NC) TYPE  CHAR1 OPTIONAL
*"  EXCEPTIONS
*"      SEND_ERROR
*"----------------------------------------------------------------------

  DATA: l_email TYPE string.
  DATA: lv_send(1) VALUE 'X'.

  DATA: send_request       TYPE REF TO cl_bcs. "" Send request
  DATA: short_message      TYPE REF TO cl_bcs.
  DATA: document           TYPE REF TO cl_document_bcs.
*  DATA: sender             TYPE REF TO cl_sapuser_bcs.
  DATA: sender             TYPE REF TO cl_cam_address_bcs.
  DATA: recipient          TYPE REF TO if_recipient_bcs.
  DATA: recipient_bcc      TYPE REF TO if_recipient_bcs.
  DATA: recipient_cc      TYPE REF TO if_recipient_bcs.
  DATA: dist_list          TYPE REF TO cl_distributionlist_bcs.
  DATA: bcs_exception      TYPE REF TO cx_bcs.
  DATA: sent_to_all        TYPE os_boolean.
  DATA: text               TYPE bcsy_text. " Mail body
  DATA:  l_attach    TYPE bcsy_text, " Attachment
         l_extension TYPE soodk-objtp VALUE 'OTF',  " TXT format
         l_size      TYPE sood-objlen, " Size of Attachment
         l_document  TYPE REF TO cl_document_bcs,   " Mail body
         wa_text     TYPE soli. " Work area for attach
  DATA: internet_address LIKE adr6-smtp_addr, bcc_email1 LIKE adr6-smtp_addr, bcc_email2 LIKE adr6-smtp_addr, bcc_email3 LIKE adr6-smtp_addr.
  DATA: addr_bcc           LIKE adr6-smtp_addr.
  DATA: usersender         TYPE user_addr-bname .
  DATA: userrecip          TYPE syst-uname.
  DATA: e_mail             TYPE adr6-smtp_addr,
        lv_spras           type lfa1-spras.
  DATA: l_butxt   LIKE t001-butxt,
        l_stceg   LIKE t001-stceg,
        l_subject LIKE thead-tdname,
        l_body    LIKE thead-tdname.
  DATA: subject            TYPE so_obj_des.
  DATA: text_table     TYPE STANDARD TABLE OF tline,
        wa_text_table  LIKE LINE OF text_table,
        l_mensagem     LIKE /sbxc/zckp_tab07-mensagem,
        l_data_doc(10),
        l_pstng_date(10).

  DATA: l_lfa1_name1 TYPE lfa1-name1.

  IF i_doc_date IS NOT INITIAL.
    CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
      EXPORTING
        date_internal            = i_doc_date
      IMPORTING
        date_external            = l_data_doc
      EXCEPTIONS
        date_internal_is_invalid = 1
        OTHERS                   = 2.
  ENDIF.

  IF i_pstng_date IS NOT INITIAL.
    CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
      EXPORTING
        date_internal            = i_pstng_date
      IMPORTING
        date_external            = l_pstng_date
      EXCEPTIONS
        date_internal_is_invalid = 1
        OTHERS                   = 2.
  ENDIF.

* Obter endereço email do fornecedor
  SELECT SINGLE smtp_addr spras INTO (e_mail, lv_spras) FROM lfa1 INNER JOIN adr6 ON lfa1~adrnr = adr6~addrnumber  "#EC CI_NOORDER
      WHERE lifnr = i_vendor.

  IF e_mail IS INITIAL.
    MESSAGE i056(/sbxc/zckp_cockpit).
    CLEAR lv_send.
    return.
  ENDIF.

  "Verificar idioma
        IF lv_spras IS INITIAL.
          lv_spras = 'P'.
        elseif lv_spras ne 'P'.
          lv_spras = 'E'.
        ENDIF.


*      Obter dados para preenchimento de variáveis no body do email
*      Obter nome da empresa e NIF
  SELECT SINGLE butxt stceg FROM t001 INTO ( l_butxt, l_stceg ) WHERE bukrs = i_bukrs.

CLEAR subject.
"Verificar se é uma nota de crédito
IF i_nc eq 'X'.
    CLEAR l_lfa1_name1.
  SELECT SINGLE name1 FROM lfa1 INTO l_lfa1_name1 WHERE lifnr = i_vendor.
*    Determinar se se trata de uma rejeição de documento ou não
  IF i_reject EQ 'X'.
    l_subject = '/SBXC/ZCKP_RESUBJECT'.
    l_body    = '/SBXC/ZCKP_REBODY'.
*    Carregar mensagem de motivo de rejeição
    SELECT SINGLE mensagem INTO l_mensagem
      FROM /sbxc/zckp_tab07 WHERE est_mensagem = i_est_mens AND cod_mes = i_cod_mens AND spras = sy-langu.
  ELSE.
    l_subject = '/SBXC/ZCKP_ESUBJECT'.
    l_body    = '/SBXC/ZCKP_EBODY'.
  ENDIF.

  REFRESH text_table.
  CLEAR wa_text_table.
*     Subject
  CALL FUNCTION 'READ_TEXT'
    EXPORTING
      id        = 'ST'
      language  = 'P'
      name      = l_subject
      object    = 'TEXT'
    TABLES
      lines     = text_table
    EXCEPTIONS
      not_found = 1.
  IF sy-subrc EQ 0.
    READ TABLE text_table INDEX 1 INTO wa_text_table.
    subject = wa_text_table-tdline.
    REPLACE FIRST OCCURRENCE OF '&f_notacredito&' IN subject WITH i_ref_doc_no.
    REPLACE FIRST OCCURRENCE OF '&f_motivo&' IN subject WITH l_mensagem.
  ENDIF.

  REFRESH text_table.
*      Body
  CALL FUNCTION 'READ_TEXT'
    EXPORTING
      id        = 'ST'
      language  = 'P'
      name      = l_body
      object    = 'TEXT'
    TABLES
      lines     = text_table
    EXCEPTIONS
      not_found = 1.

  IF sy-subrc EQ 0.
    CLEAR wa_text_table.
    LOOP AT text_table INTO wa_text_table.
      REPLACE FIRST OCCURRENCE OF '&f_nomeforn&' IN wa_text_table-tdline WITH l_lfa1_name1.
      REPLACE FIRST OCCURRENCE OF '&f_notacredito&' IN wa_text_table-tdline WITH i_ref_doc_no.
      REPLACE FIRST OCCURRENCE OF '&f_datadocumento&' IN wa_text_table-tdline WITH l_data_doc.
      REPLACE FIRST OCCURRENCE OF '&f_motivo&' IN wa_text_table-tdline WITH l_mensagem.
      REPLACE FIRST OCCURRENCE OF '&f_nomeempresa&' IN wa_text_table-tdline WITH l_butxt.
      REPLACE FIRST OCCURRENCE OF '&f_nifempresa&' IN wa_text_table-tdline WITH l_stceg.
      APPEND wa_text_table-tdline TO text.
    ENDLOOP.
  ENDIF.
else.
**    Determinar se se trata de uma rejeição de documento ou não
*  IF i_reject EQ 'X'.
*    l_subject = '/SBXC/ZCKP_RESUBJECT'.
*    l_body    = '/SBXC/ZCKP_REBODY'.
*    Carregar mensagem de motivo de rejeição
    SELECT SINGLE mensagem INTO l_mensagem
      FROM /sbxc/zckp_tab07 WHERE est_mensagem = i_est_mens AND cod_mes = i_cod_mens AND spras = lv_spras. "sy-langu.
    select SINGLE tdname INTO l_body
      from /sbxc/zckp_tab06 WHERE est_mensagem = i_est_mens AND cod_mes = i_cod_mens.

      IF l_body IS INITIAL.
        MESSAGE s066(/sbxc/zckp_cockpit).
        CLEAR lv_send.
        return.
      ENDIF.

*  ELSE.
*    l_subject = '/SBXC/ZCKP_ESUBJECT'.
*    l_body    = '/SBXC/ZCKP_EBODY'.
*  ENDIF.

*  REFRESH text_table.
*  CLEAR wa_text_table.
**     Subject
*  CALL FUNCTION 'READ_TEXT'
*    EXPORTING
*      id        = 'ST'
*      language  = 'P'
*      name      = l_subject
*      object    = 'TEXT'
*    TABLES
*      lines     = text_table
*    EXCEPTIONS
*      not_found = 1.
*  IF sy-subrc EQ 0.
*    READ TABLE text_table INDEX 1 INTO wa_text_table.
*    subject = wa_text_table-tdline.
*    REPLACE FIRST OCCURRENCE OF '&f_notacredito&' IN subject WITH i_ref_doc_no.
*    REPLACE FIRST OCCURRENCE OF '&f_motivo&' IN subject WITH l_mensagem.
*  ENDIF.

IF subject IS INITIAL.
move i_ref_doc_no to subject.
ENDIF.

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
      REPLACE FIRST OCCURRENCE OF '&REF_DOC_NO_ORIG&' IN wa_text_table-tdline WITH i_ref_doc_no.
      REPLACE FIRST OCCURRENCE OF '&DOC_DATE&' IN wa_text_table-tdline WITH l_data_doc.
      REPLACE FIRST OCCURRENCE OF '&PSTNG_DATE&' IN wa_text_table-tdline WITH l_pstng_date.
      REPLACE FIRST OCCURRENCE OF '&MENSAGEM&' IN wa_text_table-tdline WITH l_mensagem.
*      REPLACE FIRST OCCURRENCE OF '&f_notacredito&' IN wa_text_table-tdline WITH i_ref_doc_no.
*      REPLACE FIRST OCCURRENCE OF '&f_datadocumento&' IN wa_text_table-tdline WITH l_data_doc.
*      REPLACE FIRST OCCURRENCE OF '&f_motivo&' IN wa_text_table-tdline WITH l_mensagem.
*      REPLACE FIRST OCCURRENCE OF '&f_nomeempresa&' IN wa_text_table-tdline WITH l_butxt.
*      REPLACE FIRST OCCURRENCE OF '&f_nifempresa&' IN wa_text_table-tdline WITH l_stceg.
      APPEND wa_text_table-tdline TO text.
    ENDLOOP.
  ENDIF.
        ENDIF.
   "Determina o endereço do utilizador
  data: smtp_addr TYPE adr6-smtp_addr.

     usersender = sy-uname.
  select single a~smtp_addr into smtp_addr  "#EC CI_NOORDER
            from usr21 as u
              inner join adr6 as a
                 on  u~addrnumber = a~addrnumber
                and u~persnumber = a~persnumber
          where bname eq usersender.
    IF smtp_addr IS INITIAL.
    MESSAGE s065(/sbxc/zckp_cockpit).
    CLEAR lv_send.
    return.
  ENDIF.

    TRY.
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

  ENDIF.

*  Verificar se deve enviar em bcc para responsável empresa
  IF i_bcc EQ 'X'.
*    Obter endereço(s) de email do responsável
    CLEAR: bcc_email1, bcc_email2, bcc_email3.
    SELECT SINGLE smtp_addr1 smtp_addr2 smtp_addr3 INTO (bcc_email1, bcc_email2, bcc_email3) FROM /sbxc/zckp_tab18
      WHERE bukrs = i_bukrs.
    IF sy-subrc EQ 0.

      IF bcc_email1 IS NOT INITIAL.
        recipient_bcc = cl_cam_address_bcs=>create_internet_address(
                                  bcc_email1 ).
        CALL METHOD send_request->add_recipient
          EXPORTING
            i_recipient  = recipient_bcc
            i_blind_copy = 'X'.
      ENDIF.

      IF bcc_email2 IS NOT INITIAL.
        recipient_bcc = cl_cam_address_bcs=>create_internet_address(
                                  bcc_email2 ).
        CALL METHOD send_request->add_recipient
          EXPORTING
            i_recipient  = recipient_bcc
            i_blind_copy = 'X'.
      ENDIF.

      IF bcc_email3 IS NOT INITIAL.
        recipient_bcc = cl_cam_address_bcs=>create_internet_address(
                                  bcc_email3 ).
        CALL METHOD send_request->add_recipient
          EXPORTING
            i_recipient  = recipient_bcc
            i_blind_copy = 'X'.
      ENDIF.

    ENDIF.
  ENDIF.

* Exception handling
      CATCH cx_bcs into bcs_exception.
        return.
        endtry.

  SELECT * FROM  /sbxc/zckp_hmail  "#EC CI_NOORDER
    WHERE processo = header-processo AND
          ano = header-ano AND
          seqno = header-seqno.

  ENDSELECT.
  IF sy-subrc NE 0.
    /sbxc/zckp_hmail-num_mess = 1.
  ELSE.
    /sbxc/zckp_hmail-num_mess = /sbxc/zckp_hmail-num_mess + 1.
  ENDIF.
  MOVE-CORRESPONDING /sbxc/zckp_tab09 TO  /sbxc/zckp_hmail.
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

ENDFUNCTION.
