FUNCTION /SBXC/ZCKP_ANEXA_URL1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(BUKRS) TYPE  BUKRS
*"     VALUE(BELNR) TYPE  BELNR_D
*"     VALUE(GJAHR) TYPE  GJAHR
*"     VALUE(URL) TYPE  SO_TEXT255
*"  EXPORTING
*"     VALUE(SUBRC) LIKE  SY-SUBRC
*"     VALUE(SUBRC_DESC) TYPE  /SBXC/ZCKP_ZERRORMSG
*"----------------------------------------------------------------------

  data document_id type sofmk.
  data folder_id  type sofdk.
  data lt_objhead type standard table of soli.
  data lt_objcont type standard table of soli.
  data ls_objcont type soli.
  data lt_urltab  type standard table of sood-objdes.
  data l_tab_size type i.
  data l_url_id   type so_url.
  data l_obj_id   type soodk.
  data l_obj_data type sood1.
*  data url        type so_url.
  data rel_doc    type borident.
  data is_object  type borident.
  data object     type sibflporb.

  data: bukrs_belnr(14).

  call function 'SO_FOLDER_ROOT_ID_GET'
    exporting
      region    = 'B'
    importing
      folder_id = folder_id
    exceptions
      others    = 1.

  concatenate '&KEY&' url(250) into ls_objcont.
  append ls_objcont to lt_objcont.

  l_obj_data-objsns = 'O'.
  l_obj_data-objla  = sy-langu.
  l_obj_data-objdes = url.

  move bukrs to bukrs_belnr.
  move belnr to bukrs_belnr+4.

*  CONCATENATE bukrs belnr gjahr INTO is_object-objkey.
  concatenate bukrs_belnr gjahr into is_object-objkey.

  move 'BKPF' to is_object-objtype.

  call function 'SO_OBJECT_INSERT'
    exporting
      folder_id             = folder_id
      object_type           = 'URL'
      object_hd_change      = l_obj_data
    importing
      object_id             = l_obj_id
    tables
      objhead               = lt_objhead
      objcont               = lt_objcont
    exceptions
      active_user_not_exist = 35
      folder_not_exist      = 6
      object_type_not_exist = 17
      owner_not_exist       = 22
      parameter_error       = 23
      others                = 1000.

  if sy-subrc = 0.
    document_id-foltp = folder_id-foltp.
    document_id-folyr = folder_id-folyr.
    document_id-folno = folder_id-folno.
    document_id-doctp = l_obj_id-objtp.
    document_id-docyr = l_obj_id-objyr.
    document_id-docno = l_obj_id-objno.
    if not document_id is initial.
      clear rel_doc.
      rel_doc-objkey  = document_id.
      rel_doc-objtype = 'MESSAGE'.
      call function 'BINARY_RELATION_CREATE'
        exporting
          obj_rolea    = is_object
          obj_roleb    = rel_doc
          relationtype = 'URL'
        exceptions
          others       = 1.
      if sy-subrc = 0.
        subrc = 0.
        subrc_desc = text-102.
        commit work.
      else.
        subrc = 0.
        subrc_desc = text-102.
      endif.
    endif.
  else.
    subrc = 99.
    subrc_desc = text-101.
  endif.

  if bukrs is initial or gjahr is initial or
     belnr is initial or url is initial.
    subrc = 1.
    subrc_desc = text-200.
  endif.

  select single * from bkpf where bukrs = bukrs and
                                  belnr = belnr and
                                  gjahr = gjahr.
  if sy-subrc ne 0.
    subrc = 1.
    subrc_desc = text-201.
  endif.





ENDFUNCTION.
