FUNCTION /sbxc/zckp_ver_so_img.
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
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE.
  DATA: wa_header TYPE /sbxc/zckp_invh.

  DATA: item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA: item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.

  DATA: is_lporb TYPE sibflporb,
        lt_roles TYPE obl_t_role,
        ls_roles TYPE LINE OF obl_t_role,
        gt_links    TYPE TABLE OF obl_s_link,
       gs_links    TYPE          obl_s_link,
       gs_folderid TYPE          soodk,
       gs_objectid TYPE          soodk.

  REFRESH msg_cockpit.

* Preencher estrutura cabecalho
*  LOOP AT cab.
  MOVE-CORRESPONDING cab TO header.
  APPEND header.
*  ENDLOOP.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  LOOP AT header.



    CLEAR is_lporb.
    is_lporb-typeid = 'BKPF'.
    is_lporb-catid = 'BO'.
    CONCATENATE header-comp_code header-doc_fi header-ano_lanc
    INTO is_lporb-instid.

    REFRESH lt_roles.

    TRY.
        CALL METHOD cl_binary_relation=>read_links_of_binrel
          EXPORTING
            is_object    = is_lporb
*           ip_logsys    =
            ip_relation  = 'ATTA'
            ip_role      = 'GOSAPPLOBJ'
*           ip_propnam   =
*           ip_no_buffer = SPACE
          IMPORTING
            et_links     = gt_links
            et_roles     = lt_roles.
      CATCH cx_obl_parameter_error .
      CATCH cx_obl_internal_error .
      CATCH cx_obl_model_error .
    ENDTRY.
*

  ENDLOOP.

*  READ TABLE lt_roles
*  INTO ls_roles WITH KEY roletype = 'ATTACHMENT'.

  DATA: lt_docs TYPE TABLE OF sood4 WITH HEADER LINE.

  LOOP AT gt_links INTO gs_links.

    lt_docs = gs_links-instid_b.
    APPEND lt_docs.

    AT LAST.
      CALL FUNCTION 'SO_DOCUMENTS_MANAGER'
        EXPORTING
          activity    = 'DISP'
*         OFFICE_USER =
        TABLES
          documents   = lt_docs.

    ENDAT.

  ENDLOOP.

ENDFUNCTION.
