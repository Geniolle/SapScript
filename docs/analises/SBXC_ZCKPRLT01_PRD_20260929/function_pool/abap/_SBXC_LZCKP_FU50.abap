FUNCTION /sbxc/zckp_cockpit_ver_imagem.
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

  TABLES: toa01.

  DATA: l_returncode,
          lt_sval LIKE sval OCCURS 0 WITH HEADER LINE.

  REFRESH msg_cockpit.

  DATA: BEGIN OF mensagens OCCURS 0,
          cod_mes TYPE /sbxc/zckp_tab06-cod_mes,
          split,
          mensagem TYPE /sbxc/zckp_tab07-mensagem,
        END OF mensagens.

  DATA: opcao TYPE i.


  DATA: est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem,
        lv_count     TYPE syindex,
        lt_connections TYPE TABLE OF toav0, ls_connection  TYPE toav0,
        lv_objkey      TYPE swo_typeid.

* Preencher estrutura cabecalho

  MOVE-CORRESPONDING cab TO header.
  APPEND header.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  LOOP AT header.


*ARCHIVOBJECT_DISPLAY
* doc_lo
*ano_lanc
    DATA: arc_doc_id TYPE toa01-arc_doc_id,
          object_id TYPE toa01-object_id.

    CONCATENATE header-comp_code header-DOC_FI header-ano_lanc INTO object_id.
*    select * from toa01 where sap_object = 'BUS2081' and
*                              object_id = object_id.


*
    CALL FUNCTION 'ARCHIV_GET_CONNECTIONS_INT'
      EXPORTING
        objecttype    = 'BKPF'
        object_id     = object_id
      IMPORTING
        count         = lv_count
      TABLES
        connections   = lt_connections
      EXCEPTIONS
        nothing_found = 1
        OTHERS        = 2.

*    IF sy-subrc = 0.
*      em_has_attach = abap_true.
*    ENDIF.

    READ TABLE lt_connections INTO ls_connection INDEX 1.
    if sy-subrc = 0.
    CALL FUNCTION 'ARCHIVOBJECT_DISPLAY'
      EXPORTING
        archiv_doc_id            = ls_connection-arc_doc_id "toa01-arc_doc_id
        archiv_id                = ls_connection-archiv_id "toa01-archiv_id
      EXCEPTIONS
        error_archiv             = 1
        error_communicationtable = 2
        error_kernel             = 3
        OTHERS                   = 4.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
    endif.

  IF sy-subrc = 0.

  ELSE.
    CHECK header-url NE space.
    CALL FUNCTION 'CALL_BROWSER'
     EXPORTING
       url                          = header-url
       window_name                  = text-048 "'Documento'
       new_window                   = 'X'
*       BROWSER_TYPE                 =
*       CONTEXTSTRING                =
     EXCEPTIONS
       frontend_not_supported       = 1
       frontend_error               = 2
       prog_not_found               = 3
       no_batch                     = 4
       unspecified_error            = 5
       OTHERS                       = 6
              .
    IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.


  ENDIF.

  MOVE-CORRESPONDING header TO cab.

ENDLOOP.

REFRESH linha.
LOOP AT item.
  MOVE-CORRESPONDING item TO linha.
  APPEND linha.
ENDLOOP.



ENDFUNCTION.
