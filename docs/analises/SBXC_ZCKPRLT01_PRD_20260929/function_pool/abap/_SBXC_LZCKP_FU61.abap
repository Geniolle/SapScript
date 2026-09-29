FUNCTION /SBXC/ZCKP_REGISTA_LOG1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(PROCESSO) TYPE  /SBXC/ZWS_PROCESS
*"     VALUE(CHAVE) TYPE  KEYFIELD
*"     VALUE(UTILIZADOR) TYPE  UNAME
*"     VALUE(INFORM) TYPE  CHAR80
*"     REFERENCE(CODRETORNO) TYPE  CHAR1
*"     REFERENCE(UUID) TYPE  /SBXC/ZUUIDO OPTIONAL
*"     REFERENCE(ISAP_UUID) TYPE  /SBXC/ZUUID OPTIONAL
*"  EXPORTING
*"     VALUE(SAP_UUID) TYPE  /SBXC/ZUUID
*"----------------------------------------------------------------------
 tables: /SBXC/WS_LOG.
  DATA:
    l_tst       LIKE tzonref-tstampl,
    L_tstc(27)  type c,
    l_tsttb(21) type c,
    guid        like sysuuid-c.

  if ISAP_UUID is initial.
*    CALL FUNCTION 'SYSTEM_UUID_C_CREATE'
*      IMPORTING
*        uuid = SAP_UUID. "uuid SAP
    CALL METHOD cl_esh_co_unique_id=>get_guid_c32
      RECEIVING
        r_guid_c32 = sap_uuid.
  else.
    SAP_UUID = ISAP_UUID.
  endif.

  guid = uuid.   "uuid - origem

  GET TIME STAMP FIELD l_tst.
  write l_tst to L_tstc.
  replace all occurrences of '.' in L_tstc
                        with ' ' in character mode.
  replace all occurrences of ',' in L_tstc
                        with ' ' in character mode.
  condense L_tstc no-gaps.

  L_tstc(8)   = sy-datum.
  L_tstc+8(6) = sy-uzeit.

  l_tsttb     = L_tstc.

  clear /SBXC/WS_LOG.
  /SBXC/WS_LOG-PROCESSO   = processo.
  /SBXC/WS_LOG-TIMESTPL   = l_tsttb.
  /SBXC/WS_LOG-UUID_SAP   = SAP_UUID.
  /SBXC/WS_LOG-UUID_ORIG  = guid.
  /SBXC/WS_LOG-CHAVE      = chave.
  /SBXC/WS_LOG-UTILIZADOR = utilizador.
  /SBXC/WS_LOG-INFORM     = inform.
  /SBXC/WS_LOG-retorno    = codretorno.
  insert /SBXC/WS_LOG.

ENDFUNCTION.
