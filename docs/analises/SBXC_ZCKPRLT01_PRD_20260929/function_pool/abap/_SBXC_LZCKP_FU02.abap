FUNCTION /sbxc/zckp_cab_fat.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(STATUS1_TMP) TYPE  /SBXC/ZCKP_CTRL-STATUS1 OPTIONAL
*"     REFERENCE(EST_CAB) TYPE  /SBXC/ZCKP_TAB00-EST_CAB OPTIONAL
*"     REFERENCE(T_TAB12) TYPE  /SBXC/ZCKP_T_TAB12 OPTIONAL
*"     REFERENCE(T_HMAIL) TYPE  /SBXC/ZCKP_T_HMAIL OPTIONAL
*"  TABLES
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header LIKE /sbxc/zckp_invh.
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.

* Preencher estrutura cabecalho
  MOVE-CORRESPONDING cab TO header.
*
  IF header-vendor CO '0123456789 '.
    UNPACK header-vendor TO header-vendor.
  ENDIF.

  DATA: cabstat(30).
  FIELD-SYMBOLS: <fstat> TYPE any.
  cabstat = 'CAB-STATUS1'.
  ASSIGN (cabstat) TO <fstat>.
  IF <fstat> IS ASSIGNED.
    /sbxc/zckp_ctrl-status1 = <fstat>.
  ENDIF.
  IF /sbxc/zckp_ctrl-status1 = '3' OR /sbxc/zckp_ctrl-status1 = '4'.
* valida motivo bloq. factura
    CLEAR: rbkp, bkpf.
    SELECT SINGLE zlspr gsber                           "#EC CI_NOORDER
           INTO (header-motivo_bloq, header-gsber)
           FROM bseg
          WHERE bukrs = header-comp_code
            AND belnr = header-doc_fi
            AND gjahr = header-ano_lanc
            AND lifnr <> ''.

* verifica se houve actualizações realizadas fora do cockpit:
    IF header-doc_lo IS NOT INITIAL.
      SELECT SINGLE  stblg rbstat  bukrs lifnr waers rmwwr FROM  rbkp
                      INTO
        (header-doc_estorno, rbkp-rbstat, header-comp_code, header-vendor, header-currency, header-gross_amount)
                      WHERE belnr = header-doc_lo AND
                            gjahr = header-ano_lanc.
    ELSE.
      SELECT SINGLE    bldat budat xblnr stblg  FROM  bkpf
                      INTO (header-doc_date, header-pstng_date, header-ref_doc_no, header-doc_estorno)

                      WHERE bukrs =  header-comp_code AND
                            belnr = header-doc_fi AND
                            gjahr = header-ano_lanc.
    ENDIF.
* se a factura estornada fora de cockpit actualiza tabela de controle
    IF header-doc_estorno IS NOT INITIAL.
      UPDATE /sbxc/zckp_ctrl SET
      status1 = '5'

      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

      UPDATE /sbxc/zckp_invh SET
      doc_estorno = header-doc_estorno
       WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.
      COMMIT WORK.

      ctrl-status1 = '5'.
    ENDIF.

  ENDIF.

*Actualiza pré-editados
  IF /sbxc/zckp_ctrl-status1 EQ '9'.
    CLEAR: rbkp, bkpf.
    SELECT SINGLE * FROM rbkp WHERE belnr = header-doc_lo AND
                          gjahr = header-ano_lanc.
    IF rbkp-rbstat EQ '5'.
      UPDATE /sbxc/zckp_ctrl SET
         status1 = '3'
         WHERE processo = header-processo AND
         ano = header-ano AND
         seqno = header-seqno.
      COMMIT WORK.

      ctrl-status1 = '3'.
    ENDIF.
  ENDIF.
*Validar indicador de receções.
  DATA: BEGIN OF t_pedidos OCCURS 0,
          ebeln LIKE ekpo-ebeln,
          ebelp LIKE ekpo-ebelp,
          wemng LIKE eket-wemng.
  DATA: END OF t_pedidos.
  CLEAR t_pedidos. REFRESH t_pedidos.

  READ TABLE t_tab12 ASSIGNING FIELD-SYMBOL(<fs12>) WITH KEY processo = header-processo
                                                             poref    = 'X'.
  IF sy-subrc EQ 0.
    DATA(f_poref) = <fs12>-poref.
  ENDIF.

  IF f_poref = 'X'.
    SELECT * FROM /sbxc/zckp_invi
             INTO CORRESPONDING FIELDS OF TABLE item
            WHERE processo = header-processo AND
                 ano       = header-ano      AND
                 seqno     = header-seqno.
    LOOP AT item WHERE NOT po_number IS INITIAL.
      IF item-po_item IS INITIAL.
        SELECT ekpo~ebeln ekpo~ebelp eket~wemng FROM ekpo
                                                JOIN eket
                                                  ON ( ekpo~ebeln = eket~ebeln AND
                                                       ekpo~ebelp = eket~ebelp )
                                                INTO CORRESPONDING FIELDS OF t_pedidos
                                               WHERE ekpo~webre = 'X'
                                                 AND ekpo~ebeln = item-po_number
          ORDER BY ekpo~ebelp.
          COLLECT t_pedidos.
          CLEAR t_pedidos.
        ENDSELECT.
      ELSE.
        SELECT ekpo~ebeln ekpo~ebelp eket~wemng FROM ekpo
                                                JOIN eket
                                                  ON ( ekpo~ebeln = eket~ebeln AND
                                                       ekpo~ebelp = eket~ebelp )
                                                INTO CORRESPONDING FIELDS OF t_pedidos
                                               WHERE ekpo~webre = 'X'
                                                 AND ekpo~ebeln = item-po_number
                                                 AND ekpo~ebelp = item-po_item
          ORDER BY ekpo~ebelp.
          COLLECT t_pedidos.
          CLEAR t_pedidos.
        ENDSELECT.
      ENDIF.
      READ TABLE t_pedidos WITH KEY wemng = 0.
      IF sy-subrc <> 0.
*Todas as linhas têm receção.
        header-recep_linha = '2'.
      ELSE.
*Assumir inicialmente que não ha receção.
        header-recep_linha = '0'.
        LOOP AT t_pedidos WHERE wemng <> 0.
          header-recep_linha = '1'.
          EXIT.                                         "#EC CI_NOORDER
        ENDLOOP.
      ENDIF.
    ENDLOOP.

  ENDIF.
** valida se forn tem IRF
  SELECT SINGLE * FROM lfbw WHERE                       "#EC CI_NOORDER
      lifnr = header-vendor AND
      bukrs = header-comp_code.

  IF sy-subrc = 0.
    header-irf = 'X'.
    IF header-cat_irf IS INITIAL.
      header-cat_irf =  lfbw-witht.
    ENDIF.
    IF  header-cod_irf  IS INITIAL.
      header-cod_irf = lfbw-wt_withcd.
    ENDIF.
  ELSE.
    CLEAR: header-irf, header-cat_irf, header-cod_irf.
  ENDIF.

  SELECT SINGLE * INTO @DATA(ls_hmail)
    FROM /sbxc/zckp_hmail
    WHERE processo = @header-processo
      AND ano = @header-ano
      AND seqno = @header-seqno.

** verifica se existe email para o processo
*  READ TABLE t_hmail ASSIGNING FIELD-SYMBOL(<fs_hm>) WITH KEY  processo = header-processo
*                                                               ano = header-ano
*                                                               seqno = header-seqno.
  IF sy-subrc = 0.
    header-icon_email = '@E0@'.
  ENDIF.

  IF header-doc_fi = 'VARIOS'.
    header-doc_fi = '@B1@'.
  ENDIF.

  IF header-doc_lo = 'VARIOS'.
    header-doc_lo = '@B1@'.
  ENDIF.

  IF header-ref_doc_no IS NOT INITIAL.
    SELECT SINGLE zterm FROM ekko  INTO header-zterm WHERE
      ebeln = header-ref_doc_no .
  ENDIF.
  MOVE-CORRESPONDING header TO cab.

  REFRESH cor_tab.

  PERFORM set_color TABLES cor_tab
    USING header-processo /sbxc/zckp_ctrl-status1 est_cab.

  DATA: lv_icon_status LIKE icon-id.

  IF status1_tmp NE ctrl-status1.

* Actualizar icon na tabela de controlo
    READ TABLE gt_tab10
    WITH KEY processo = header-processo
      status = /sbxc/zckp_ctrl-status1.
    IF sy-subrc EQ 0.
      lv_icon_status = gt_tab10-icon_status.
    ENDIF.

    UPDATE /sbxc/zckp_ctrl SET
    icon_status = lv_icon_status
    WHERE processo = header-processo AND
    ano = header-ano AND
    seqno = header-seqno.

  ENDIF.
  cor[] = cor_tab[].

ENDFUNCTION.
