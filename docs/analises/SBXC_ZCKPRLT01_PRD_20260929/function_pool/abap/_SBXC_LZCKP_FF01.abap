*----------------------------------------------------------------------*
***INCLUDE /SBXC/LZCKP_FF01 .
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Form  MIR4_EL_PRE_EDITADA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_DOC_LO  text
*      -->P_HEADER_ANO_LANC  text
*----------------------------------------------------------------------*
FORM mir4_el_pre_editada  USING invoicedocnumber TYPE re_belnr
                                fiscalyear TYPE gjahr.

  CLEAR: bdcdata, messtab.
  REFRESH: bdcdata, messtab.

  CLEAR opt.
  opt-defsize  = 'X'.
  opt-dismode  = 'N'.
  opt-updmode  = 'E'.
  opt-racommit = 'X'.
  REFRESH messtab.
  PERFORM dynpro USING:  'X' 'SAPLMR1M'    '6150',
                           ' ' 'BDC_OKCODE'  '/00',
                           ' ' 'RBKP-BELNR'  invoicedocnumber,
                           ' ' 'RBKP-GJAHR' fiscalyear,
                           'X' 'SAPLMR1M'    '6000',
                           ' ' 'BDC_OKCODE'  '/EPPCH',
                           'X' 'SAPLMR1M'    '6000',
                           ' ' 'BDC_OKCODE'  '/EDELE'.


  CALL TRANSACTION 'MIR4' USING bdcdata
                          OPTIONS FROM opt
                          MESSAGES INTO messtab.

ENDFORM.                    " MIR4_EL_PRE_EDITADA
*&---------------------------------------------------------------------*
*&      Form  DYNPRO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_0049   text
*      -->P_0050   text
*      -->P_0051   text
*----------------------------------------------------------------------*
FORM dynpro  USING VALUE(dynbegin) VALUE(name) VALUE(value).

  save_sy_tabix = sy-tabix.
  CLEAR bdcdata.
  IF dynbegin = 'X'.
    bdcdata-program  = name.
    bdcdata-dynpro   = value.
    bdcdata-dynbegin = 'X'.
  ELSE .
    bdcdata-fnam = name.
    bdcdata-fval = value.
  ENDIF.
  APPEND bdcdata.
  sy-tabix = save_sy_tabix.

ENDFORM.                    " DYNPRO
*&---------------------------------------------------------------------*
*&      Form  VALIDA_STATUS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_0038   text
*      -->P_HEADER_PROCESSO  text
*      -->P_HEADER_ANO  text
*      -->P_HEADER_SEQNO  text
*      -->P_HEADER_STATUS1  text
*      <--P_STATUS_OK  text
*----------------------------------------------------------------------*
FORM valida_status USING status
                         processo
                         ano
                         seqno
                         status1
                CHANGING status_ok.

  status_ok = 'X'.

ENDFORM.                    " VALIDA_STATUS
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_BDCDATA
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_DOC_FI  text
*      -->P_HEADER_COMP_CODE  text
*      -->P_HEADER_ANO_LANC  text
*      -->P_REASON_REV  text
*----------------------------------------------------------------------*
FORM preenche_bdcdata  USING  p_doc_fi
                              p_comp_code
                              p_ano_lanc
                              reason_rev.
  PERFORM inserir_dados USING:
          'SAPMF05A' '0105' 'X',
          'RF05A-BELNS' p_doc_fi '',
          'BKPF-BUKRS'  p_comp_code '',
          'RF05A-GJAHS' p_ano_lanc '',
          'UF05A-STGRD' reason_rev '',
          'BDC_OKCODE' '=BU' ''.

  CALL TRANSACTION 'FB08' USING bdcdata
                          MODE  'E'
                          MESSAGES INTO t_messtab.
ENDFORM.                    " PREENCHE_BDCDATA
*&---------------------------------------------------------------------*
*&      Form  INSERIR_DADOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_1626   text
*      -->P_1627   text
*      -->P_1628   text
*----------------------------------------------------------------------*
FORM inserir_dados  USING nome numero ecran.

  CLEAR bdcdata.
  IF ecran = 'X'.
    MOVE: nome    TO bdcdata-program,
          numero  TO bdcdata-dynpro,
          ecran   TO bdcdata-dynbegin.
  ELSE.
    MOVE: nome    TO bdcdata-fnam,
          numero  TO bdcdata-fval.
  ENDIF.
  APPEND bdcdata.

ENDFORM.                    " INSERIR_DADOS
*&---------------------------------------------------------------------*
*&      Form  LOG_DELETE2
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_IT_FIMSG  text
*      -->P_HEADER_PROCESSO  text
*      -->P_HEADER_ANO  text
*      -->P_HEADER_SEQNO  text
*----------------------------------------------------------------------*
FORM log_delete2   TABLES  msg STRUCTURE fimsg
           USING   processo
                   ano
                   seqno.

  DELETE FROM /sbxc/zckp_tab08 WHERE processo = processo AND ano = ano AND
         seqno = seqno.

ENDFORM.                    " LOG_DELETE2
*&---------------------------------------------------------------------*
*&      Form  LOG2
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_IT_FIMSG  text
*      -->P_HEADER_PROCESSO  text
*      -->P_HEADER_ANO  text
*      -->P_HEADER_SEQNO  text
*----------------------------------------------------------------------*
FORM log2  TABLES  msg STRUCTURE  fimsg
            USING  processo
                   ano
                   seqno.

  CLEAR /sbxc/zckp_tab08-num.

  LOOP AT msg.
    ADD 1 TO /sbxc/zckp_tab08-num.
    /sbxc/zckp_tab08-processo = processo.
    /sbxc/zckp_tab08-ano = ano.
    /sbxc/zckp_tab08-seqno = seqno.

    /sbxc/zckp_tab08-msgid = msg-msgid.
    /sbxc/zckp_tab08-msgty = msg-msgty.
    /sbxc/zckp_tab08-msgno = msg-msgno.
    /sbxc/zckp_tab08-msgv1 = msg-msgv1.
    /sbxc/zckp_tab08-msgv2 = msg-msgv2.
    /sbxc/zckp_tab08-msgv3 = msg-msgv3.
    /sbxc/zckp_tab08-msgv4 = msg-msgv4.

    /sbxc/zckp_tab08-uname = sy-uname.
    /sbxc/zckp_tab08-data = sy-datum.
    /sbxc/zckp_tab08-hora = sy-uzeit.
    INSERT /sbxc/zckp_tab08.
  ENDLOOP.

ENDFORM.                                                    " LOG2
*&---------------------------------------------------------------------*
*&      Form  INICIALIZA_ESTRUTURAS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_C_NODATA  text
*----------------------------------------------------------------------*
FORM inicializa_estruturas  USING   nodata .

  CLEAR i_bgr00.
  i_bgr00-stype  = '0'.
  i_bgr00-nodata = nodata.


  PERFORM init_structures USING 'BBKPF' i_bbkpf nodata.
  i_bbkpf-stype = '1'.

  PERFORM init_structures USING 'BBSEG' i_bbseg nodata. "#EC CI_FLDEXT_OK[2610650]
  i_bbseg-stype = '2'.
  i_bbseg-tbnam = 'BBSEG'.

ENDFORM.                    " INICIALIZA_ESTRUTURAS
*&---------------------------------------------------------------------*
*&      Form  INIT_STRUCTURES
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_0352   text
*      -->P_I_BBKPF  text
*      -->P_NODATA  text
*----------------------------------------------------------------------*
FORM init_structures USING tabname tab i_nodata.


  DATA:lt_nametab TYPE TABLE OF dfies,
       ls_namet   TYPE dfies.

  FIELD-SYMBOLS: <f1> TYPE any.

  REFRESH nametab.

  CALL FUNCTION 'DDIF_NAMETAB_GET'
    EXPORTING
      tabname   = tabname
    TABLES
      dfies_tab = lt_nametab
    EXCEPTIONS
      not_found = 1
      OTHERS    = 2.

  LOOP AT lt_nametab INTO ls_namet.
    CLEAR char.
    CONCATENATE 'I_' ls_namet-tabname '-' ls_namet-fieldname INTO char.
    ASSIGN (char) TO <f1>.
    <f1> = i_nodata.
  ENDLOOP.

ENDFORM.                    " INIT_STRUCTURES
*&---------------------------------------------------------------------*
*&      Form  INIT_BBKPF
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_BBKPF  text
*----------------------------------------------------------------------*
FORM init_bbkpf  USING  bbkpf.
  bbkpf = i_bbkpf.
ENDFORM.                    " INIT_BBKPF
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_CAB
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_E_UCOMM  text
*      -->P_IT_CAB  text
*----------------------------------------------------------------------*
FORM preenche_cab  USING e_ucomm
                        it_cab STRUCTURE /sbxc/zckp_invh.

  DATA: dia(10).
  DATA: fat_estornada_mm(1).
  CASE e_ucomm.
    WHEN 'F-41' OR 'F-41B'.
      bbkpf-tcode     = 'FB01'.
    WHEN 'F-43' OR 'F-43B'.
      bbkpf-tcode     = 'FB01'.
    WHEN 'FB01B'.
      bbkpf-tcode     = 'FB01'.
    WHEN OTHERS.
      bbkpf-tcode     = e_ucomm.
  ENDCASE.

  WRITE it_cab-doc_date TO dia .
  TRANSLATE dia USING '. / - ' .
  CONDENSE dia NO-GAPS.
  bbkpf-bldat = dia.

  CLEAR dia.
  IF it_cab-pstng_date NE '00000000'.
    WRITE it_cab-pstng_date TO dia.
  ELSE.
    WRITE sy-datum TO dia .
  ENDIF.
  TRANSLATE dia USING '. / - ' .
  CONDENSE dia NO-GAPS.
  bbkpf-budat = dia.

  IF  it_cab-doc_type NE space.
    bbkpf-blart     = it_cab-doc_type.
  ENDIF.

  IF it_cab-comp_code NE space.
    bbkpf-bukrs     = it_cab-comp_code.
  ENDIF.

  IF it_cab-ref_doc_no NE space.
    bbkpf-xblnr(16) = it_cab-ref_doc_no.
  ENDIF.

  IF it_cab-header_txt NE space.
    bbkpf-bktxt = it_cab-header_txt.
  ENDIF.

  bbkpf-waers     = it_cab-currency.
  bbkpf-xmwst     = 'X'.

  TRANSFER bbkpf TO ds_name.

  "Valida duplicados
  DATA: lv_return TYPE sy-subrc.
  CALL FUNCTION '/SBXC/ZCKP_VALIDA_DUPLICADO'
    EXPORTING
      doctypesaphety = it_cab-doctypesaphety
    IMPORTING
      return         = lv_return
    CHANGING
      cab            = it_cab.
*         ctrl           = ctrl.
  IF lv_return EQ 4.
    RETURN.
  ENDIF.

*  IF it_cab-doctypesaphety EQ 'NC' OR it_cab-doctypesaphety EQ 'CREDITNOTE'
*  OR it_cab-doctypesaphety EQ 'CN' OR it_cab-doctypesaphety EQ '4'.
*    "Verificar Nota de crédito dupla
*    DATA flag_bkpf_n(1).
*    DATA flag_aux_n(1).
*    CLEAR flag_aux_n.
*    TRANSLATE wa_header-ref_doc_no  TO UPPER CASE.
*    SELECT * FROM bsip WHERE
*                  bukrs = it_cab-comp_code AND
*                  lifnr = it_cab-vendor AND
*                  waers = it_cab-currency AND
*                  bldat = it_cab-doc_date AND
*                  xblnr = it_cab-ref_doc_no AND
*                  wrbtr = it_cab-gross_amount AND
*                  gjahr =  it_cab-pstng_date(4) AND
*               shkzg = 'S'.
*      CLEAR flag_bkpf_n.
*      CLEAR bkpf.
*      flag_aux_n =  abap_true.
*      SELECT SINGLE * FROM bkpf WHERE belnr = bsip-belnr AND
*                                      bukrs = bsip-bukrs AND
*                                      gjahr = bsip-gjahr AND
*                                      xreversal EQ ' '.
*      IF sy-subrc = 0.
*        flag_bkpf_n = 'X'.
*      ELSE.
*        SELECT SINGLE  * FROM rbkp WHERE
*              bukrs = bkpf-bukrs AND
*              belnr = bkpf-awkey(10) AND
*              gjahr = bkpf-awkey+10(4) AND
*              stblg = ''.
*        IF sy-subrc NE 0.
*          fat_estornada_mm = 'X'.
*        ENDIF.
*      ENDIF.
*
*    ENDSELECT.
*    IF  ( fat_estornada_mm = ' ' AND flag_bkpf_n = 'X' ) OR ( fat_estornada_mm = ' ' AND flag_bkpf_n = ' '  AND flag_aux_n EQ abap_true )
*    OR ( fat_estornada_mm = 'X' AND flag_bkpf_n = 'X'  AND flag_aux_n EQ abap_true ).  "nao foi estornada nem em MM nem FI
*      MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr DISPLAY LIKE 'E'.
*      RETURN.
*    ENDIF.
*    "Fim Verificar Nota de crédito dupla
*  ELSE.
** Valida Duplicados
*    DATA flag_bkpf(1).
*    DATA flag_aux(1).
*    CLEAR flag_aux.
*    SELECT * FROM bsip WHERE
*          bukrs = it_cab-comp_code AND
*          lifnr = it_cab-vendor AND
*          waers = it_cab-currency AND
*          bldat = it_cab-doc_date AND
*          xblnr = it_cab-ref_doc_no AND
*          wrbtr = it_cab-gross_amount AND
*          gjahr =  it_cab-pstng_date(4) AND
*          shkzg = 'H'.
*      CLEAR flag_bkpf.
*      CLEAR bkpf.
*      flag_aux =  abap_true.
*      SELECT SINGLE * FROM bkpf WHERE belnr = bsip-belnr AND
*                                      bukrs = bsip-bukrs AND
*                                      gjahr = bsip-gjahr AND
*                                      xreversal EQ ' '.
*      IF sy-subrc = 0.
*        flag_bkpf = 'X'.
*      ELSE.
*        SELECT SINGLE  * FROM rbkp WHERE
*                  bukrs = bkpf-bukrs AND
*                  belnr = bkpf-awkey(10) AND
*                  gjahr = bkpf-awkey+10(4) AND
*                  stblg = ''.
*        IF sy-subrc NE 0.
*          fat_estornada_mm = 'X'.
*        ENDIF.
*      ENDIF.
*    ENDSELECT.
*
*    IF  ( fat_estornada_mm = ' ' AND flag_bkpf = 'X' ) OR ( fat_estornada_mm = ' ' AND flag_bkpf = ' '  AND flag_aux EQ abap_true )
*    OR ( fat_estornada_mm = 'X' AND flag_bkpf = 'X'  AND flag_aux EQ abap_true ).  "nao foi estornada nem em MM nem FI
*      MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr DISPLAY LIKE 'E'.
*      RETURN.
*    ENDIF.
*  ENDIF.
ENDFORM.                    " PREENCHE_CAB
*&---------------------------------------------------------------------*
*&      Form  BBSEG_FORN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_IT_CAB  text
*----------------------------------------------------------------------*
FORM bbseg_forn  USING e_ucomm it_cab STRUCTURE /sbxc/zckp_invh
                       data_base LIKE sy-datum.

  PERFORM init_bbseg  USING bbseg.           "#EC CI_FLDEXT_OK[2610650]

  CASE e_ucomm.
    WHEN 'F-41'.
      bbseg-newbs = '21'.
    WHEN 'F-43'.
      bbseg-newbs = '31'.
    WHEN OTHERS.
      IF it_cab-doctypesaphety = '4' OR it_cab-doctypesaphety = '381'. "Nota de crédito
        bbseg-newbs = '21'.
        IF it_cab-item_text IS INITIAL.
          bbseg-sgtxt = TEXT-020. "'V/ Nota de Crédito'.
        ELSE.
          bbseg-sgtxt = it_cab-item_text.
        ENDIF.
      ELSE.
        bbseg-newbs = '31'.
        IF it_cab-item_text IS INITIAL.
          bbseg-sgtxt = TEXT-021. "'V/ Fatura'.
        ELSE.
          bbseg-sgtxt = it_cab-item_text.
        ENDIF.
      ENDIF.
  ENDCASE.

  bbseg-newko     = it_cab-vendor.

  DATA: lv_amount TYPE bbseg-wrbtr.
  WRITE it_cab-gross_amount TO  lv_amount DECIMALS 2 CURRENCY it_cab-currency.
  WRITE lv_amount TO  bbseg-wrbtr CURRENCY it_cab-currency .

  IF it_cab-zterm NE space.
    bbseg-zterm     = it_cab-zterm.
  ENDIF.
*  WRITE it_cab-doc_date TO bbseg-zfbdt.
* CCF 31.01.2022
  WRITE data_base TO bbseg-zfbdt.
  bbseg-zlspr = it_cab-pmnt_block.
  bbseg-zuonr = it_cab-alloc_nmbr.
  bbseg-sgtxt = it_cab-item_text.

  TRANSFER bbseg TO ds_name.

ENDFORM.                    " BBSEG_FORN
*&---------------------------------------------------------------------*
*&      Form  BBSEG_LIN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_IT_LIN  text
*----------------------------------------------------------------------*
FORM bbseg_lin USING e_ucomm
                     it_lin STRUCTURE /sbxc/zckp_invi.


  PERFORM init_bbseg  USING bbseg.           "#EC CI_FLDEXT_OK[2610650]

*  CASE e_ucomm.
*    WHEN 'F-41' OR 'F-41B'.
  IF it_lin-anln1 IS NOT INITIAL.
    IF it_lin-db_cr_ind EQ 'S'.
      bbseg-newbs = '70'.
    ENDIF.
    IF it_lin-db_cr_ind EQ 'H'.
      bbseg-newbs = '75'.
    ENDIF.
  ELSE.
    IF it_lin-db_cr_ind EQ 'S'.
      bbseg-newbs = '40'.
    ENDIF.
    IF it_lin-db_cr_ind EQ 'H'.
      bbseg-newbs = '50'.
    ENDIF.
  ENDIF.
*    WHEN 'F-43' OR 'F-43B'.
*      IF it_lin-db_cr_ind EQ 'S'.
*        bbseg-newbs = '40'.
*      ENDIF.
*      IF it_lin-db_cr_ind EQ 'H'.
*        bbseg-newbs = '50'.
*      ENDIF.
*    WHEN OTHERS.
*      IF it_lin-db_cr_ind EQ 'S'.
*        bbseg-newbs = '81'.
*      ENDIF.
*      IF it_lin-db_cr_ind EQ 'H'.
*        bbseg-newbs = '91'.
*      ENDIF.
*  ENDCASE.

  bbseg-newko(17) = it_lin-gl_account.
  IF  bbseg-newko IS INITIAL.
    bbseg-newko(17) = it_lin-anln1. "Imobilizado
    IF  bbseg-newko IS INITIAL.
      bbseg-newko = '0'.
    ENDIF.
  ENDIF.

  DATA: lv_amount TYPE bbseg-wrbtr.
  WRITE it_lin-item_amount TO  lv_amount DECIMALS 2 CURRENCY bbkpf-waers.
  WRITE lv_amount TO  bbseg-wrbtr CURRENCY bbkpf-waers .

*  bbseg-wrbtr     = '*'.
  bbseg-mwskz     = it_lin-tax_code_sap.
  bbseg-kostl     = it_lin-costcenter.
  bbseg-matnr     = it_lin-matnr.
  bbseg-zuonr = it_lin-zuonr.
  bbseg-sgtxt = it_lin-item_text.
  TRANSFER bbseg TO ds_name.                 "#EC CI_FLDEXT_OK[2610650]

ENDFORM.                    " BBSEG_LIN
*&---------------------------------------------------------------------*
*&      Form  INIT_BBSEG
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_BBSEG  text
*----------------------------------------------------------------------*
FORM init_bbseg  USING     bbseg.
  bbseg = i_bbseg.
ENDFORM.                    " INIT_BBSEG
*&---------------------------------------------------------------------*
*&      Form  EXIBE_DOC_FI
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_DOC_FI  text
*      -->P_HEADER_COMP_CODE  text
*      -->P_HEADER_ANO_LANC  text
*----------------------------------------------------------------------*
FORM exibe_doc_fi  USING    doc_fi comp_code ano_lanc.

  SET PARAMETER ID: 'BLN' FIELD doc_fi,
                    'BUK' FIELD comp_code,
                    'GJR' FIELD ano_lanc.

  CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.

ENDFORM.                    " EXIBE_DOC_FI
*&---------------------------------------------------------------------*
*&      Form  EXIBE_DOC_LO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_DOC_ESTORNO  text
*      -->P_HEADER_COMP_CODE  text
*      -->P_HEADER_ANO_LANC  text
*----------------------------------------------------------------------*
FORM exibe_doc_lo  USING    doc_lo comp_code ano_lanc.


  SET PARAMETER ID: 'RBN' FIELD doc_lo,
                    'GJR' FIELD ano_lanc.

  CALL TRANSACTION 'MIR4' AND SKIP FIRST SCREEN.

ENDFORM.                    " EXIBE_DOC_LO
*&---------------------------------------------------------------------*
*&      Form  VALIDA_IMP_MULT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_ITEM  text
*      -->P_HEADER  text
*----------------------------------------------------------------------*
FORM valida_imp_mult  TABLES itemdata STRUCTURE /sbxc/zckp_invi
         USING headerdata LIKE /sbxc/zckp_invh.

  DATA: gl_account LIKE itemdata-gl_account,
        anln1      LIKE itemdata-anln1,
        anln2      LIKE itemdata-anln2,
        costcenter LIKE itemdata-costcenter,
        wbs_elem   LIKE itemdata-wbs_elem,
        orderid    LIKE itemdata-orderid.

  CLEAR indice.
  LOOP AT itemdata.
    ADD 1 TO indice.

    SELECT sakto kostl ps_psp_pnr aufnr anln1 anln2 FROM ekkn INTO
    (gl_account, costcenter, wbs_elem, orderid, anln1, anln2)
    WHERE ebeln = itemdata-po_number
                        AND ebelp = itemdata-po_item.

      IF sy-dbcnt > 1.
        itemdata-clas_mult = '@1E@'.
        CLEAR  itemdata-gl_account.
        EXIT.
      ENDIF.

    ENDSELECT.
    IF itemdata-gl_account IS INITIAL. itemdata-gl_account = gl_account. ENDIF.

    IF itemdata-anln1 IS INITIAL. itemdata-anln1 = anln1. ENDIF.
    IF itemdata-anln2 IS INITIAL. itemdata-anln2 = anln2. ENDIF.

    IF itemdata-costcenter IS INITIAL.
      itemdata-costcenter = costcenter.
    ENDIF.
    IF itemdata-wbs_elem IS INITIAL.
      itemdata-wbs_elem = wbs_elem.
    ENDIF.
    IF itemdata-orderid IS INITIAL. itemdata-orderid = orderid. ENDIF.
    MODIFY itemdata INDEX indice TRANSPORTING clas_mult gl_account costcenter wbs_elem orderid anln1 anln2.
    CLEAR:  gl_account, costcenter, wbs_elem, orderid, anln1, anln2.
  ENDLOOP.

ENDFORM.                    " VALIDA_IMP_MULT
**&---------------------------------------------------------------------*
**&      Form  VALIDA_IVA
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->P_ITEM  text
**      -->P_HEADER  text
**----------------------------------------------------------------------*
*FORM valida_iva  TABLES  itemdata STRUCTURE /sbxc/zckp_invi
*         USING headerdata LIKE /sbxc/zckp_invh.
*
*  DATA: x_country LIKE t001-land1.
*  DATA: lt_ftaxp TYPE STANDARD TABLE OF ftaxp WITH HEADER LINE.
*
*
*  LOOP AT itemdata.
*    SELECT SINGLE werks FROM ekpo INTO itemdata-stge_loc
*      WHERE ebeln = itemdata-po_number
*        AND ebelp = itemdata-po_item.
*
*    indice = sy-tabix.
*    IF  headerdata-doc_type <> 'NC'. "Se se tratar de uma factura
*
*      SELECT SINGLE iva_fatura INTO itemdata-tax_code_sap
*        FROM /sbxc/zckp_miva
*        WHERE comp_code = headerdata-comp_code
*          AND stge_loc = itemdata-stge_loc
*          AND serv_saphety = itemdata-gl_account.
*
*      IF sy-subrc NE 0.
*        SELECT SINGLE iva_fatura INTO itemdata-tax_code_sap
*          FROM /sbxc/zckp_miva
*          WHERE comp_code = headerdata-comp_code
*            AND serv_saphety = itemdata-gl_account.
*        IF sy-subrc NE 0.
*          SELECT SINGLE iva_fatura INTO itemdata-tax_code_sap
*            FROM /sbxc/zckp_miva
*            WHERE serv_saphety = itemdata-gl_account.
*        ENDIF.
*
*      ENDIF.
*    ELSE. " se se tratar de uma NC
*      SELECT SINGLE iva_nc INTO itemdata-tax_code_sap
*        FROM /sbxc/zckp_miva
*        WHERE comp_code = headerdata-comp_code
*          AND stge_loc = itemdata-stge_loc
*          AND serv_saphety = itemdata-gl_account.
*      IF sy-subrc NE 0.
*        SELECT SINGLE iva_nc INTO itemdata-tax_code_sap
*          FROM /sbxc/zckp_miva
*          WHERE comp_code = headerdata-comp_code
*            AND serv_saphety = itemdata-gl_account.
*        IF sy-subrc NE 0.
*          SELECT SINGLE iva_nc INTO itemdata-tax_code_sap
*            FROM /sbxc/zckp_miva
*            WHERE serv_saphety = itemdata-gl_account.
*        ENDIF.
*      ENDIF.
*    ENDIF.
** valida taxa de IVA
** valida pais da empresa
*    SELECT SINGLE land1 INTO x_country FROM  t001
*      WHERE bukrs =  headerdata-comp_code.
*
*    CALL FUNCTION 'GET_TAX_PERCENTAGE'
*      EXPORTING
*        aland   = x_country
*        datab   = sy-datum
*        mwskz   = itemdata-tax_code_sap
*        txjcd   = '0'
*      TABLES
*        t_ftaxp = lt_ftaxp.
*
*    SORT lt_ftaxp BY kbetr.
*
*    LOOP AT lt_ftaxp.
*      MOVE lt_ftaxp-kbetr TO itemdata-tax_imposto_sap.
*      DIVIDE itemdata-tax_imposto_sap BY 10.
*itemdata-tax_amount = itemdata-item_amount *  itemdata-tax_imposto_sap / 100.
*      MODIFY itemdata INDEX indice TRANSPORTING tax_code_sap
*      tax_imposto_sap.
*    ENDLOOP.
*  ENDLOOP.
*
*ENDFORM.                    " VALIDA_IVA
**&---------------------------------------------------------------------*
**&      Form  VALIDA_QTD_FATURA2
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->P_ITEM  text
**      -->P_HEADER  text
**----------------------------------------------------------------------*
*FORM valida_qtd_fatura2   TABLES itemdata STRUCTURE /sbxc/zckp_invi
*         USING headerdata LIKE /sbxc/zckp_invh.
*
*
** Declaração de variaveis para rotina de calculo de qtd a facturar
*  DATA: lt_xekbe  TYPE TABLE OF ekbe,
**        ls_xekbe  TYPE ekbe,
*        lt_xekbes TYPE TABLE OF ekbes,
*        ls_xekbes TYPE ekbes.
*
*  DATA: qtd_entrada     LIKE ls_xekbes-wemng,
*        val_entrada     LIKE ls_xekbes-wewwr,
*        qtd_facturada   LIKE ls_xekbes-remng,
*        val_facturado   LIKE ls_xekbes-rewwr,
*        qtd_ped         LIKE ls_xekbes-wemng,
*        valor_ped       LIKE ls_xekbes-wewwr,
*        qtd_remanesc    LIKE ls_xekbes-remng,
*        dc_posterior(1),
*        indice(1).
*
*
*  LOOP AT itemdata.
*
*    indice = sy-tabix.
*    itemdata-quant_fact = itemdata-quantity.
** para pedido item verifica quantidades facturadas até ao momento
*    REFRESH: lt_xekbes,  lt_xekbe.
*    CLEAR: lt_xekbes,  lt_xekbe.
*
*    CALL FUNCTION 'ME_READ_HISTORY'
*      EXPORTING
*        ebeln  = itemdata-po_number
*        ebelp  = itemdata-po_item
*        webre  = 'X'
*      TABLES
*        xekbe  = lt_xekbe
*        xekbes = lt_xekbes.
*
*    READ TABLE lt_xekbes INTO ls_xekbes
*    WITH  KEY ebelp = itemdata-po_item
*              zekkn = '00'.
*
*    qtd_entrada  =  ls_xekbes-wemng.
*    val_entrada  = ls_xekbes-wewwr.
*    qtd_facturada = ls_xekbes-remng.
*    val_facturado = ls_xekbes-rewwr.
*
*    IF headerdata-doc_type = 'NC'.
*      itemdata-quant_fact  = qtd_facturada.
*      itemdata-dc_posterior = 'X'.
*    ELSE.
*      SELECT SINGLE menge brtwr INTO (qtd_ped, valor_ped) FROM ekpo
*             WHERE ebeln = itemdata-po_number AND
*                   ebelp = itemdata-po_item.
*
** 1ª Factura e factura final
*      IF  qtd_facturada = 0.
*        IF itemdata-fact_final = 'X'.
*          itemdata-dc_posterior = ' '.
*          itemdata-quant_fact = qtd_ped.
*
*        ELSE. " Factura não final
*          itemdata-dc_posterior = ' '.
*
*          PERFORM calcula_x USING qtd_ped valor_ped
*                           itemdata-item_amount qtd_facturada
*                           CHANGING itemdata-quant_fact qtd_remanesc.
*        ENDIF.
*
*      ELSEIF qtd_facturada < qtd_ped  . " qtd_facturada <> 0.
*        itemdata-dc_posterior = ' '.
*        PERFORM calcula_x USING qtd_ped valor_ped
*                  itemdata-item_amount qtd_facturada
*                  CHANGING itemdata-quant_fact qtd_remanesc.
*
*        IF itemdata-fact_final = 'X'.
*          itemdata-quant_fact = qtd_remanesc.
*        ELSE.
*          IF itemdata-quant_fact > qtd_remanesc.
*            itemdata-quant_fact = qtd_remanesc.
*          ENDIF.
*        ENDIF.
*      ELSEIF qtd_facturada GE  qtd_ped.
*        itemdata-dc_posterior = 'X'. " --> Debito posterior
*        itemdata-quant_fact = qtd_ped.
*      ENDIF.
*
*    ENDIF.
*   MODIFY  itemdata INDEX indice TRANSPORTING quant_fact dc_posterior  .
*  ENDLOOP.
*
*ENDFORM.                    " VALIDA_QTD_FATURA2
*&---------------------------------------------------------------------*
*&      Form  valida_qtd_fatura_sc
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->ITEMDATA   text
*      -->HEADERDATA text
*----------------------------------------------------------------------*
FORM valida_qtd_fatura_sc   TABLES itemdata STRUCTURE /sbxc/zckp_invi
         USING headerdata LIKE /sbxc/zckp_invh.

* Declaração de variaveis para rotina de calculo de qtd a facturar
  DATA: lt_xekbe  TYPE TABLE OF ekbe,
*        ls_xekbe  TYPE ekbe,
        lt_xekbes TYPE TABLE OF ekbes,
        t_ekpo    TYPE TABLE OF ekpo,
        ls_xekbes TYPE ekbes.
  DATA: indice_i TYPE sy-tabix.
  DATA: "qtd_entrada     LIKE ls_xekbes-wemng,
*        val_entrada     LIKE ls_xekbes-wewwr,
*        qtd_facturada   LIKE ls_xekbes-remng,
*        val_facturado   LIKE ls_xekbes-rewwr,
*        qtd_ped         LIKE ls_xekbes-wemng,
*        valor_ped       LIKE ls_xekbes-wewwr,
*        qtd_remanesc    LIKE ls_xekbes-remng,
*        dc_posterior(1),
*        indice          LIKE sy-tabix,
    ped_item(20).

  DATA: lv_error TYPE c.

  LOOP AT itemdata.
    indice_i = sy-tabix. "ODC - 19_05_2020
* para pedido item verifica quantidades facturadas até ao momento
    REFRESH: lt_xekbes,  lt_xekbe.
    CLEAR: lt_xekbes,  lt_xekbe, lv_error.
    IF itemdata-po_number IS NOT INITIAL.
      REFRESH t_ekpo.
      IF itemdata-po_item IS INITIAL.
        SELECT * FROM ekpo INTO CORRESPONDING FIELDS OF TABLE t_ekpo
                 WHERE ebeln = itemdata-po_number AND
                        loekz = ' ' AND
                        netpr <> '0' ORDER BY ebelp.
      ELSE.
        SELECT * FROM ekpo INTO CORRESPONDING FIELDS OF TABLE t_ekpo
             WHERE ebeln = itemdata-po_number AND
                   ebelp = itemdata-po_item AND
                    loekz = ' ' AND
                    netpr <> '0'
          ORDER BY PRIMARY KEY.
      ENDIF.

      IF itemdata-po_item IS NOT INITIAL . "ODC - 08_06_2021
* Propoe Entradas de Material/Serviços
        CLEAR wa_rbkpv.
        REFRESH: ydrseg , xmsel_best, xmsel_lifs,
        xmsel_frbr, xmsel_werk, xmsel_erfb, xmsel_tran,
        t_errprot, t_ebelntab.

        MOVE: headerdata-pstng_date(4) TO wa_rbkpv-gjahr.
        wa_xmsel_best-gjahr =  wa_rbkpv-gjahr.
        wa_rbkpv-bldat = headerdata-pstng_date.
        wa_rbkpv-budat = headerdata-pstng_date.
        wa_rbkpv-bukrs = headerdata-comp_code.
        wa_rbkpv-blart = 'RE'.
        MOVE 'X' TO: wa_rbkpv-xzuordli, wa_rbkpv-xzuordrt, wa_rbkpv-xbest.
        wa_xmsel_best-ebeln = itemdata-po_number.
        wa_xmsel_best-ebelp = itemdata-po_item.
        IF headerdata-xware_bnk = '1'.
          "Mercadorias.
          MOVE 'X' TO: wa_rbkpv-xware.
        ELSEIF headerdata-xware_bnk = '2'.
          "Custos Complementares
          MOVE 'X' TO: wa_rbkpv-xbnk.
        ELSEIF headerdata-xware_bnk = '3'.
          "Mercadorias + Custos Complementares
          MOVE 'X' TO: wa_rbkpv-xware, wa_rbkpv-xbnk.
        ENDIF.

        APPEND wa_rbkpv TO rbkpv. ", xmsel_best.
        APPEND wa_xmsel_best TO xmsel_best.

        REFRESH t_errprot.
        CALL FUNCTION 'MRM_ASSIGNMENT'
          EXPORTING
            i_no_material_lock = 'X' "ODC - 25_07_2022
          TABLES
            t_drseg            = ydrseg
            t_rbselbest        = xmsel_best
            t_rbsellifs        = xmsel_lifs
            t_rbselfrbr        = xmsel_frbr
            t_rbselwerk        = xmsel_werk
            t_rbselerfb        = xmsel_erfb
            t_rbseltran        = xmsel_tran
            t_errprot          = t_errprot
            t_ebelntab         = t_ebelntab
          CHANGING
            c_rbkpv            = wa_rbkpv
            t_limit            = xlimit
          EXCEPTIONS
            error_message      = 01.
        IF sy-subrc NE 0.
          lv_error = abap_true.
          EXIT.                                         "#EC CI_NOORDER
        ENDIF.

        DO 10 TIMES.
          CALL FUNCTION 'DEQUEUE_EMEKKOS'
            EXPORTING
              ebeln = itemdata-po_number.
        ENDDO.

        DO 10 TIMES.
          CALL FUNCTION 'DEQUEUE_EMEKPOE'
            EXPORTING
              ebeln = itemdata-po_number
              ebelp = itemdata-po_item.
        ENDDO.

*"ODC - 12_07_2022
*LOOP AT YDRSEG ASSIGNING FIELD-SYMBOL(<fsy>) WHERE ebeln = itemdata-po_number and ebelp = itemdata-po_item.
*DO 3 TIMES.
*  CALL FUNCTION 'DEQUEUE_EMMBEWE'
*   EXPORTING
*     MATNR           = itemdata-MATNR
*     BWKEY           = <fsy>-werks
*            .
*  ENDDO.
*ENDLOOP.
*"Fim ODC - 12_07_2022

        IF ydrseg[] IS INITIAL.
          CONCATENATE itemdata-po_number ekpo-ebelp INTO ped_item SEPARATED BY space.
          MESSAGE s035(m8) WITH  ped_item.
        ENDIF.
        LOOP AT ydrseg INTO wa_ydrseg WHERE ebeln = itemdata-po_number AND
                                            ebelp = itemdata-po_item.
          CLEAR: itemdata-ref_doc, itemdata-ref_doc_year, itemdata-ref_doc_item , itemdata-cond_type.
          CLEAR: item_aux-ref_doc, item_aux-ref_doc_year, item_aux-ref_doc_item,  item_aux-cond_type.

          itemdata-ref_doc = wa_ydrseg-lfbnr.
          itemdata-ref_doc_year = wa_ydrseg-lfgja.
          itemdata-ref_doc_item = wa_ydrseg-lfpos.

          READ  TABLE item_aux WITH  KEY
              po_number = itemdata-po_number
              po_item = itemdata-po_item
              ref_doc = itemdata-ref_doc
              ref_doc_year = itemdata-ref_doc_year
              ref_doc_item = itemdata-ref_doc_item.
          IF sy-subrc NE 0.


            READ  TABLE item_aux WITH  KEY
                po_number = itemdata-po_number
                po_item = itemdata-po_item.
            IF sy-subrc EQ 0.
              item_aux-ref_doc = wa_ydrseg-lfbnr.
              item_aux-ref_doc_year = wa_ydrseg-lfgja.
              item_aux-ref_doc_item = wa_ydrseg-lfpos.
              IF itemdata-quant_fact > 0.
                item_aux-quantity = itemdata-quant_fact.
              ENDIF.


              READ TABLE t_ekpo ASSIGNING FIELD-SYMBOL(<fs_ep>) WITH KEY ebeln = itemdata-po_number
                                                                         ebelp = itemdata-po_item.
              IF sy-subrc EQ 0.
                item_aux-po_unit = <fs_ep>-meins.
                item_aux-item_text = <fs_ep>-txz01.
                item_aux-idnlf = <fs_ep>-idnlf.
              ENDIF.


              CLEAR: item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2.
              SELECT SINGLE sakto kostl ps_psp_pnr aufnr anln1 anln2 FROM ekkn INTO "#EC CI_NOORDER
                (item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem,
                 item_aux-orderid, item_aux-anln1, item_aux-anln2)
                WHERE ebeln = itemdata-po_number
                AND ebelp = itemdata-po_item.

              MODIFY item_aux INDEX sy-tabix.
            ENDIF.

          ENDIF.
          CLEAR wa_ydrseg.
        ENDLOOP.
      ELSE. " Fim - ODC - 08_06_2021

        LOOP AT t_ekpo INTO ekpo.
          CLEAR: itemdata-ref_doc, itemdata-ref_doc_year, itemdata-ref_doc_item, itemdata-cond_type .
*        CLEAR itemdata-tax_code_sap. "ODC - 04_06_2020
          CLEAR: item_aux-ref_doc, item_aux-ref_doc_year, item_aux-ref_doc_item, item_aux-cond_type .
          REFRESH ydrseg.
          IF ekpo-webre = ' '.

            IF headerdata-xware_bnk NE '2'. "Se existirem itens de mercadorias
              REFRESH: lt_xekbes,  lt_xekbe.
              CLEAR: lt_xekbes,  lt_xekbe.

              CALL FUNCTION 'ME_READ_HISTORY'
                EXPORTING
                  ebeln  = itemdata-po_number
                  ebelp  = ekpo-ebelp
                  webre  = 'X'
                TABLES
                  xekbe  = lt_xekbe
                  xekbes = lt_xekbes.

              READ TABLE lt_xekbes INTO ls_xekbes
              WITH  KEY ebelp = ekpo-ebelp
                        zekkn = '00'.
* se não houver historico então coloca os montantes e quantidades do pedido
              IF ( ls_xekbes-wemng = 0 AND ls_xekbes-remng = 0 ) OR sy-subrc <> 0.
                qtd_remanesc = ekpo-menge.
                valor_ped = ekpo-netwr.
* se pisco de em = ' ' então qtd proposta = qt. pedido -  qt. fat
* valor pu * qt propposta
              ELSEIF ekpo-wepos = ' '.
                qtd_remanesc = ekpo-menge - ls_xekbes-remng.
                valor_ped = ekpo-netpr * qtd_remanesc.
              ELSE.
                qtd_remanesc = ls_xekbes-wemng -  ls_xekbes-remng.
                valor_ped =  ( ( ekpo-netwr / ekpo-menge ) * qtd_remanesc )  - ls_xekbes-rewwr.
                "( ( ekpo-netpr / ekpo-menge ) * qtd_remanesc )  - ls_xekbes-rewwr. "ls_xekbes-wewwr - ls_xekbes-rewwr.
              ENDIF.
              IF valor_ped < 0.
                valor_ped = 0.
              ENDIF.
              IF qtd_remanesc < 0.
                qtd_remanesc = 0.
              ENDIF.

* se operação é NC
              IF headerdata-doctypesaphety = 'NC' OR headerdata-doctypesaphety  = 'CREDITNOTE'
              OR headerdata-doctypesaphety  = '4'
              OR headerdata-doctypesaphety = '381'. "ODC - 12_03_2021
                qtd_remanesc = ls_xekbes-remng.
                valor_ped = ls_xekbes-rewwr.
              ENDIF.

              itemdata-quantity = qtd_remanesc.
***Conversão quantidade em UPP
              itemdata-bpmng = ( itemdata-quantity * ekpo-bpumz ) / ekpo-bpumn.
              itemdata-item_amount =  valor_ped .
              itemdata-po_unit = ekpo-meins.
              itemdata-bprme   = ekpo-bprme.
              itemdata-matnr   = ekpo-matnr.

              "ODC - 19_05_2020
              IF itemdata-tax_code_sap IS INITIAL OR ekpo-mwskz IS NOT INITIAL.
                itemdata-tax_code_sap = ekpo-mwskz.
                MODIFY itemdata INDEX indice_i TRANSPORTING tax_code_sap.
              ENDIF.
              "Fim ODC 19_05_2020

              READ  TABLE item_aux WITH  KEY
                                 po_number = itemdata-po_number
                                 po_item = ekpo-ebelp
                                 ref_doc = itemdata-ref_doc
                                 ref_doc_year = itemdata-ref_doc_year
                                 ref_doc_item = itemdata-ref_doc_item.
              IF sy-subrc NE 0.
                ADD 1 TO   indice .

                IF itemdata-item_amount  < 0.
                  itemdata-item_amount = 0.
                ENDIF.

                MOVE-CORRESPONDING itemdata TO item_aux.
                item_aux-po_item = ekpo-ebelp.

                CLEAR: item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2.
                SELECT SINGLE sakto kostl ps_psp_pnr aufnr anln1 anln2 FROM ekkn INTO "#EC CI_NOORDER
                  (item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem,
                  item_aux-orderid, item_aux-anln1, item_aux-anln2)
                  WHERE ebeln = itemdata-po_number
                  AND ebelp = ekpo-ebelp.

                MOVE-CORRESPONDING headerdata TO item_aux.
                item_aux-invoice_doc_item = indice.
                item_aux-item_text = ekpo-txz01.
                item_aux-po_number = itemdata-po_number.
                item_aux-po_unit = ekpo-meins. "ODC - 04_03_2021
                item_aux-idnlf = ekpo-idnlf.
                IF qtd_remanesc = 0 AND valor_ped = 0.
                ELSE.
                  APPEND item_aux.
                ENDIF.
                "ODC - 19_05_2020
              ELSE.
                item_aux-po_unit = ekpo-meins. "ODC - 04_03_2021
                item_aux-tax_code_sap = itemdata-tax_code_sap.
                MODIFY item_aux INDEX sy-tabix TRANSPORTING tax_code_sap.
                "Fim ODC - 19_05_2020
              ENDIF.
            ENDIF.
            IF headerdata-xware_bnk NE '1'. "Se for para apresentar custos complementares.
              CLEAR wa_rbkpv.
              REFRESH: ydrseg , xmsel_best, xmsel_lifs,
              xmsel_frbr, xmsel_werk, xmsel_erfb, xmsel_tran,
              t_errprot, t_ebelntab.

              MOVE: headerdata-pstng_date(4) TO wa_rbkpv-gjahr.
              wa_xmsel_best-gjahr =  wa_rbkpv-gjahr.
              wa_rbkpv-bldat = headerdata-pstng_date.
              wa_rbkpv-budat = headerdata-pstng_date.
              wa_rbkpv-bukrs = headerdata-comp_code.
              wa_rbkpv-blart = 'RE'.
              MOVE 'X' TO: wa_rbkpv-xzuordli, wa_rbkpv-xzuordrt, wa_rbkpv-xbest.
              wa_xmsel_best-ebeln = itemdata-po_number.
              wa_xmsel_best-ebelp = ekpo-ebelp.
              MOVE 'X' TO: wa_rbkpv-xbnk.
              APPEND wa_rbkpv TO rbkpv.
              APPEND wa_xmsel_best TO xmsel_best.

              REFRESH t_errprot.
              CALL FUNCTION 'MRM_ASSIGNMENT'
                EXPORTING
                  i_no_material_lock = 'X' "ODC - 25_07_2022
                TABLES
                  t_drseg            = ydrseg
                  t_rbselbest        = xmsel_best
                  t_rbsellifs        = xmsel_lifs
                  t_rbselfrbr        = xmsel_frbr
                  t_rbselwerk        = xmsel_werk
                  t_rbselerfb        = xmsel_erfb
                  t_rbseltran        = xmsel_tran
                  t_errprot          = t_errprot
                  t_ebelntab         = t_ebelntab
                CHANGING
                  c_rbkpv            = wa_rbkpv
                  t_limit            = xlimit
                EXCEPTIONS
                  error_message      = 01.
              IF sy-subrc NE 0.
                lv_error =  abap_true.
*              READ TABLE t_errprot WITH KEY msgty = 'E'.
*MESSAGE ID t_errprot-msgid TYPE 'S' NUMBER t_errprot-msgno DISPLAY LIKE 'E'
*WITH t_errprot-msgv1 t_errprot-msgv2 t_errprot-msgv3 t_errprot-msgv4.
                EXIT.                                   "#EC CI_NOORDER
              ENDIF.

              DO 10 TIMES.
                CALL FUNCTION 'DEQUEUE_EMEKKOS'
                  EXPORTING
                    ebeln = itemdata-po_number.
              ENDDO.

              DO 10 TIMES.
                CALL FUNCTION 'DEQUEUE_EMEKPOE'
                  EXPORTING
                    ebeln = itemdata-po_number
                    ebelp = ekpo-ebelp.
*                EXCEPTIONS
*                  foreign_lock   = 2
*                  system_failure = 3.
              ENDDO.

*              "ODC - 12_07_2022
**READ TABLE YDRSEG ASSIGNING FIELD-SYMBOL(<fsz>) with key ebeln = itemdata-po_number ebelp = itemdata-po_item.
**IF sy-subrc eq 0.
**DO 10 TIMES.
**  CALL FUNCTION 'DEQUEUE_EMMBEWE'
**   EXPORTING
**     MATNR           = itemdata-MATNR
**     BWKEY           = <fsz>-werks
**            .
**  ENDDO.
**ENDIF.
*LOOP AT YDRSEG ASSIGNING FIELD-SYMBOL(<fsz>) WHERE ebeln = itemdata-po_number and ebelp = itemdata-po_item.
*DO 3 TIMES.
*  CALL FUNCTION 'DEQUEUE_EMMBEWE'
*   EXPORTING
*     MATNR           = itemdata-MATNR
*     BWKEY           = <fsz>-werks
*            .
*  ENDDO.
*ENDLOOP.
*"Fim ODC - 12_07_2022

              IF ydrseg[] IS INITIAL.
                CONCATENATE itemdata-po_number ekpo-ebelp INTO ped_item SEPARATED BY space.
                MESSAGE s035(m8) WITH  ped_item.
              ENDIF.
              LOOP AT ydrseg INTO wa_ydrseg WHERE ebeln = itemdata-po_number AND
                                                  ebelp = ekpo-ebelp.
                CLEAR: itemdata-ref_doc, itemdata-ref_doc_year, itemdata-ref_doc_item , itemdata-cond_type.
                CLEAR: item_aux-ref_doc, item_aux-ref_doc_year, item_aux-ref_doc_item,  item_aux-cond_type.
                IF ekpo-weunb = 'X'.
                  itemdata-item_amount = ( ( ekpo-netwr / ekpo-menge ) * wa_ydrseg-wemng ) - wa_ydrseg-refwr.
                ELSE.
                  itemdata-item_amount = wa_ydrseg-wewwr - wa_ydrseg-refwr.
                  IF itemdata-item_amount IS NOT INITIAL.
                    itemdata-item_amount = ( ( ekpo-netwr / ekpo-menge ) * wa_ydrseg-wemng ) - wa_ydrseg-refwr.
                  ENDIF.
                ENDIF.
                itemdata-costcenter = wa_ydrseg-kostl.
                itemdata-po_unit =  wa_ydrseg-meins.
                itemdata-bprme   =  wa_ydrseg-bprme.
                IF itemdata-tax_code_sap IS INITIAL.
                  itemdata-tax_code_sap = wa_ydrseg-mwskz.
                ENDIF.
                itemdata-ref_doc = wa_ydrseg-lfbnr.
                itemdata-ref_doc_year = wa_ydrseg-lfgja.
                itemdata-ref_doc_item = wa_ydrseg-lfpos.
                itemdata-quantity = wa_ydrseg-wemng - wa_ydrseg-remng.
                itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
                itemdata-cond_type = wa_ydrseg-kschl.
                IF tipo_doc = 'NC' OR tipo_doc = 'CREDITNOTE' OR headerdata-doctypesaphety  = '4'
                OR headerdata-doctypesaphety  = 'NC' OR  headerdata-doctypesaphety  = 'CREDITNOTE'
                  OR headerdata-doctypesaphety = '381'. "ODC - 12_03_2021.
                  itemdata-quantity = wa_ydrseg-remng.
                  itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
                  itemdata-item_amount = wa_ydrseg-refwr.
                ENDIF.
                IF itemdata-item_amount  < 0.
                  itemdata-item_amount = 0.
                ENDIF.

                IF itemdata-quantity  < 0.
                  itemdata-quantity = 0.
                  itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
                ENDIF.
                READ  TABLE item_aux WITH  KEY
                       po_number = itemdata-po_number
                       po_item = ekpo-ebelp
                       ref_doc = itemdata-ref_doc
                       ref_doc_year = itemdata-ref_doc_year
                       ref_doc_item = itemdata-ref_doc_item
                       cond_type    = itemdata-cond_type.
                IF sy-subrc NE 0.
                  IF itemdata-item_amount  < 0.
                    itemdata-item_amount = 0.
                  ENDIF.

                  MOVE-CORRESPONDING itemdata TO item_aux.
                  item_aux-po_item = ekpo-ebelp.
                  item_aux-po_unit = ekpo-meins. "ODC - 04_03_2021

                  CLEAR: item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2.
                  SELECT SINGLE sakto kostl ps_psp_pnr aufnr anln1 anln2 FROM ekkn INTO "#EC CI_NOORDER
               (item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2)
               WHERE ebeln = itemdata-po_number
               AND ebelp = ekpo-ebelp.
                  MOVE-CORRESPONDING headerdata TO item_aux.
                  ADD 1 TO   indice .
                  item_aux-invoice_doc_item = indice.
                  item_aux-item_text = ekpo-txz01.
                  item_aux-po_number = itemdata-po_number.
                  item_aux-idnlf = ekpo-idnlf.
                  APPEND item_aux.

                ENDIF.
                CLEAR wa_ydrseg.
              ENDLOOP.
            ENDIF.
          ELSE." EF/EM  esta activo
* Propoe Entradas de Material/Serviços
            CLEAR wa_rbkpv.
            REFRESH: ydrseg , xmsel_best, xmsel_lifs,
            xmsel_frbr, xmsel_werk, xmsel_erfb, xmsel_tran,
            t_errprot, t_ebelntab.

            MOVE: headerdata-pstng_date(4) TO wa_rbkpv-gjahr.
            wa_xmsel_best-gjahr =  wa_rbkpv-gjahr.
            wa_rbkpv-bldat = headerdata-pstng_date.
            wa_rbkpv-budat = headerdata-pstng_date.
            wa_rbkpv-bukrs = headerdata-comp_code.
            wa_rbkpv-blart = 'RE'.
            MOVE 'X' TO: wa_rbkpv-xzuordli, wa_rbkpv-xzuordrt, wa_rbkpv-xbest.
            wa_xmsel_best-ebeln = itemdata-po_number.
            wa_xmsel_best-ebelp = ekpo-ebelp.
            IF headerdata-xware_bnk = '1'.
              "Mercadorias.
              MOVE 'X' TO: wa_rbkpv-xware.
            ELSEIF headerdata-xware_bnk = '2'.
              "Custos Complementares
              MOVE 'X' TO: wa_rbkpv-xbnk.
            ELSEIF headerdata-xware_bnk = '3'.
              "Mercadorias + Custos Complementares
              MOVE 'X' TO: wa_rbkpv-xware, wa_rbkpv-xbnk.
            ENDIF.

            APPEND wa_rbkpv TO rbkpv. ", xmsel_best.
            APPEND wa_xmsel_best TO xmsel_best.

            REFRESH t_errprot.
            CALL FUNCTION 'MRM_ASSIGNMENT'
              EXPORTING
                i_no_material_lock = 'X' "ODC - 25_07_2022
              TABLES
                t_drseg            = ydrseg
                t_rbselbest        = xmsel_best
                t_rbsellifs        = xmsel_lifs
                t_rbselfrbr        = xmsel_frbr
                t_rbselwerk        = xmsel_werk
                t_rbselerfb        = xmsel_erfb
                t_rbseltran        = xmsel_tran
                t_errprot          = t_errprot
                t_ebelntab         = t_ebelntab
              CHANGING
                c_rbkpv            = wa_rbkpv
                t_limit            = xlimit
              EXCEPTIONS
                error_message      = 01.
            IF sy-subrc NE 0.
              lv_error = abap_true.
*            READ TABLE t_errprot WITH KEY msgty = 'E'.
*            MESSAGE ID t_errprot-msgid TYPE 'S' NUMBER t_errprot-msgno DISPLAY LIKE 'E'
*            WITH t_errprot-msgv1 t_errprot-msgv2 t_errprot-msgv3 t_errprot-msgv4.
              EXIT.                                     "#EC CI_NOORDER
            ENDIF.

            DO 10 TIMES.
              CALL FUNCTION 'DEQUEUE_EMEKKOS'
                EXPORTING
                  ebeln = itemdata-po_number.
            ENDDO.

            DO 10 TIMES.
              CALL FUNCTION 'DEQUEUE_EMEKPOE'
                EXPORTING
                  ebeln = itemdata-po_number
                  ebelp = ekpo-ebelp.
*              EXCEPTIONS
*                foreign_lock   = 2
*                system_failure = 3.
            ENDDO.

*            "ODC - 12_07_2022
**READ TABLE YDRSEG ASSIGNING FIELD-SYMBOL(<fsw>) with key ebeln = itemdata-po_number ebelp = itemdata-po_item.
**IF sy-subrc eq 0.
**DO 10 TIMES.
**  CALL FUNCTION 'DEQUEUE_EMMBEWE'
**   EXPORTING
**     MATNR           = itemdata-MATNR
**     BWKEY           = <fsw>-werks
**            .
**  ENDDO.
**ENDIF.
*
*LOOP AT YDRSEG ASSIGNING FIELD-SYMBOL(<fsw>) WHERE ebeln = itemdata-po_number and ebelp = itemdata-po_item.
*DO 3 TIMES.
*  CALL FUNCTION 'DEQUEUE_EMMBEWE'
*   EXPORTING
*     MATNR           = itemdata-MATNR
*     BWKEY           = <fsw>-werks
*            .
*  ENDDO.
*ENDLOOP.
*"Fim ODC - 12_07_2022

            IF ydrseg[] IS INITIAL.
              CONCATENATE itemdata-po_number ekpo-ebelp INTO ped_item SEPARATED BY space.
              MESSAGE s035(m8) WITH  ped_item.
            ENDIF.

            LOOP AT ydrseg INTO wa_ydrseg WHERE ebeln = itemdata-po_number AND
                                                ebelp = ekpo-ebelp.
              CLEAR: itemdata-ref_doc, itemdata-ref_doc_year, itemdata-ref_doc_item , itemdata-cond_type.
              CLEAR: item_aux-ref_doc, item_aux-ref_doc_year, item_aux-ref_doc_item,  item_aux-cond_type.
              IF ekpo-weunb = 'X'.
                itemdata-item_amount = ( ekpo-netwr / ekpo-menge ) * wa_ydrseg-wemng .
              ELSE.
                itemdata-item_amount = wa_ydrseg-wewwr - wa_ydrseg-refwr.
                IF itemdata-item_amount IS NOT INITIAL.
                  itemdata-item_amount = ( ( ekpo-netwr / ekpo-menge ) * wa_ydrseg-wemng ) - wa_ydrseg-refwr.
                ENDIF.
              ENDIF.
              itemdata-costcenter = wa_ydrseg-kostl.
              itemdata-po_unit =  wa_ydrseg-meins.
              itemdata-bprme   =  wa_ydrseg-bprme.
              IF itemdata-tax_code_sap IS INITIAL.
                itemdata-tax_code_sap = wa_ydrseg-mwskz.
              ENDIF.
              itemdata-ref_doc = wa_ydrseg-lfbnr.
              itemdata-ref_doc_year = wa_ydrseg-lfgja.
              itemdata-ref_doc_item = wa_ydrseg-lfpos.
              itemdata-quantity = wa_ydrseg-wemng - wa_ydrseg-remng.
***Conversão quantidade em UPP
              itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
              itemdata-cond_type = wa_ydrseg-kschl.
              IF tipo_doc = 'NC' OR tipo_doc = 'CREDITNOTE' OR headerdata-doctypesaphety  = '4'
              OR headerdata-doctypesaphety  = 'NC' OR  headerdata-doctypesaphety  = 'CREDITNOTE'
                OR headerdata-doctypesaphety = '381'. "ODC - 12_03_2021.
                itemdata-quantity = wa_ydrseg-remng.
***Conversão quantidade em UPP
                itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
                itemdata-item_amount = wa_ydrseg-refwr.
              ENDIF.
              IF itemdata-item_amount  < 0.
                itemdata-item_amount = 0.
              ENDIF.
              itemdata-matnr   = ekpo-matnr.
              IF itemdata-quantity  < 0.
                itemdata-quantity = 0.
***Conversão quantidade em UPP
                itemdata-bpmng = ( itemdata-quantity * wa_ydrseg-bpumz ) / wa_ydrseg-bpumn.
              ENDIF.
              READ  TABLE item_aux WITH  KEY
                     po_number = itemdata-po_number
                     po_item = ekpo-ebelp
                     ref_doc = itemdata-ref_doc
                     ref_doc_year = itemdata-ref_doc_year
                     ref_doc_item = itemdata-ref_doc_item.
              IF sy-subrc NE 0.
                IF itemdata-item_amount  < 0.
                  itemdata-item_amount = 0.
                ENDIF.

                MOVE-CORRESPONDING itemdata TO item_aux.
                item_aux-po_item = ekpo-ebelp.
                item_aux-po_unit = ekpo-meins. "ODC - 04_03_2021
                CLEAR: item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2.
                SELECT SINGLE sakto kostl ps_psp_pnr aufnr anln1 anln2 FROM ekkn INTO "#EC CI_NOORDER
                  (item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem,
                   item_aux-orderid, item_aux-anln1, item_aux-anln2)
                  WHERE ebeln = itemdata-po_number
                  AND ebelp = ekpo-ebelp.

                MOVE-CORRESPONDING headerdata TO item_aux.
                ADD 1 TO   indice .
                item_aux-invoice_doc_item = indice.
                item_aux-item_text = ekpo-txz01.
                item_aux-po_number = itemdata-po_number.
                item_aux-idnlf = ekpo-idnlf.
                APPEND item_aux.
              ENDIF.
              CLEAR wa_ydrseg.
            ENDLOOP.

          ENDIF.
        ENDLOOP.
      ENDIF. "ODC - 08_06_2021
      IF lv_error EQ abap_true.
        READ TABLE t_errprot WITH KEY msgty = 'E'.
        IF sy-subrc EQ 0.
          MESSAGE ID t_errprot-msgid TYPE 'S' NUMBER t_errprot-msgno DISPLAY LIKE 'E'
          WITH t_errprot-msgv1 t_errprot-msgv2 t_errprot-msgv3 t_errprot-msgv4.
        ENDIF.
      ENDIF.

    ELSE.
      "ODC - 19_05_2020
      IF itemdata-tax_code_sap IS INITIAL.
        itemdata-tax_code_sap = ekpo-mwskz.
        MODIFY itemdata INDEX indice_i TRANSPORTING tax_code_sap.
      ENDIF.
      "Fim ODC 19_05_2020
      READ  TABLE item_aux WITH  KEY
                         po_number = itemdata-po_number
                         po_item = ekpo-ebelp
                         ref_doc = itemdata-ref_doc
                         ref_doc_year = itemdata-ref_doc_year
                         ref_doc_item = itemdata-ref_doc_item.
      IF sy-subrc NE 0.
        ADD 1 TO   indice .

        MOVE-CORRESPONDING itemdata TO item_aux.
        item_aux-po_item = ekpo-ebelp.
        item_aux-po_unit = ekpo-meins. "ODC - 04_03_2021
        CLEAR: item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid, item_aux-anln1, item_aux-anln2.
        SELECT SINGLE sakto kostl ps_psp_pnr aufnr  anln1 anln2  FROM ekkn INTO "#EC CI_NOORDER
       (item_aux-gl_account, item_aux-costcenter, item_aux-wbs_elem, item_aux-orderid,item_aux-anln1, item_aux-anln2)
       WHERE ebeln = itemdata-po_number
       AND ebelp = ekpo-ebelp.

        MOVE-CORRESPONDING headerdata TO item_aux.
        item_aux-item_text = itemdata-item_text.
        item_aux-invoice_doc_item = indice.
        item_aux-po_number = itemdata-po_number.
        item_aux-idnlf = ekpo-idnlf.
        APPEND item_aux.
        "ODC - 19_05_2020
      ELSE.
        item_aux-tax_code_sap = itemdata-tax_code_sap.
        MODIFY item_aux INDEX sy-tabix TRANSPORTING tax_code_sap.
        "Fim ODC - 19_05_2020
      ENDIF.
    ENDIF.
  ENDLOOP.
  DELETE item_aux WHERE po_number IS NOT INITIAL AND po_item IS INITIAL.
ENDFORM.                    " VALIDA_QTD_FATURA2
*&---------------------------------------------------------------------*
*&      Form  CALCULA_X
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_QTD_PED  text
*      -->P_VALOR_PED  text
*      -->P_ITEMDATA_ITEM_AMOUNT  text
*      -->P_QTD_FACTURADA  text
*      <--P_ITEMDATA_QUANT_FACT  text
*      <--P_QTD_REMANESC  text
*----------------------------------------------------------------------*
FORM calcula_x  USING    qtd_ped
                         valor_ped
                         itemdata_item_amount
                         qtd_facturada
                CHANGING itemdata_quantity
                         qtd_remanesc .

  itemdata_quantity = ( qtd_ped * itemdata_item_amount ) / valor_ped.

  qtd_remanesc = qtd_ped - qtd_facturada.

ENDFORM.                    " CALCULA_X
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_ESTRS_ADIANT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADERDATA  text
*----------------------------------------------------------------------*
FORM preenche_estrs_adiant  USING    headerdata LIKE /sbxc/zckp_adiantamento.

  REFRESH: lt_accountpayable, lt_currencyamount, lt_extension1.
*Se o campo invoice doc. number estiver preenchido e o campo ped.
*saphety
*não estiver preenchido então passa para o campo referencia o nº da
*fact.
  IF  headerdata-ref_1 IS INITIAL.
    headerdata-ref_doc_no = headerdata-ref_doc_no.
  ENDIF.

  MOVE-CORRESPONDING headerdata TO ls_documentheader.
  ls_documentheader-fisc_year = ls_documentheader-pstng_date(4).
* verifica na tabela ZCKP_MOV_FI as contas a utilizar no lançamento
  PERFORM valida_dados_fi USING headerdata-comp_code 'AD'.

  CLEAR indice.
  indice = indice + 1.

* preenche estrutura accountpayable
  PERFORM  preenche_est_lt_accountpayable  USING headerdata indice.

* Dados Itens da moeda do 1º item
  CLEAR sinal.

  PERFORM  preenche_est_lt_currencyamount USING indice '-' 'D'
  headerdata-amt_doccur  .

* preenche estrutura lt_EXTENSION1
  PERFORM  preenche_est_lt_extension1 USING 'CKP//F-47//K//A'  .

ENDFORM.                    " PREENCHE_ESTRS_ADIANT
*&---------------------------------------------------------------------*
*&      Form  VALIDA_DADOS_FI
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADERDATA_COMP_CODE  text
*      -->P_1197   text
*----------------------------------------------------------------------*
FORM valida_dados_fi  USING comp_code
                       tipo_op LIKE /sbxc/zckp_mfi-tipo_op.

  SELECT SINGLE * FROM /sbxc/zckp_mfi CLIENT SPECIFIED
         WHERE
            mandt = sy-mandt AND
            tipo_op = tipo_op AND
            bukrs = comp_code.

  IF sy-subrc NE 0.
    SELECT SINGLE * FROM /sbxc/zckp_mfi CLIENT SPECIFIED "#EC CI_NOORDER
       WHERE
          mandt = sy-mandt AND
          tipo_op = tipo_op.
  ENDIF.
* analisa moeda da empresa
  SELECT SINGLE waers FROM t001 INTO moeda
  WHERE
     bukrs = comp_code.

ENDFORM.                    " VALIDA_DADOS_FI
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_EST_LT_ACCOUNTPAYABLE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADERDATA  text
*      -->P_INDICE  text
*----------------------------------------------------------------------*
FORM preenche_est_lt_accountpayable  USING
      headerdata LIKE /sbxc/zckp_adiantamento
                 indice.

  lt_accountpayable-itemno_acc = indice.

  MOVE-CORRESPONDING headerdata TO  lt_accountpayable.
  lt_accountpayable-vendor_no = headerdata-vendor.
  IF headerdata-pstng_date IS INITIAL.
    headerdata-pstng_date =  headerdata-doc_date.
  ENDIF.

  lt_accountpayable-bline_date = headerdata-pstng_date.

  lt_accountpayable-sp_gl_ind = 'K'.


* passa para o campo texto a concatenação do nº pedido saphety
  CONCATENATE headerdata-ref_doc_no headerdata-ref_1 INTO
  lt_accountpayable-item_text SEPARATED BY space.
* passa nº ped. saphety e factura saphety para o campo xref3
  lt_accountpayable-ref_key_2 =  headerdata-ref_doc_no.
  lt_accountpayable-ref_key_3 =  headerdata-ref_1.

  APPEND lt_accountpayable.

ENDFORM.                    " PREENCHE_EST_LT_ACCOUNTPAYABLE
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_EST_LT_CURRENCYAMOUNT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_INDICE  text
*      -->P_1216   text
*      -->P_1217   text
*      -->P_HEADERDATA_AMT_DOCCUR  text
*----------------------------------------------------------------------*
FORM preenche_est_lt_currencyamount  USING    itemno_acc sinal d_c
valor.
* Verificar valor da entrada do item do pedido, será este o valor a
* lançar de diferimento
  lt_currencyamount-itemno_acc = itemno_acc.
  lt_currencyamount-currency   = moeda.
  lt_currencyamount-curr_type  = '00'.


  IF ( sinal = '+' AND d_c = 'D' ) OR ( sinal = '-' AND d_c = 'C' ).
    lt_currencyamount-amt_doccur = valor.
  ELSEIF ( sinal = '+' AND d_c = 'C' ) OR ( sinal = '-' AND d_c =
'D' ).
    lt_currencyamount-amt_doccur = -1 * valor.
  ENDIF.

  APPEND lt_currencyamount.
  CLEAR lt_currencyamount.

ENDFORM.                    " PREENCHE_EST_LT_CURRENCYAMOUNT
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_EST_LT_EXTENSION1
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_1222   text
*----------------------------------------------------------------------*
FORM preenche_est_lt_extension1  USING   field1.

*para que o módulo de função Z_SAMPLE_INTERFACE_RWBAPI01 apenas seja
*executado  quando se trata de uma referencia porveniente do
* SAPHETY colocamos o campo extension-field1 = 'CKP//F-47//K//A'.


*Transacção: FIBF
*Opções --> Modulos de processo --> cliente
*Inclusão do processo: RWBAPI01
*Criação da função Z_SAMPLE_INTERFACE_RWBAPI01
*Criação do prod. ZCKP (op --> produtos --> de um cliente)
*  lt_extension1-field1 = field1.
*  APPEND lt_extension1.

ENDFORM.                    " PREENCHE_EST_LT_EXTENSION1
*&---------------------------------------------------------------------*
*&      Form  LANCA_DOC_FI
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADERDATA_COMP_CODE  text
*      -->P_HEADERDATA_PSTNG_DATE  text
*      -->P_HEADERDATA_DOC_DATE  text
*      -->P_0015   text
*      -->P_HEADERDATA_USERNAME  text
*      -->P_HEADERDATA_PROCESSO  text
*      -->P_HEADERDATA_ANO  text
*      -->P_HEADERDATA_SEQNO  text
*----------------------------------------------------------------------*
FORM lanca_doc_fi  USING   comp_code pstng_date  doc_date texto
                        uname processo ano seqno.

  REFRESH t_return.

* Dados de Cabeçalho
  ls_documentheader-bus_act    = 'RFBU'.
  ls_documentheader-username   = uname.
  ls_documentheader-header_txt = texto.
  ls_documentheader-comp_code  = comp_code.
  ls_documentheader-doc_date   = doc_date.
  ls_documentheader-pstng_date = pstng_date.
  ls_documentheader-doc_type   = /sbxc/zckp_mfi-blart.



* Grava documento contábil
  CALL FUNCTION 'BAPI_ACC_DOCUMENT_POST' "#EC CI_USAGE_OK[2628704]
    EXPORTING
      documentheader = ls_documentheader
    IMPORTING
      obj_key        = l_awkey
    TABLES
      accountgl      = lt_accountgl "#EC CI_USAGE_OK[2438131]
      accountpayable = lt_accountpayable
      currencyamount = lt_currencyamount
      extension1     = lt_extension1
      return         = t_return.

  CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'.

  LOOP AT t_return.
    IF t_return-number = '605'.

      IF texto = TEXT-023. "'CKP:Diferimento'.
        UPDATE /sbxc/zckp_invh SET
*         doc_fi_diferimen = t_return-message_v2(10)
         doc_estorno = ' '
**         doc_estorno_dife = ' '
         ano_lanc = t_return-message_v2+14(4)
         enviado_saphety = ' '
   WHERE processo = processo AND
         ano = ano AND
         seqno = seqno.
      ELSEIF texto = TEXT-022. "'CKP: Adiantamento'.
        UPDATE /sbxc/zckp_adnt SET
              doc_fi = t_return-message_v2(10)
              doc_estorno = ' '
              ano_lanc = t_return-message_v2+14(4)
              enviado_saphety = ' '
        WHERE processo = processo AND
              ano = ano AND
              seqno = seqno.
      ENDIF.
      UPDATE /sbxc/zckp_ctrl SET
      status1 = '3'
      WHERE processo = processo AND
      ano = ano AND
      seqno = seqno.

    ENDIF.
    ADD 1 TO wa_msg-lineno.
    wa_msg-msgid = t_return-id.
    wa_msg-msgno = t_return-number.
    wa_msg-msgty = t_return-type.

    wa_msg-msgv1 =  t_return-message_v1.
    wa_msg-msgv2 =  t_return-message_v2.
    wa_msg-msgv3 =  t_return-message_v3.
    wa_msg-msgv4 =  t_return-message_v4.

    APPEND wa_msg TO msg_ckp .
  ENDLOOP.

  COMMIT WORK.

  REFRESH: lt_accountgl, lt_currencyamount.
  CLEAR: lt_accountgl, ls_documentheader.

ENDFORM.                    " LANCA_DOC_FI
*&---------------------------------------------------------------------*
*&      Form  LOG_DELETE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_RETURN  text
*      -->P_HEADERDATA_PROCESSO  text
*      -->P_HEADERDATA_ANO  text
*      -->P_HEADERDATA_SEQNO  text
*----------------------------------------------------------------------*
FORM log_delete  TABLES msg STRUCTURE bapiret2
          USING    processo
                   ano
                   seqno.

  DELETE FROM /sbxc/zckp_tab08 WHERE processo = processo AND ano = ano AND
         seqno = seqno.

ENDFORM.                    " LOG_DELETE
*&---------------------------------------------------------------------*
*&      Form  LOG
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_RETURN  text
*      -->P_HEADERDATA_PROCESSO  text
*      -->P_HEADERDATA_ANO  text
*      -->P_HEADERDATA_SEQNO  text
*----------------------------------------------------------------------*
FORM log  TABLES  msg STRUCTURE bapiret2
          USING    processo
                   ano
                   seqno.

  LOOP AT msg.
    ADD 1 TO /sbxc/zckp_tab08-num.
    /sbxc/zckp_tab08-processo = processo.
    /sbxc/zckp_tab08-ano = ano.
    /sbxc/zckp_tab08-seqno = seqno.
    /sbxc/zckp_tab08-msgid = msg-id.
    /sbxc/zckp_tab08-msgno = msg-number.
    /sbxc/zckp_tab08-msgty = msg-type.
    /sbxc/zckp_tab08-msgv1 = msg-message_v1.
    /sbxc/zckp_tab08-msgv2 = msg-message_v2.
    /sbxc/zckp_tab08-msgv3 = msg-message_v3.
    /sbxc/zckp_tab08-msgv4 = msg-message_v4.
    /sbxc/zckp_tab08-uname = sy-uname.
    /sbxc/zckp_tab08-data = sy-datum.
    /sbxc/zckp_tab08-hora = sy-uzeit.
    INSERT /sbxc/zckp_tab08.
  ENDLOOP.

ENDFORM.                    " LOG
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_ESTRUTURAS_INV
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_ITEMDATA  text
*      -->P_HEADERDATA  text
*----------------------------------------------------------------------*
FORM preenche_estruturas_inv   TABLES
                             itemdata STRUCTURE /sbxc/zckp_invi
                              USING headerdata LIKE /sbxc/zckp_invh.

*  RANGES: r_bukrs FOR t001-bukrs,
*          r_bsart FOR ekko-bsart.

  DATA: linha_fat           LIKE it_itemdata-invoice_doc_item,
        fat_estornada_mm(1).

  DATA: lt_itemdata     TYPE TABLE OF /sbxc/zckp_invi,
        lt_itemdata_aux TYPE TABLE OF /sbxc/zckp_invi,
        ls_itemdata     TYPE /sbxc/zckp_invi.

  MOVE-CORRESPONDING headerdata TO wa_header.
  wa_header-pmnttrms = headerdata-zterm.
  wa_header-diff_inv = headerdata-vendor.

* CCF 21.01.2022
* Se a condição de pagamento tem o campo “Dia fixo” preenchido (T052 - ZFAEL), o sistema faz EOMONTH da “data base”.
  DATA: zfael    LIKE t052-zfael,
        lv_ldate TYPE sy-datum.

  SELECT SINGLE zfael INTO zfael FROM t052 WHERE zterm = headerdata-zterm.
  IF zfael = '31'.
    CALL FUNCTION 'LAST_DAY_OF_MONTHS'
      EXPORTING
        day_in            = headerdata-baseline_date
      IMPORTING
        last_day_of_month = lv_ldate
*  EXCEPTIONS
*       DAY_IN_NO_DATE    = 1
*       OTHERS            = 2
      .
*    IF sy-subrc = 0.
*      ls_cab-baseline_date = lv_ldate.
*    ENDIF.

  ENDIF.
  wa_header-bline_date   =  lv_ldate .
*  wa_header-bline_date   = headerdata-baseline_date.
  "FIM CCF

*  Se é factura ou Nota de Crédito - se for 'X' e fact se for' ' é NC
* este campo deixou de fazer sentido, uma vez que o tipo de documento
* vem na mensagem
  IF headerdata-doctypesaphety = 'NC' OR headerdata-doctypesaphety  = 'CREDITNOTE'
  OR headerdata-doctypesaphety  = '4'
    OR headerdata-doctypesaphety = '381'. "ODC - 12_03_2021.
    wa_header-invoice_ind = ' '.
  ELSE.
    wa_header-invoice_ind = 'X'.
  ENDIF.

** se tipo de documento for de introdução manual, retirar esta linha
*  wa_header-doc_type = 'RE'.

** verifica existencia de factura dupla
*  DATA flag_bkpf(1).
*  TRANSLATE wa_header-ref_doc_no  TO UPPER CASE.
*  SELECT * FROM bsip WHERE
*      bukrs = wa_header-comp_code AND
*      lifnr = headerdata-vendor AND
*      waers = headerdata-currency AND
*      bldat = headerdata-doc_date AND
*      xblnr = wa_header-ref_doc_no AND
*      wrbtr = headerdata-gross_amount AND
*      gjahr = headerdata-ano_lanc AND
*     shkzg = 'H'.
*    CLEAR bkpf.
*    SELECT SINGLE * FROM bkpf WHERE bukrs = wa_header-comp_code AND
*                                    belnr = bsip-belnr AND
*                                    gjahr = bsip-gjahr AND
*                                    xreversal EQ ' '   AND
*                                    bstat NE 'V'.
*    IF sy-subrc = 0.
*      flag_bkpf = 'X'.
*    ENDIF.
*
*    SELECT SINGLE  * FROM rbkp WHERE
*              bukrs = bkpf-bukrs AND
*              belnr = bkpf-awkey(10) AND
*              gjahr = bkpf-awkey+10(4) AND
*              stblg = ''.
*    IF sy-subrc NE 0.
*      fat_estornada_mm = 'X'.
*    ENDIF.
*
*  ENDSELECT.
*
**  IF (  sy-subrc = 0 ) OR ( sy-subrc NE 0 AND flag_bkpf = 'X' ).
*  IF  fat_estornada_mm = ' ' AND flag_bkpf = 'X'.  "nao foi estornada nem em MM nem FI
*    MESSAGE s108(m8) WITH bsip-belnr bsip-gjahr DISPLAY LIKE 'E'.
*    RETURN.
**   Verificar se fatura já foi registrada sob documento contábil & &
*  ENDIF.


  IF  headerdata-pstng_date IS NOT INITIAL.
    wa_header-pstng_date = headerdata-pstng_date .
  ELSE.
    wa_header-pstng_date = sy-datum.
  ENDIF.
  IF wa_header-item_text IS INITIAL.
    CONCATENATE headerdata-processo headerdata-ano headerdata-seqno INTO
    wa_header-item_text SEPARATED BY '/'.
  ENDIF.

  wa_header-calc_tax_ind = 'X'.
* wa_header-pmnt_block = ' '.
*preenche estrutura de itens da factura ou itens de lançamentos directos
*a CR
  SORT itemdata BY po_number.
  CLEAR linha_fat.

  "Verificar se existem itens repetidos
  DATA: lv_pedido TYPE /sbxc/zckp_invi-po_number,
        lv_item   TYPE /sbxc/zckp_invi-po_item.
  DATA: lv_rep TYPE flag.
*  DATA: lv_cnt TYPE i.
  lt_itemdata_aux[] = itemdata[].
  LOOP AT itemdata WHERE po_number IS NOT INITIAL AND po_item IS NOT INITIAL .
    CLEAR lv_rep.

    LOOP AT lt_itemdata_aux INTO ls_itemdata WHERE po_number = itemdata-po_number AND po_item = itemdata-po_item AND tax_code_sap NE itemdata-tax_code_sap.
*      invoice_doc_item NE itemdata-invoice_doc_item.
      lv_rep = abap_true.
      APPEND ls_itemdata TO lt_itemdata.
    ENDLOOP.

    IF lv_rep EQ abap_true.
      CLEAR ls_itemdata.
      MOVE-CORRESPONDING itemdata TO ls_itemdata.
      APPEND ls_itemdata TO lt_itemdata.
      DELETE itemdata WHERE po_number = itemdata-po_number AND po_item = itemdata-po_item.
    ENDIF.

  ENDLOOP.
  SORT lt_itemdata BY po_number po_item invoice_doc_item.
  "Adicionar os itens repetidos nas tabelas
  LOOP AT lt_itemdata INTO ls_itemdata.

    IF ls_itemdata-po_number IS NOT INITIAL.

      IF ( lv_pedido NE ls_itemdata-po_number AND lv_item NE ls_itemdata-po_item AND lv_pedido IS NOT INITIAL AND lv_item IS NOT INITIAL )
      OR ( lv_pedido EQ ls_itemdata-po_number AND lv_item NE ls_itemdata-po_item AND lv_pedido IS NOT INITIAL AND lv_item IS NOT INITIAL ).
        APPEND it_itemdata.
        MOVE-CORRESPONDING ls_itemdata TO it_itemdata.
        it_itemdata-ref_doc_it = ls_itemdata-ref_doc_item.
        ADD 1 TO linha_fat.
        it_itemdata-invoice_doc_item = linha_fat.
        it_itemdata-de_cre_ind = ls_itemdata-dc_posterior.
        it_itemdata-po_pr_uom = ekpo-bprme.
        CLEAR: it_itemdata-item_amount,it_itemdata-quantity.
        CLEAR it_accountingdata.
        ADD 1 TO it_accountingdata-serial_no.
        CLEAR: it_accountingdata-xunpl.
      ELSEIF lv_pedido IS INITIAL AND lv_item IS INITIAL.
        MOVE-CORRESPONDING ls_itemdata TO it_itemdata.
        it_itemdata-ref_doc_it = ls_itemdata-ref_doc_item.
        ADD 1 TO linha_fat.
        it_itemdata-invoice_doc_item = linha_fat.
        it_itemdata-de_cre_ind = ls_itemdata-dc_posterior.
        it_itemdata-po_pr_uom = ekpo-bprme.
        CLEAR: it_itemdata-item_amount,it_itemdata-quantity.
        CLEAR it_accountingdata.
        ADD 1 TO it_accountingdata-serial_no.
        CLEAR: it_accountingdata-xunpl.
      ELSE.
        it_accountingdata-xunpl = abap_true.
        CLEAR it_accountingdata-serial_no.
      ENDIF.

      it_itemdata-item_amount = it_itemdata-item_amount + ls_itemdata-item_amount.
      it_itemdata-quantity    = it_itemdata-quantity + ls_itemdata-quantity.

      it_accountingdata-invoice_doc_item = linha_fat.
*      IF it_accountingdata-serial_no IS NOT INITIAL.
*        it_accountingdata-xunpl = abap_true.
*        CLEAR it_accountingdata-serial_no.
*      ELSE.
*        ADD 1 TO it_accountingdata-serial_no.
*        CLEAR: it_accountingdata-xunpl.
*      ENDIF.

      it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.
      it_accountingdata-tax_code = ls_itemdata-tax_code_sap.
      it_accountingdata-po_unit = ls_itemdata-po_unit.
      it_accountingdata-quantity = ls_itemdata-quantity .
      it_accountingdata-item_amount = ls_itemdata-item_amount.
      it_accountingdata-gl_account = ls_itemdata-gl_account.
      it_accountingdata-asset_no = ls_itemdata-anln1.
      it_accountingdata-sub_number = ls_itemdata-anln2.
      it_accountingdata-costcenter =  ls_itemdata-costcenter.
      it_accountingdata-wbs_elem =  ls_itemdata-wbs_elem.
      it_accountingdata-orderid = ls_itemdata-orderid.

      it_accountingdata-po_pr_uom = ls_itemdata-bprme.
      APPEND it_accountingdata.
    ENDIF.

    lv_pedido = ls_itemdata-po_number.
    lv_item   = ls_itemdata-po_item.
  ENDLOOP.

  IF lv_pedido EQ ls_itemdata-po_number AND lv_item EQ ls_itemdata-po_item AND lv_pedido IS NOT INITIAL AND lv_item IS NOT INITIAL.
    APPEND it_itemdata.
    CLEAR: lv_pedido, lv_item.
  ENDIF.

  CLEAR it_accountingdata.
  LOOP AT itemdata.

    SELECT SINGLE pstyp bprme INTO (ekpo-pstyp, ekpo-bprme) FROM ekpo
        WHERE ebeln = itemdata-po_number AND
              ebelp = itemdata-po_item.

    IF ekpo-pstyp = '1'. " item de limite
      MOVE-CORRESPONDING itemdata TO it_itemdata.
      ADD 1 TO linha_fat.
      it_itemdata-invoice_doc_item = linha_fat.
      IF itemdata-tax_code_sap IS NOT INITIAL.
        it_itemdata-tax_code = itemdata-tax_code_sap.
      ELSE.
        it_itemdata-tax_code = headerdata-mwskz.
      ENDIF.

      it_itemdata-de_cre_ind = itemdata-dc_posterior.
      it_itemdata-po_pr_uom = ekpo-bprme.
      CLEAR: it_itemdata-po_unit, it_itemdata-quantity.

      APPEND it_itemdata.
      MOVE-CORRESPONDING itemdata TO it_accountingdata.

      SELECT zekkn sakto kostl ps_psp_pnr aufnr wrbtr menge anln1 anln2 FROM /sbxc/zckp_ekkn INTO
      (it_accountingdata-serial_no, it_accountingdata-gl_account, it_accountingdata-costcenter, it_accountingdata-wbs_elem,
      it_accountingdata-orderid, it_accountingdata-item_amount, it_accountingdata-quantity, it_accountingdata-asset_no, it_accountingdata-sub_number)
      WHERE
      processo = headerdata-processo
      AND ano = headerdata-ano
      AND seqno = headerdata-seqno
      AND ebeln = itemdata-po_number
      AND ebelp = itemdata-po_item.

        it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.
        it_accountingdata-tax_code = itemdata-tax_code_sap.

        CLEAR: it_accountingdata-quantity, it_accountingdata-po_unit.

        it_accountingdata-po_pr_uom = itemdata-bprme.

        APPEND it_accountingdata.
        CLEAR it_accountingdata.
      ENDSELECT.
    ELSE." IF ekpo-pstyp = '1'
      IF itemdata-po_number IS NOT INITIAL.

        MOVE-CORRESPONDING itemdata TO it_itemdata.
        it_itemdata-ref_doc_it = itemdata-ref_doc_item.
        ADD 1 TO linha_fat.
        it_itemdata-invoice_doc_item = linha_fat.
*        it_itemdata-tax_code = itemdata-tax_code_sap.
        IF itemdata-tax_code_sap IS NOT INITIAL.
          it_itemdata-tax_code = itemdata-tax_code_sap.
        ELSE.
          it_itemdata-tax_code = headerdata-mwskz.
        ENDIF.
        it_itemdata-de_cre_ind = itemdata-dc_posterior.
        it_itemdata-po_pr_uom = ekpo-bprme.
        APPEND it_itemdata.

        IF  itemdata-gl_account IS NOT INITIAL AND  itemdata-clas_mult IS INITIAL
        OR ( ekkn-anln1 IS NOT INITIAL AND ekkn-anln2 IS NOT INITIAL AND itemdata-clas_mult IS INITIAL ) .
          SELECT SINGLE zekkn sakto ps_psp_pnr aufnr kostl anln1 anln2 FROM ekkn "#EC CI_NOORDER
        INTO (it_accountingdata-serial_no, ekkn-sakto, ekkn-ps_psp_pnr, ekkn-aufnr, ekkn-kostl, ekkn-anln1, ekkn-anln2)
        WHERE ebeln = itemdata-po_number
        AND ebelp = itemdata-po_item.

          IF ekkn-sakto IS INITIAL OR ( itemdata-gl_account <>  ekkn-sakto ) OR
       ( itemdata-wbs_elem <>  ekkn-ps_psp_pnr ) OR ( itemdata-orderid <> ekkn-aufnr )
       OR ( itemdata-costcenter <> ekkn-kostl )  OR ( itemdata-anln1 <> ekkn-anln1 )
       OR ( itemdata-anln2 <> ekkn-anln2 ).

            it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.

            it_accountingdata-tax_code = itemdata-tax_code_sap.
            it_accountingdata-po_unit = itemdata-po_unit.
            it_accountingdata-quantity = itemdata-quantity .
            it_accountingdata-item_amount = itemdata-item_amount.
            it_accountingdata-gl_account = itemdata-gl_account.
            it_accountingdata-asset_no = itemdata-anln1.
            it_accountingdata-sub_number = itemdata-anln2.
            it_accountingdata-costcenter =  itemdata-costcenter.
            it_accountingdata-wbs_elem =  itemdata-wbs_elem.
            it_accountingdata-orderid = itemdata-orderid.

            it_accountingdata-po_pr_uom = itemdata-bprme.
            APPEND it_accountingdata.
            CLEAR it_accountingdata.
*            elseIF ekkn-sakto IS NOT INITIAL and ( itemdata-gl_account =  ekkn-sakto ) and
*( itemdata-wbs_elem =  ekkn-ps_psp_pnr ) and ( itemdata-orderid = ekkn-aufnr )
*     and ( itemdata-costcenter = ekkn-kostl )  and ( itemdata-anln1 = ekkn-anln1 )
*     and ( itemdata-anln2 = ekkn-anln2 ).
*
*      it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.
*
*            it_accountingdata-tax_code = itemdata-tax_code_sap.
*            it_accountingdata-po_unit = itemdata-po_unit.
*            it_accountingdata-quantity = itemdata-quantity .
*            it_accountingdata-item_amount = itemdata-item_amount.
*            it_accountingdata-gl_account = itemdata-gl_account.
*            it_accountingdata-asset_no = itemdata-anln1.
*            it_accountingdata-sub_number = itemdata-anln2.
*            it_accountingdata-costcenter =  itemdata-costcenter.
*            it_accountingdata-wbs_elem =  itemdata-wbs_elem.
*            it_accountingdata-orderid = itemdata-orderid.
*
*            it_accountingdata-po_pr_uom = itemdata-bprme.
*            APPEND it_accountingdata.
*            CLEAR it_accountingdata.
          ENDIF.
        ELSEIF itemdata-clas_mult IS NOT INITIAL.
* Item com classificação contabil multipla


* Verifica se ja foram feitas alterações aos montantes e valores para imp. cont. mult
          SELECT zekkn sakto kostl ps_psp_pnr aufnr wrbtr menge anln1 anln2 FROM /sbxc/zckp_ekkn INTO
          (it_accountingdata-serial_no, it_accountingdata-gl_account, it_accountingdata-costcenter, it_accountingdata-wbs_elem,
          it_accountingdata-orderid, it_accountingdata-item_amount, it_accountingdata-quantity, it_accountingdata-asset_no, it_accountingdata-sub_number)
          WHERE
          processo = headerdata-processo
          AND ano = headerdata-ano
          AND seqno = headerdata-seqno
          AND ebeln = itemdata-po_number
          AND ebelp = itemdata-po_item.


            it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.

            it_accountingdata-tax_code = itemdata-tax_code_sap.

            it_accountingdata-po_unit = itemdata-po_unit.
            it_accountingdata-po_pr_uom = itemdata-bprme.
            APPEND it_accountingdata.
            CLEAR it_accountingdata.

          ENDSELECT.
          IF sy-subrc NE 0.
            SELECT zekkn sakto kostl ps_psp_pnr aufnr vproz anln1 anln2 FROM ekkn INTO
            (it_accountingdata-serial_no, it_accountingdata-gl_account, it_accountingdata-costcenter, it_accountingdata-wbs_elem,
            it_accountingdata-orderid, ekkn-vproz, it_accountingdata-asset_no, it_accountingdata-sub_number)
            WHERE ebeln = itemdata-po_number
            AND ebelp = itemdata-po_item.

              it_accountingdata-invoice_doc_item = it_itemdata-invoice_doc_item.

              it_accountingdata-tax_code = itemdata-tax_code_sap.

              it_accountingdata-item_amount = ( ekkn-vproz * itemdata-item_amount ) / 100.
              it_accountingdata-quantity =  ( ekkn-vproz * itemdata-quantity ) / 100.
              it_accountingdata-po_unit = itemdata-po_unit.
              it_accountingdata-po_pr_uom = itemdata-bprme.
              APPEND it_accountingdata.
              CLEAR it_accountingdata.
            ENDSELECT.
          ENDIF.
        ENDIF.

      ELSEIF itemdata-gl_account IS NOT INITIAL. " IF itemdata-po_number IS NOT INITIAL.
        MOVE-CORRESPONDING itemdata TO it_glaccountdata.
*        it_glaccountdata-db_cr_ind = itemdata-dc.
        it_glaccountdata-tax_code = itemdata-tax_code_sap.
        it_glaccountdata-comp_code = headerdata-comp_code.
        it_accountingdata-po_pr_uom = itemdata-bprme.
        CLEAR: it_glaccountdata-quantity.

        ADD 1 TO lin.
        it_glaccountdata-invoice_doc_item = lin.
        APPEND it_glaccountdata.

      ENDIF.
    ENDIF.

  ENDLOOP.
  DESCRIBE TABLE it_glaccountdata LINES lin.

  itemdata[] = lt_itemdata_aux[].

  SORT it_accountingdata BY invoice_doc_item xunpl.

ENDFORM.                    " PREENCHE_ESTRUTURAS_INV
**&---------------------------------------------------------------------*
**&      Form  VALIDA_QTD_FATURA
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->P_HEADERDATA  text
**      <--P_ITEMDATA  text
**----------------------------------------------------------------------*
*FORM valida_qtd_fatura  USING headerdata LIKE /sbxc/zckp_invh
*        CHANGING itemdata LIKE /sbxc/zckp_invi.
*
*
*
** para pedido item verifica quantidades facturadas até ao momento
*  REFRESH: lt_xekbes,  lt_xekbe.
*  CLEAR: lt_xekbes,  lt_xekbe.
*
*  CALL FUNCTION 'ME_READ_HISTORY'
*    EXPORTING
*      ebeln  = itemdata-po_number
*      ebelp  = itemdata-po_item
*      webre  = 'X'
*    TABLES
*      xekbe  = lt_xekbe
*      xekbes = lt_xekbes.
*
*  READ TABLE lt_xekbes INTO ls_xekbes WITH  KEY ebelp = itemdata-po_item
*                                       zekkn = '00'.
*
*  qtd_entrada  =  ls_xekbes-wemng.
*  val_entrada  = ls_xekbes-wewwr.
*  qtd_facturada = ls_xekbes-remng.
*  val_facturado = ls_xekbes-rewwr.
*
*  IF headerdata-doc_type = 'NC'.
*    itemdata-quantity = qtd_facturada.
*    dc_posterior = 'X'.
*  ELSE.
*    SELECT SINGLE menge brtwr INTO (qtd_ped, valor_ped) FROM ekpo
*           WHERE ebeln = itemdata-po_number AND
*                 ebelp = itemdata-po_item.
*
** 1ª Factura e factura final
*    IF  qtd_facturada = 0.
*      IF itemdata-fact_final = 'X'.
*        dc_posterior = ' '.
*        itemdata-quantity = qtd_ped.
*
*      ELSE. " Factura não final
*        dc_posterior = ' '.
*        PERFORM calcula_x USING qtd_ped valor_ped
*                         itemdata-item_amount qtd_facturada
*                         CHANGING itemdata-quantity qtd_remanesc.
*      ENDIF.
*
*    ELSEIF qtd_facturada < qtd_ped  . " qtd_facturada <> 0.
*      dc_posterior = ' '.
*      PERFORM calcula_x USING qtd_ped valor_ped
*                itemdata-item_amount qtd_facturada
*                CHANGING itemdata-quantity qtd_remanesc.
*
*      IF itemdata-fact_final = 'X'.
*        itemdata-quantity = qtd_remanesc.
*      ELSE.
*        IF itemdata-quantity > qtd_remanesc.
*          itemdata-quantity = qtd_remanesc.
*        ENDIF.
*      ENDIF.
*    ELSEIF qtd_facturada GE  qtd_ped.
*      dc_posterior = 'X'. " --> Debito posterior
*      itemdata-quantity = qtd_ped.
*    ENDIF.
*
*  ENDIF.
*
*ENDFORM.                    " VALIDA_QTD_FATURA
*&---------------------------------------------------------------------*
*&      Form  ACTUALIZA_BD
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_ITEMDATA  text
*      -->P_HEADERDATA  text
*      -->P_INVOICEDOCNUMBER  text
*      -->P_FISCALYEAR  text
*----------------------------------------------------------------------*
FORM actualiza_bd  TABLES itemdata STRUCTURE /sbxc/zckp_invi
                    USING headerdata TYPE /sbxc/zckp_invh
                          invoicedocnumber
                          fiscalyear
                 CHANGING t_ctrl TYPE /sbxc/zckp_ctrl.

  COMMIT WORK AND WAIT .

  CLEAR indice.
  REFRESH item2.
  SELECT SINGLE  rbstat  bukrs lifnr waers rmwwr FROM  rbkp
                  INTO (rbkp-rbstat,
headerdata-comp_code, headerdata-vendor, headerdata-currency, headerdata-gross_amount)
WHERE belnr = invoicedocnumber AND
gjahr = fiscalyear.

  IF sy-subrc = 0.
*  CHECK sy-subrc = 0.
    SELECT SINGLE    bldat budat xblnr  FROM  bkpf
INTO (headerdata-doc_date, headerdata-pstng_date, headerdata-ref_doc_no)
WHERE bukrs =  headerdata-comp_code AND
      belnr = headerdata-doc_fi AND
      gjahr = fiscalyear.


    IF rbkp-rbstat = '5'.
      UPDATE /sbxc/zckp_ctrl SET
             status1 = '3'
             data_chg_st1 = sy-datum
             hora_chg_st1 = sy-uzeit
             user_chg_st1 = sy-uname
             WHERE processo = headerdata-processo AND
             ano = fiscalyear AND
             seqno = headerdata-seqno.

      COMMIT WORK AND WAIT.
      t_ctrl-status1 = 'E'. "processado com exito


    ELSEIF rbkp-rbstat = 'A'.
      UPDATE /sbxc/zckp_ctrl SET
           status1 = '9'
           WHERE processo = headerdata-processo AND
           ano = fiscalyear AND
           seqno = headerdata-seqno.
      COMMIT WORK AND WAIT.
*     headerdata-status1 = '9'.
      t_ctrl-status1 = '9'.
      CLEAR lt_accdn. REFRESH lt_accdn.
      MOVE  fiscalyear TO ld_aworg.
      CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
        EXPORTING
          i_awtyp      = 'RMRP'
          i_awref      = invoicedocnumber
          i_aworg      = ld_aworg
        TABLES
          e_accdn      = lt_accdn
        EXCEPTIONS
          no_doc_found = 1
          OTHERS       = 2.


      IF sy-subrc = 0.
        LOOP AT lt_accdn INTO ls_accdn
          WHERE ldgrp IS INITIAL. "ODC - 04_11_2020
          headerdata-doc_fi = ls_accdn-belnr.
          headerdata-ano_lanc = ls_accdn-gjahr.
        ENDLOOP.
      ENDIF.

    ENDIF.

    UPDATE /sbxc/zckp_invh SET
              doc_date = headerdata-doc_date
              pstng_date = headerdata-pstng_date
              vendor = headerdata-vendor
              currency = headerdata-currency
              comp_code = headerdata-comp_code
              ref_doc_no = headerdata-ref_doc_no
              gross_amount = headerdata-gross_amount
              doc_fi = headerdata-doc_fi
              doc_lo =  headerdata-doc_lo
              doc_estorno = ' '
              ano_lanc = headerdata-ano_lanc
              data_criacao = sy-datum
*            enviado_saphety = ' '
       WHERE processo = headerdata-processo AND
             ano = headerdata-ano AND
             seqno = headerdata-seqno.

* Valida/actualiza lançamentos com ref a Pedido de compra

    DELETE FROM /sbxc/zckp_invi WHERE
            processo = headerdata-processo AND
            ano = headerdata-ano AND
            seqno = headerdata-seqno.
    COMMIT WORK AND WAIT.

    SELECT ebeln ebelp sgtxt wrbtr mwskz menge bstme tbtkz lfbnr lfgja lfpos
    INTO (wa_/sbxc/zckp_invi-po_number,wa_/sbxc/zckp_invi-po_item, wa_/sbxc/zckp_invi-item_text, wa_/sbxc/zckp_invi-item_amount,
    wa_/sbxc/zckp_invi-tax_code_sap, wa_/sbxc/zckp_invi-quantity, wa_/sbxc/zckp_invi-po_unit,
    wa_/sbxc/zckp_invi-dc_posterior, wa_/sbxc/zckp_invi-ref_doc, wa_/sbxc/zckp_invi-ref_doc_year, wa_/sbxc/zckp_invi-ref_doc_item)
    FROM rseg WHERE
    belnr = headerdata-doc_lo AND
    gjahr = headerdata-ano_lanc.

      ADD 1 TO indice.

      READ TABLE itemdata WITH KEY
          processo = headerdata-processo
          ano = headerdata-ano
          seqno = headerdata-seqno
          po_number = wa_/sbxc/zckp_invi-po_number
          po_item = wa_/sbxc/zckp_invi-po_item
          ref_doc = wa_/sbxc/zckp_invi-ref_doc
          ref_doc_year = wa_/sbxc/zckp_invi-ref_doc_year
          ref_doc_item = wa_/sbxc/zckp_invi-ref_doc_item.

* item ja existia
      IF sy-subrc = 0.
        MOVE-CORRESPONDING itemdata TO /sbxc/zckp_invi.
      ENDIF.
* item  novo
      MOVE-CORRESPONDING headerdata TO /sbxc/zckp_invi.
      MOVE-CORRESPONDING wa_/sbxc/zckp_invi TO /sbxc/zckp_invi.
      /sbxc/zckp_invi-costcenter = itemdata-costcenter.
      /sbxc/zckp_invi-gl_account = itemdata-gl_account.
      /sbxc/zckp_invi-anln1 = itemdata-anln1.
      /sbxc/zckp_invi-anln2 = itemdata-anln2.
      /sbxc/zckp_invi-tax_code_sap = itemdata-tax_code_sap.
      /sbxc/zckp_invi-db_cr_ind = itemdata-db_cr_ind.
      /sbxc/zckp_invi-invoice_doc_item = indice.
      /sbxc/zckp_invi-quant_fact = wa_/sbxc/zckp_invi-quantity.
      /sbxc/zckp_invi-item_amount = wa_/sbxc/zckp_invi-item_amount.

      REFRESH it_mwdat.
      CALL FUNCTION 'CALCULATE_TAX_FROM_NET_AMOUNT'
        EXPORTING
          i_bukrs           = headerdata-comp_code
          i_mwskz           = /sbxc/zckp_invi-tax_code_sap
          i_waers           = headerdata-currency
          i_wrbtr           = /sbxc/zckp_invi-item_amount
        TABLES
          t_mwdat           = it_mwdat
        EXCEPTIONS
          bukrs_not_found   = 1
          country_not_found = 2
          mwskz_not_defined = 3
          mwskz_not_valid   = 4
          ktosl_not_found   = 5
          kalsm_not_found   = 6
          parameter_error   = 7
          knumh_not_found   = 8
          kschl_not_found   = 9
          unknown_error     = 10
          account_not_found = 11
          txjcd_not_valid   = 12
          OTHERS            = 13.

      LOOP AT it_mwdat INTO wa_mwdat.
        /sbxc/zckp_invi-tax_imposto_sap = wa_mwdat-msatz.
        /sbxc/zckp_invi-tax_amount =  wa_mwdat-wmwst.
      ENDLOOP.

      INSERT /sbxc/zckp_invi.

      MOVE-CORRESPONDING /sbxc/zckp_invi TO item2.
      APPEND item2.
      CLEAR: /sbxc/zckp_invi, wa_/sbxc/zckp_invi.

    ENDSELECT.

* verifica lançamentos a CR
    SELECT wrbtr kostl saknr mwskz shkzg FROM rbco
INTO (wa_/sbxc/zckp_invi-item_amount, wa_/sbxc/zckp_invi-costcenter, wa_/sbxc/zckp_invi-gl_account,
wa_/sbxc/zckp_invi-tax_code_sap, wa_/sbxc/zckp_invi-db_cr_ind)
WHERE  belnr = headerdata-doc_lo AND
gjahr = headerdata-ano_lanc AND
buzei = '000000'.

      ADD 1 TO indice.
      MOVE-CORRESPONDING headerdata TO /sbxc/zckp_invi.
      MOVE-CORRESPONDING wa_/sbxc/zckp_invi TO /sbxc/zckp_invi.

      /sbxc/zckp_invi-invoice_doc_item = indice.

      INSERT /sbxc/zckp_invi.

      MOVE-CORRESPONDING /sbxc/zckp_invi TO item2.
      APPEND item2.
      CLEAR: /sbxc/zckp_invi, wa_/sbxc/zckp_invi.

    ENDSELECT.
* Verifica imputações contabilisticas mutiplas

    SELECT buzei ebeln ebelp FROM rseg INTO
       (rseg-buzei, rseg-ebeln, rseg-ebelp)
        WHERE  belnr = headerdata-doc_lo AND
              gjahr = headerdata-ano_lanc.

      DELETE FROM /sbxc/zckp_ekkn  WHERE
           ebeln = rseg-ebeln AND
           ebelp = rseg-ebelp AND
           processo = headerdata-processo AND
           ano = headerdata-ano AND
           seqno = headerdata-seqno.



      MOVE-CORRESPONDING rseg TO /sbxc/zckp_ekkn.
      MOVE-CORRESPONDING headerdata TO /sbxc/zckp_ekkn.
      SELECT  cobl_nr wrbtr menge kostl saknr ps_psp_pnr aufnr anln1 anln2
        FROM rbco
  INTO (/sbxc/zckp_ekkn-zekkn, /sbxc/zckp_ekkn-wrbtr, /sbxc/zckp_ekkn-menge, /sbxc/zckp_ekkn-kostl,
  /sbxc/zckp_ekkn-sakto, /sbxc/zckp_ekkn-ps_psp_pnr, /sbxc/zckp_ekkn-aufnr, /sbxc/zckp_ekkn-anln1, /sbxc/zckp_ekkn-anln2)
  WHERE  belnr = headerdata-doc_lo AND
  gjahr = headerdata-ano_lanc AND
  buzei = rseg-buzei.

        INSERT /sbxc/zckp_ekkn.


      ENDSELECT.
    ENDSELECT.
    CLEAR /sbxc/zckp_ekkn.
    COMMIT WORK AND WAIT.

    REFRESH itemdata.
    itemdata[] = item2[].
    COMMIT WORK AND WAIT.
  ENDIF.
ENDFORM.                    " ACTUALIZA_BD
*&---------------------------------------------------------------------*
*&      Form  BI_MIR4
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_INVOICEDOCNUMBER  text
*      -->P_FISCALYEAR  text
*----------------------------------------------------------------------*
FORM bi_mir4  USING    invoicedocnumber fiscalyear.


  CLEAR messtab. REFRESH messtab.

  CLEAR opt.
  opt-dismode  = 'E'.
  opt-updmode  = 'S'.
  opt-racommit = 'X'.
  REFRESH messtab.

  PERFORM dynpro USING:  'X' 'SAPLMR1M'    '6150',
                           ' ' 'BDC_OKCODE'  '/00',
                           ' ' 'RBKP-BELNR'  invoicedocnumber,
                           ' ' 'RBKP-GJAHR' fiscalyear,
                           'X' 'SAPLMR1M'    '6000',
                           ' ' 'BDC_OKCODE'  '/EPPCH'.

  CALL TRANSACTION 'MIR4' USING bdcdata
                          OPTIONS FROM opt
                          MESSAGES INTO messtab.

ENDFORM.                                                    " BI_MIR4
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_ESTRUT_FUNC
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LISTA_FACTURAS  text
*----------------------------------------------------------------------*
FORM preenche_estrut_func  USING   lista_facturas LIKE /sbxc/zckp_bloq_fatura.

  DATA doc_log LIKE bkpf-awkey.
* procura documento financeiro associado ao doc. logistica

  MOVE lista_facturas-fisc_year TO ld_aworg.
  CLEAR lt_accdn. REFRESH lt_accdn.
  CALL FUNCTION 'FI_DOCUMENT_FIND_FOR_INTERFACE'
    EXPORTING
      i_awtyp      = 'RMRP'
      i_awref      = lista_facturas-inv_doc_no
      i_aworg      = ld_aworg
    TABLES
      e_accdn      = lt_accdn
    EXCEPTIONS
      no_doc_found = 1
      OTHERS       = 2.
* se não encontra doc de fi então é porque o documento enviado já é o
* documento financeiro
  IF sy-subrc = 0.
    LOOP AT lt_accdn INTO ls_accdn
      WHERE ldgrp IS INITIAL. "ODC - 04_11_2020
      lista_facturas-inv_doc_no = ls_accdn-awref.
    ENDLOOP.
  ELSE.
    SELECT SINGLE awkey FROM bkpf INTO doc_log WHERE
         bukrs = lista_facturas-comp_code AND
         belnr = lista_facturas-inv_doc_no  AND
         gjahr = lista_facturas-fisc_year.
    lista_facturas-inv_doc_no = doc_log(10).
  ENDIF.


  ls_accchg-fdname = 'ZLSPR'.
  ls_accchg-newval = lista_facturas-mot_bloq.

  APPEND ls_accchg TO lt_accchg.

ENDFORM.                    " PREENCHE_ESTRUT_FUNC
*&---------------------------------------------------------------------*
*&      Form  VERIFICA_NIF_EMP
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_VENDORCOMPANY_COMP_CODE  text
*      <--P_VENDORCOMPANY_NIF_EMPRESA_COMP  text
*----------------------------------------------------------------------*
FORM verifica_nif_emp   USING    p_comp_code
                       CHANGING p_nif_empresa_comp.

  SELECT SINGLE stceg INTO p_nif_empresa_comp FROM t001
        CLIENT SPECIFIED
         WHERE  mandt = sy-mandt   AND
         bukrs = p_comp_code.

ENDFORM.                    " VERIFICA_NIF_EMP
*&---------------------------------------------------------------------*
*&      Form  DERIVA_CAMPOS_PO_CHANGE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_POCHANGE  text
*----------------------------------------------------------------------*
FORM deriva_campos_po_change  USING     pochange LIKE /sbxc/zbapipochange.
* Verifica nº de pedido e item utilizando nº pedido e item saphety

  UNPACK pochange-vendor TO pochange-vendor.
  SELECT SINGLE ebeln FROM ekko INTO pedido  WHERE      "#EC CI_NOORDER
         lifnr = pochange-vendor AND
         bukrs = pochange-comp_code AND
         ihrez = pochange-ref_1.

ENDFORM.                    " DERIVA_CAMPOS_PO_CHANGE
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_ESTRUTURAS_PO_CHANGE
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_POCHANGE  text
*----------------------------------------------------------------------*
FORM preenche_estruturas_po_change  USING    pochange LIKE /sbxc/zbapipochange.

* Se o campo PO_ITEM vier preenchido é apenas para marcar para eliminar
* este item, caso contrario é para marcar para eliminar todos os itens

  IF pochange-po_item IS NOT INITIAL.
    pochange-po_item = pochange-po_item * 10.
    ls_poitem-po_item = pochange-po_item.
    ls_poitem-delete_ind = 'L'.
    APPEND ls_poitem TO lt_poitem.

    ls_poitemx-po_item = pochange-po_item.
    ls_poitemx-delete_ind = 'X'.
    APPEND ls_poitemx TO lt_poitemx.

  ELSE.
    SELECT ebelp FROM ekpo INTO ls_poitem-po_item WHERE
           ebeln = pedido.

      ls_poitem-delete_ind = 'L'.
      APPEND ls_poitem TO lt_poitem.

      ls_poitemx-po_item = ls_poitem-po_item.
      ls_poitemx-delete_ind = 'X'.
      APPEND ls_poitemx TO lt_poitemx.

    ENDSELECT.
  ENDIF.

ENDFORM.                    " PREENCHE_ESTRUTURAS_PO_CHANGE
*&---------------------------------------------------------------------*
*&      Form  DECODE_BASE64
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_BASE64_STRING  text
*      <--P_XML_OUT  text
*      <--P_L_XSTR  text
*----------------------------------------------------------------------*
FORM decode_base64   USING
      base_64     TYPE string
    CHANGING
      plaintext  TYPE string
      l_xstr     TYPE xstring.

  CHECK base_64  IS NOT INITIAL.

  CONSTANTS:
    lc_op_dec TYPE x VALUE 37.
  DATA: lr_conv TYPE REF TO cl_abap_conv_in_ce.

  CALL 'SSF_ABAP_SERVICE'
    ID 'OPCODE' FIELD   lc_op_dec
    ID 'BINDATA' FIELD   l_xstr
    ID 'B64DATA' FIELD   base_64.                         "#EC CI_CCALL

  TRY.

      lr_conv = cl_abap_conv_in_ce=>create( input = l_xstr ).
      lr_conv->read( IMPORTING data = plaintext ).

    CATCH cx_sy_conversion_codepage.
      CLEAR plaintext.
      MESSAGE i999(samx) WITH TEXT-004 TEXT-005.
  ENDTRY.

ENDFORM.                    " DECODE_BASE64
*&---------------------------------------------------------------------*
*&      Form  ERROR_PROCESSING
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_RETURN  text
*      -->P_0136   text
*      -->P_0137   text
*      -->P_0138   text
*      -->P_W_RETURN_MESSAGE_V1  text
*      -->P_0140   text
*      -->P_0141   text
*      -->P_0142   text
*----------------------------------------------------------------------*
FORM error_processing  TABLES   p_return STRUCTURE bapiret2
                                  "Introduzir nome correto para <...>
                       USING    tipo LIKE sy-msgty
                                classe LIKE sy-msgid
                                numero LIKE sy-msgno
                                var1 LIKE sy-msgv1
                                var2 LIKE sy-msgv2
                                var3 LIKE sy-msgv3
                                var4 LIKE sy-msgv4.



  CALL FUNCTION 'BALW_BAPIRETURN_GET2'
    EXPORTING
      type   = tipo
      cl     = classe
      number = numero
      par1   = var1
      par2   = var2
      par3   = var3
      par4   = var4
*     parameter = imp_parameter
*     row    = imp_row
*     field  = imp_field
    IMPORTING
      return = w_return.

  APPEND w_return TO p_return.

ENDFORM.                    " ERROR_PROCESSING
*&---------------------------------------------------------------------*
*&      Form  DERIVA_CAMPOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_POITEM  text
*      -->P_POSCHEDULE  text
*      -->P_POACCOUNT  text
*      -->P_POHEADER  text
*----------------------------------------------------------------------*
FORM deriva_campos TABLES poitem STRUCTURE /sbxc/zbapimepoitem
                   poschedule  STRUCTURE /sbxc/zbapimeposchedul
                   poaccount STRUCTURE /sbxc/zbapimepoaccount
             USING poheader TYPE /sbxc/zbapimepoheader.

  UNPACK poheader-vendor TO poheader-vendor.

* Grupo de compradores; Categoria Class. Contabil

* ler tabela /SBXC/zckp_mped

  SELECT SINGLE pur_group acctasscat asset_no stge_loc FROM
  /sbxc/zckp_mped INTO
         (poheader-pur_group, /sbxc/zckp_mped-acctasscat,
         /sbxc/zckp_mped-asset_no, /sbxc/zckp_mped-stge_loc)
         WHERE doc_type = poheader-doc_type AND bukrs =
         poheader-comp_code.

  IF sy-subrc NE 0.
    SELECT SINGLE pur_group acctasscat asset_no stge_loc FROM "#EC CI_NOORDER
    /sbxc/zckp_mped INTO
           (poheader-pur_group, /sbxc/zckp_mped-acctasscat,
           /sbxc/zckp_mped-asset_no, /sbxc/zckp_mped-stge_loc)
           WHERE doc_type = poheader-doc_type.
  ENDIF.

ENDFORM.                    " DERIVA_CAMPOS
*&---------------------------------------------------------------------*
*&      Form  PREENCHE_ESTRUTURAS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_POITEM  text
*      -->P_POSCHEDULE  text
*      -->P_POACCOUNT  text
*      -->P_POHEADER  text
*----------------------------------------------------------------------*
FORM preenche_estruturas  TABLES poitem STRUCTURE /sbxc/zbapimepoitem
                   poschedule  STRUCTURE /sbxc/zbapimeposchedul
                   poaccount STRUCTURE /sbxc/zbapimepoaccount
             USING poheader LIKE /sbxc/zbapimepoheader.


  DATA lin TYPE i.

* verifica campos da estrutura bapi
  SELECT tabname fieldname FROM dd03l INTO CORRESPONDING FIELDS OF
  TABLE t_field
          WHERE ( tabname = '/SBXC/ZBAPIMEPOHEADER' OR  tabname =
          '/SBXC/ZBAPIMEPOITEM' OR
tabname  = '/SBXC/ZBAPIMEPOSCHEDUL' OR tabname  = '/SBXC/ZBAPIMEPOACCOUNT'
).

* Preenche estrutura X


  LOOP AT t_field  WHERE tabname = '/SBXC/ZBAPIMEPOHEADER'.
    CONCATENATE 'POHEADER-' t_field-fieldname INTO l_fieldname.
    ASSIGN (l_fieldname) TO <field>.
    IF <field> IS NOT INITIAL.
      CONCATENATE 'POHEADERX-' t_field-fieldname INTO l_fieldname.
      ASSIGN (l_fieldname) TO <field>.
      <field> = 'X'.
    ENDIF.
  ENDLOOP.

  LOOP AT poschedule.
    poschedule-po_item = poschedule-po_item * 10.
    LOOP AT t_field  WHERE tabname = '/SBXC/ZBAPIMEPOSCHEDUL'.
      CONCATENATE 'POSCHEDULE-' t_field-fieldname INTO l_fieldname.
      ASSIGN (l_fieldname) TO <field>.
      IF <field> IS NOT INITIAL.
        CONCATENATE 'POSCHEDULEX-' t_field-fieldname INTO l_fieldname.
        ASSIGN (l_fieldname) TO <field>.
        IF t_field-fieldname =   'PO_ITEM'.
          poschedulex-po_item  = poschedule-po_item.
        ELSEIF t_field-fieldname =   'SCHED_LINE'.
          poschedulex-sched_line =  poschedule-sched_line.
        ELSE.
          <field> = 'X'.
        ENDIF.
      ENDIF.
    ENDLOOP.
*    poschedulex-del_datcat_ext = 'X'.
    APPEND poschedulex.
*    bapi_schedule-del_datcat_ext = 'D'.
    MOVE-CORRESPONDING poschedule TO bapi_schedule.
    APPEND bapi_schedule.

  ENDLOOP.

  LOOP AT poaccount.
    poaccount-po_item = poaccount-po_item * 10.
    MODIFY poaccount TRANSPORTING po_item.

    LOOP AT t_field  WHERE tabname = '/SBXC/ZBAPIMEPOACCOUNT'.
      CONCATENATE 'POACCOUNT-' t_field-fieldname INTO l_fieldname.
      ASSIGN (l_fieldname) TO <field>.
      IF <field> IS NOT INITIAL.
        IF t_field-fieldname =  'DISTR_PERC' AND <field> = '100.0'.
          CLEAR poaccount-distr_perc.
          MODIFY poaccount TRANSPORTING distr_perc.
        ENDIF.

*      else.
        CONCATENATE 'POACCOUNTX-' t_field-fieldname INTO l_fieldname.
        ASSIGN (l_fieldname) TO <field>.
        IF t_field-fieldname =   'PO_ITEM'.
          poaccountx-po_item = poaccount-po_item.
        ELSEIF t_field-fieldname =  'SERIAL_NO'.
          poaccountx-serial_no = poaccount-serial_no.
        ELSE.
          <field> = 'X'.

        ENDIF.
      ENDIF.
    ENDLOOP.


    IF /sbxc/zckp_mped-asset_no IS NOT INITIAL.
      poaccountx-asset_no = 'X'.
    ENDIF.

    APPEND  poaccountx.
    MOVE-CORRESPONDING poaccount TO bapi_account.

    IF poaccount-wbs_element IS NOT INITIAL.
      CALL FUNCTION 'CONVERSION_EXIT_ABPSN_OUTPUT'
        EXPORTING
          input  = poaccount-wbs_element
        IMPORTING
          output = bapi_account-wbs_element.
    ENDIF.

    bapi_account-asset_no = /sbxc/zckp_mped-asset_no.

    APPEND bapi_account.

  ENDLOOP.

  LOOP AT poitem.
    poitem-po_item = poitem-po_item * 10.
    LOOP AT t_field  WHERE tabname = '/SBXC/ZBAPIMEPOITEM'.
      CONCATENATE 'POITEM-' t_field-fieldname INTO l_fieldname.
      ASSIGN (l_fieldname) TO <field>.

      IF <field> IS NOT INITIAL.
        CONCATENATE 'POITEMX-' t_field-fieldname INTO l_fieldname.
        ASSIGN (l_fieldname) TO <field>.
        IF t_field-fieldname =   'PO_ITEM'.
          poitemx-po_item = poitem-po_item.
          poitemx-po_itemx = 'X'.
        ELSE.
          <field> = 'X'.
        ENDIF.
      ENDIF.



    ENDLOOP.
    IF /sbxc/zckp_mped-acctasscat IS NOT INITIAL.
      poitemx-acctasscat = 'X'.
    ENDIF.

*POITEM-PO_PRICE
* Note 580225 - Purchasing BAPIs: Conditions and pricing
* You can use the PO_PRICE field to control, at item level,
* if the value should be copied from the POITEM-NET_PRICE
* field to the conditions. PO_PRICE can have the values ' ', '1' or '2' with the following meaning:
*
*PO_PRICE = ' ': The conditions are determined automatically, the value from the NET_PRICE
*field is only copied if the system cannot determine a condition.
*PO_PRICE = '1': The value transferred in field NET_PRICE is copied as a gross price
*that is, it is set with the condition type specified as base price in the calculation schema.
*In the SAP Standard System, these are condition types PB00 or PBXX.
*All other condition types remain unchanged. No conditions are copied from the last document.
*PO_PRICE = '2': The value transferred in field NET_PRICE is copied as a net price that is,
*it is set with the condition type specified as base price in the calculation procedure.
*All other condition types are deleted.


    bapi_item-po_price = '2'.
    poitemx-po_price = 'X'.

    MOVE-CORRESPONDING poitem TO bapi_item.

    bapi_item-stge_loc = /sbxc/zckp_mped-stge_loc.

    poitemx-stge_loc = 'X'.

*Se pisco de não espera entrada etiver vazio significa que vai haver
*entrada de material e como
*tal deverá passar para BAPI com 'X' e o campo EM n/avaliada  tem de
*estar ' '
    IF bapi_item-gr_ind IS INITIAL.
      bapi_item-gr_non_val = ' '.
      poitemx-gr_non_val = 'X'.
      bapi_item-gr_ind = 'X'.
      poitemx-gr_ind = 'X'.
    ELSE.
      bapi_item-gr_non_val = 'X'.
      poitemx-gr_non_val = 'X'.
      bapi_item-gr_ind = 'X'.
      poitemx-gr_ind = 'X'.
    ENDIF.


    poitemx-plant = 'X'.

    bapi_item-acctasscat = /sbxc/zckp_mped-acctasscat.
* Verifica se o item tem imputação contabilistica múltipla
    READ TABLE poaccount WITH KEY po_item =  bapi_item-po_item
                                  serial_no = '02'.

    IF sy-subrc = 0.
      bapi_item-distrib = '2'.
      bapi_item-part_inv  = '2'.
      poitemx-distrib = 'X'.
      poitemx-part_inv  = 'X'.

    ENDIF.
    APPEND bapi_item.
    APPEND poitemx.


  ENDLOOP.

  MOVE-CORRESPONDING  poheader TO bapi_header.

ENDFORM.                    " PREENCHE_ESTRUTURAS
**&---------------------------------------------------------------------*
**&      Form  PREENCHE_ESTRUTURAS_GM
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->P_GOODSMVT_ITEM  text
**      -->P_GOODSMVT_HEADER  text
**----------------------------------------------------------------------*
*FORM preenche_estruturas_gm  TABLES
*                        itab_item STRUCTURE /sbxc/zbapi_gm_item
*                        USING    itab_header LIKE /sbxc/zbapi_gm_head.
*
*
** Preenche estrutura do BAPI
** Verifica nº de pedido e item utilizando nº pedido e item saphety
*
*  UNPACK itab_header-vendor TO itab_header-vendor.
*  SELECT SINGLE ebeln FROM ekko INTO pedido  WHERE
**         lifnr = itab_header-vendor AND
*         bukrs = itab_header-comp_code AND
*         ihrez = itab_header-ref_doc_no.
*
*  MOVE-CORRESPONDING itab_header TO gm_header.
*
*
*  LOOP AT itab_item.
*    indice = sy-tabix.
*
*    CLEAR w_gm_item.
*    MOVE-CORRESPONDING itab_item TO w_gm_item.
*    MOVE-CORRESPONDING itab_item TO gm_header.
*
*    w_gm_item-po_number = pedido.
*
** Verifica nº de item pedido e item utilizando nº pedido e item saphety
*    SELECT SINGLE ebelp meins FROM ekpo INTO
*          (w_gm_item-po_item,  w_gm_item-entry_uom)
*    WHERE
*           ebeln = w_gm_item-po_number AND
*           bednr = itab_item-trackingno.
*
*    IF itab_item-sinal = '+' OR itab_item-sinal IS INITIAL .
*      w_gm_item-move_type = '101'.
*    ELSE.
*      w_gm_item-move_type = '102'.
*    ENDIF.
*    w_gm_item-mvt_ind = 'B'.
*    APPEND w_gm_item TO gm_items.
*
*
*    itab_item-po_number = w_gm_item-po_number.
*    itab_item-po_item = w_gm_item-po_item.
*    MODIFY itab_item INDEX indice TRANSPORTING po_number po_item .
*  ENDLOOP.
*
*ENDFORM.                    " PREENCHE_ESTRUTURAS_GM
*&---------------------------------------------------------------------*
*&      Form  ENVIA_FATURAS_CRIADAS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_DOC  text
*      -->P_DATA  text
*----------------------------------------------------------------------*
FORM envia_faturas_criadas  TABLES doc
            STRUCTURE /sbxc/zckp_status_oc_fat_pag
            USING data .

  SELECT  processo ano seqno doc_type comp_code ref_doc_no vendor doc_fi ano_lanc
  FROM /sbxc/zckp_invh CLIENT SPECIFIED
  INTO (/sbxc/zckp_invh-processo, /sbxc/zckp_invh-ano, /sbxc/zckp_invh-seqno,
  doc-doc_type, doc-comp_code, doc-ref_doc_no, doc-vendor, doc-doc_fi, doc-ano_lanc )
  WHERE  mandt = sy-mandt AND
  data_criacao = data AND
  enviado_saphety = ' '.

    SELECT SINGLE status1 FROM /sbxc/zckp_ctrl          "#EC CI_NOORDER
      CLIENT SPECIFIED
      INTO /sbxc/zckp_ctrl-status1
      WHERE processo = /sbxc/zckp_invh-processo AND
      ano = /sbxc/zckp_invh-ano AND
      seqno = /sbxc/zckp_invh-seqno.

    IF /sbxc/zckp_ctrl-status1 = '3' OR /sbxc/zckp_ctrl-status1 = '4'.
* verifica data da factura
      SELECT SINGLE bldat FROM bkpf INTO doc-data_pagamento WHERE
            bukrs = doc-comp_code AND
            belnr = doc-doc_fi AND
            gjahr = doc-ano_lanc.

      doc-tipo_actualiza = '2'.

      APPEND doc.
      CLEAR doc.
    ENDIF.

  ENDSELECT.

ENDFORM.                    " ENVIA_FATURAS_CRIADAS
*&---------------------------------------------------------------------*
*&      Form  ENVIA_PAGAMENTOS_CRIADAS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_DOC  text
*      -->P_DATA  text
*----------------------------------------------------------------------*
FORM envia_pagamentos_criadas  TABLES   doc
            STRUCTURE /sbxc/zckp_status_oc_fat_pag
            USING data .

  SELECT processo ano seqno doc_type  comp_code ref_doc_no vendor doc_fi ano_lanc  FROM /sbxc/zckp_invh
  CLIENT SPECIFIED
  INTO (/sbxc/zckp_invh-processo, /sbxc/zckp_invh-ano, /sbxc/zckp_invh-seqno,
  doc-doc_type, doc-comp_code, doc-ref_doc_no, doc-vendor, doc-doc_fi, doc-ano_lanc )
  WHERE  mandt = sy-mandt AND
  env_paga_saphety NE 'X'.

    CHECK doc-doc_fi <> ' '.

    SELECT SINGLE status1 FROM /sbxc/zckp_ctrl          "#EC CI_NOORDER
            CLIENT SPECIFIED
            INTO /sbxc/zckp_ctrl-status1
            WHERE processo = /sbxc/zckp_invh-processo AND
                  ano = /sbxc/zckp_invh-ano AND
                  seqno = /sbxc/zckp_invh-seqno.

    IF /sbxc/zckp_ctrl-status1 = '3' OR /sbxc/zckp_ctrl-status1 = '4'.
* Verifica se documento de faCLIENT SPECIFIEDctura já tem documento de
*compensação
      SELECT SINGLE augdt augbl FROM bsak  CLIENT SPECIFIED "#EC CI_NOORDER
        INTO (doc-data_pagamento, doc-doc_pagamento)
        WHERE mandt = sy-mandt AND
              bukrs = doc-comp_code AND
              gjahr = doc-ano_lanc AND
              belnr = doc-doc_fi.



      CHECK sy-subrc = 0 AND doc-doc_pagamento IS NOT INITIAL.
      doc-tipo_actualiza = '3'.

      APPEND doc.
      CLEAR doc.

    ENDIF.

  ENDSELECT.

ENDFORM.                    " ENVIA_PAGAMENTOS_CRIADAS
*&---------------------------------------------------------------------*
*&      Form  ACT_TAB_COCKPIT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_DOC  text
*----------------------------------------------------------------------*
FORM act_tab_cockpit  TABLES   doc
            STRUCTURE /sbxc/zckp_status_oc_fat_pag.


  LOOP AT doc.
* verifica NIF empresa compradora
    IF doc-nif_empresa_comp IS INITIAL.
      PERFORM verifica_nif_emp USING doc-comp_code
               CHANGING doc-nif_empresa_comp.
      MODIFY doc TRANSPORTING nif_empresa_comp
      WHERE comp_code = doc-comp_code.
    ENDIF.
* verifica NIF fornecedor
    PERFORM verifica_nif_forn USING doc-vendor
                CHANGING doc-nif_fornecedor.
    MODIFY doc TRANSPORTING nif_fornecedor
     WHERE vendor = doc-vendor.

    IF doc-tipo_actualiza = '2'.

    ELSEIF   doc-tipo_actualiza = '3'.

    ENDIF.
  ENDLOOP.

ENDFORM.                    " ACT_TAB_COCKPIT
*&---------------------------------------------------------------------*
*&      Form  VERIFICA_NIF_FORN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_DOC_VENDOR  text
*      <--P_DOC_NIF_FORNECEDOR  text
*----------------------------------------------------------------------*
FORM verifica_nif_forn   USING p_vendor
                        CHANGING p_nif_fornecedor.


  SELECT  SINGLE stceg  stcd1 stcd2  stenr
            FROM lfa1  CLIENT SPECIFIED
                 INTO (p_nif_fornecedor,
                         stcd1, stcd2,  stenr)
                  WHERE mandt = sy-mandt  AND
                       lifnr = p_vendor.

*se campo stceg vazio preenche com stcd1 se este estiver vazio preenche
*stcd2 se vazio preenche stenr
  IF p_nif_fornecedor IS INITIAL AND stcd1 IS NOT INITIAL.
    p_nif_fornecedor = stcd1.
  ELSEIF p_nif_fornecedor IS INITIAL AND stcd2 IS NOT INITIAL.
    p_nif_fornecedor = stcd2.
  ELSEIF p_nif_fornecedor IS INITIAL AND stenr IS NOT INITIAL.
    p_nif_fornecedor = stenr.
  ENDIF.

ENDFORM.                    " VERIFICA_NIF_FORN
*&---------------------------------------------------------------------*
*&      Form  SET_COLOR
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_PROCESSO  text
*      -->P_/SBXC/ZCKP_CTRL_STATUS1  text
*----------------------------------------------------------------------*
FORM set_color
                TABLES   cor_tab STRUCTURE /sbxc/zckp_color
                USING    processo
                         status
                         est_cab.

*  SELECT SINGLE * FROM /sbxc/zckp_tab00
*    WHERE processo = processo.

  READ TABLE gt_tab10 WITH KEY processo = processo
                              status = status.

  IF sy-subrc EQ 0 AND NOT gt_tab10-cor IS INITIAL.
*   Campos da tabela de cabeçalho
    LOOP AT gt_dd03l
      WHERE tabname = est_cab "/sbxc/zckp_tab00-est_cab
      AND comptype NE 'S'.
      cor_tab-tabix = 1.
      cor_tab-int = 1.
      cor_tab-col = gt_tab10-cor.
      cor_tab-fname = gt_dd03l-fieldname.
      APPEND cor_tab.
    ENDLOOP.
*   Campos da tabela de controlo
    LOOP AT gt_dd03l
    WHERE tabname = '/SBXC/ZCKP_CTRL'
    AND comptype NE 'S'.
      cor_tab-tabix = 1.
      cor_tab-int = 1.
      cor_tab-col = gt_tab10-cor.
      cor_tab-fname = gt_dd03l-fieldname.
      APPEND cor_tab.
    ENDLOOP.

    cor_tab-tabix = 1.
    cor_tab-int = 1.
    cor_tab-col = gt_tab10-cor.
    cor_tab-fname = 'STATUS_OUT'.
    APPEND cor_tab.
  ENDIF.
ENDFORM.                    " SET_COLOR
*&---------------------------------------------------------------------*
*&      Form  set_status
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->RT_EXTAB   text
*----------------------------------------------------------------------*
FORM status USING rt_extab TYPE slis_t_extab.
  SET PF-STATUS 'STATUS'.
ENDFORM.                    "set_status
*&---------------------------------------------------------------------*
*&      Form  set_status2
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->RT_EXTAB   text
*----------------------------------------------------------------------*
FORM set_status2 USING rt_extab TYPE slis_t_extab.
  SET PF-STATUS '001'.
ENDFORM.                    "set_status
*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
FORM user_command                                           "#EC CALLED
  USING  l_ucomm        LIKE sy-ucomm
         ls_selfield    TYPE slis_selfield.

  CLEAR: ls_selfield-exit.
  CASE l_ucomm.
    WHEN 'GOON' OR 'GRAVA' OR 'POST' OR 'OK'.


*      * to reflect the data changed into internal table
      DATA : ref_grid TYPE REF TO cl_gui_alv_grid. "new

      IF ref_grid IS INITIAL.
        CALL FUNCTION 'GET_GLOBALS_FROM_SLVC_FULLSCR'
          IMPORTING
            e_grid = ref_grid.
      ENDIF.

      IF NOT ref_grid IS INITIAL.
        CALL METHOD ref_grid->check_changed_data.
      ENDIF.

      ls_selfield-exit     = 'X'.

    WHEN 'CANCEL' OR '&AC1'.
      ls_selfield-exit     = 'X'.
    WHEN OTHERS.
  ENDCASE.

  ls_selfield-refresh = 'X'.
  ls_selfield-row_stable = 'X'.
  ls_selfield-col_stable = 'X'.

ENDFORM.                    " F4_USER_COMMAND

*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
FORM user_command1
  USING  l_ucomm        LIKE sy-ucomm
         ls_selfield    TYPE slis_selfield.
  CLEAR: ls_selfield-exit.
  CASE l_ucomm.
    WHEN 'CANCEL' OR 'OK'.
      ls_selfield-exit     = 'X'.
    WHEN OTHERS.
  ENDCASE.

  ls_selfield-refresh = 'X'.
  ls_selfield-row_stable = 'X'.
  ls_selfield-col_stable = 'X'.
ENDFORM.                    " USER_COMMAND1
*FORM data_changed USING  ir_data_changed TYPE REF TO cl_alv_changed_data_protocol.
*
*  DATA ls_modi TYPE lvc_s_modi.
** Check each modification:
*  LOOP AT ir_data_changed->mt_mod_cells INTO ls_modi.
*    CASE ls_modi-fieldname.
*      WHEN 'GJAHR'.
*        READ TABLE itab_doc INDEX ls_modi-row_id.
*        CHECK sy-subrc EQ 0.
*        itab_doc-gjahr = ls_modi-value.
*        MODIFY itab_doc INDEX  ls_modi-row_id.
*    ENDCASE.
*  ENDLOOP.
*ENDFORM. "Data_changed
**&---------------------------------------------------------------------*
**&      Form  ANEXOS
**&---------------------------------------------------------------------*
**       text
**----------------------------------------------------------------------*
**      -->P_HEADERDATA_ID_PANAGON  text
**      -->P_LS_ACCDN_BUKRS  text
**      -->P_LS_ACCDN_BELNR  text
**      -->P_LS_ACCDN_GJAHR  text
**----------------------------------------------------------------------*
*FORM anexos   USING  id_panagon
*                    bukrs
*                     belnr
*                     gjahr.
*
*  DATA object_id LIKE sapb-sapobjid.
*  IF id_panagon IS NOT INITIAL.
*    CONCATENATE bukrs belnr gjahr INTO object_id.
** anexa documento de panagon ao documento financeiro criado
*    CALL FUNCTION 'ARCHIV_CONNECTION_INSERT'
*      EXPORTING
*        archiv_id  = 'A3'
*        arc_doc_id = id_panagon
**       AR_DATE    = ' '
*        ar_object  = 'ZINV_CONSU'
**       DEL_DATE   = ' '
**       MANDANT    = ' '
*        object_id  = object_id
*        sap_object = 'BKPF'
*        doc_type   = 'ZFAX06'
**       BARCODE    = ' '
* EXCEPTIONS
*       ERROR_CONNECTIONTABLE       = 1
*       OTHERS     = 2
*      .
*    IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
*    ENDIF.
*
*  ENDIF.
*
*ENDFORM.                    " ANEXOS
*&---------------------------------------------------------------------*
*&      Form  EXIBE_FORN
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_VENDOR  text
*      -->P_HEADER_COMP_CODE  text
*      -->P_HEADER_ANO_LANC  text
*----------------------------------------------------------------------*
FORM exibe_forn  USING   vendor
                          comp_code.


  SET PARAMETER ID: 'LIF' FIELD vendor,
                    'BUK' FIELD comp_code.


  CALL TRANSACTION 'XK03' AND SKIP FIRST SCREEN.
ENDFORM.                    " EXIBE_FORN
*&---------------------------------------------------------------------*
*&      Form  EXIBE_USER
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER_USER_TRATAMENTO  text
*----------------------------------------------------------------------*
FORM exibe_user  USING    user.
  SET PARAMETER ID: 'XUS' FIELD user.


  CALL TRANSACTION 'SU01' AND SKIP FIRST SCREEN.
ENDFORM.                    " EXIBE_USER
*&---------------------------------------------------------------------*
*&      Form  EXIBE_VARIOS
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_HEADER  text
*----------------------------------------------------------------------*
FORM exibe_varios  USING    header TYPE /sbxc/zckp_invh.
  DATA: "fieldcat_linha TYPE slis_fieldcat_alv,
    fieldcat_tab TYPE slis_t_fieldcat_alv,
    grupos       TYPE slis_t_sp_group_alv,
*        wa_grupos      TYPE slis_sp_group_alv,
    wa_eventos   TYPE slis_alv_event,
    eventos      TYPE slis_t_event,
    layout       TYPE slis_layout_alv,
    is_variant   TYPE disvariant,
*        reprepid       TYPE slis_reprep_id,
    grid_set     TYPE lvc_s_glay.

  DATA:    wa_fieldcat       LIKE LINE OF fieldcat_tab.

  DATA: programa LIKE sy-repid.

  DATA: BEGIN OF itab_doc_alv OCCURS 0,
          comp_code LIKE /sbxc/zckp_tab14-comp_code,
          doc_fi    LIKE /sbxc/zckp_tab14-doc_fi,
          ano_lanc  LIKE /sbxc/zckp_tab14-ano_lanc,
          doc_lo    LIKE /sbxc/zckp_tab14-doc_lo,
        END OF itab_doc_alv.

  DATA: itab_doc LIKE itab_doc_alv OCCURS 0 WITH HEADER LINE.
  CLEAR layout.
*    layout-colwidth_optimize = 'X'.
  layout-zebra = ' '.
  layout-no_vline = ' '.
  layout-no_hline = ' '.
  layout-def_status = 'A'.
  layout-edit = ' '.
  layout-edit_mode = ' '.
  CONCATENATE TEXT-024 header-processo header-seqno header-ano INTO
  layout-window_titlebar SEPARATED BY space .
  is_variant-report = sy-repid.

  programa = sy-repid.
  grid_set-edt_cll_cb = 'X'.

* Eventos a capturar
  REFRESH eventos.
  REFRESH fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'COMP_CODE'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_TAB14'.
  wa_fieldcat-col_pos = 1.
*  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'DOC_FI'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_TAB14'.
  wa_fieldcat-col_pos = 2.
*  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'ANO_LANC'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_TAB14'.
  wa_fieldcat-col_pos = 3.
*  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO fieldcat_tab.

  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'DOC_LO'.
  wa_fieldcat-ref_tabname = '/SBXC/ZCKP_TAB14'.
  wa_fieldcat-col_pos = 4.
*  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO fieldcat_tab.
  wa_eventos-name = 'DATA_CHANGED'.
  wa_eventos-form = 'F_DATA_CHANGED'.
  APPEND wa_eventos TO eventos.

  SELECT * FROM /sbxc/zckp_tab14 INTO CORRESPONDING FIELDS OF TABLE  itab_doc
  WHERE processo = header-processo AND
  seqno = header-seqno AND
  ano = header-ano.
  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      it_fieldcat              = fieldcat_tab
      i_callback_user_command  = 'USER_COMMAND1'
*     i_callback_top_of_page   = 'TOP_OF_PAGE'
*     i_background_id          = 'ALV_BACKGROUND'
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
    TABLES
      t_outtab                 = itab_doc
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.



ENDFORM.                    " EXIBE_VARIOS
*&---------------------------------------------------------------------*
*&      Form  SEND_2_SAPHETY
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LS_XML  text
*      -->P_HEADER_UUID  text
*----------------------------------------------------------------------*
FORM send_2_saphety USING xml TYPE /sbxc/zsbx_st_img
                    ref_doc TYPE /sbxc/zckp_invh-ref_doc_no_orig
                    barcode TYPE /sbxc/zckp_invh-barcode.

  DATA: lr_saphety TYPE REF TO /sbxc/co_icnddocument_service.
*  DATA: lv_reason  TYPE string.
*  DATA: lv_content TYPE string.
*  DATA: lv_code    TYPE i.
*  DATA: lv_coden(3) TYPE n.
*  DATA: lv_messid  TYPE string.
  DATA: lv_process TYPE /sbxc/zckp_processo.
*  DATA: lv_chave   TYPE keyfield.
*  DATA: lv_user    TYPE uname.
*  DATA: lv_info    TYPE /bev1/rpmsgline.
*  DATA: lv_uuid    TYPE /sbxc/uuido.
  DATA: "lv_dummy TYPE /sbxc/uuido,
    base64 TYPE /sbxc/z_data_response.
*    t_uuid TYPE /sbxc/icnddocument_service_g27.

  DATA: lo_systemfault TYPE REF TO cx_ai_system_fault,
        lv_errortext   TYPE string.

  DATA: lv_data TYPE REF TO data.

  CONSTANTS: lc_error TYPE symsgty VALUE 'E'.

  CREATE DATA lv_data TYPE /sbxc/zsbx_st_img.

*  SELECT SINGLE * INTO @DATA(ls_tab30)                  "#EC CI_NOORDER
*  FROM /sbxc/zckp_tab30.

***  CLEAR lr_saphety.
***  lv_process = 'WSIMG'.
****
***  TRY .
***      CREATE OBJECT lr_saphety
***        EXPORTING
***          logical_port_name = 'BASICHTTPBINDING_ICNDDOCUMENTSERVICE'.
***    CATCH cx_ai_system_fault INTO lo_systemfault.
***      lv_errortext = lo_systemfault->get_text( ).
***      MESSAGE lv_errortext TYPE lc_error.
***  ENDTRY.
***
***  IF sy-subrc <> 0.
***    RETURN.
***  ENDIF.
***  FIELD-SYMBOLS: <fs> TYPE any.
***  ASSIGN lv_data->* TO <fs>.
***  MOVE-CORRESPONDING xml TO <fs>.
**** Criar XML em BASE64
***
* [REDACTED: sensitive line omitted from public export]
* [REDACTED: sensitive line omitted from public export]
***  t_uuid-in_transport_document_id = uuid.
***  t_uuid-doc_type = '2'.
***
***  TRY.
***      lr_saphety->get_document_data_on_in_transp(
****  lr_saphety->xml_create(
***          EXPORTING
***            input                  = t_uuid
***             IMPORTING
***                output = base64 ).
***    CATCH cx_ai_system_fault INTO lo_systemfault.
***      lv_errortext = lo_systemfault->errortext.
***      MESSAGE i081(/sbxc/zckp_cockpit).
***      RETURN.
***  ENDTRY.

  """"""
  DATA: go_saphety  TYPE REF TO /sbxc/zcl_read_doc_saphety,
        gtp_doc_b64 TYPE string.

  TRY.
      CREATE OBJECT go_saphety.
      IF go_saphety IS NOT BOUND. "Objecto não se encontra instanciado.
      ELSE.
        base64 = go_saphety->read_doc( i_barcode = barcode i_ref_doc = ref_doc ).
      ENDIF.
    CATCH zcx_config_not_found.

  ENDTRY.

  """"" FIM

  DATA: img_base64 TYPE /sbxc/z_data_response-contentdatabytes.
  DATA:
    v_url(255)        TYPE c.

  DATA: lt_tmp_content TYPE bapidoccontentab,
        comp           TYPE i.
*  DATA: lo_dialog_container TYPE REF TO cl_gui_dialogbox_container.
  DATA: lo_docking_container TYPE REF TO cl_gui_docking_container.
  DATA: lo_html    TYPE REF TO cl_gui_html_viewer.


*-DOC_DATA
  CLEAR: img_base64,  v_url .

  img_base64 = base64-contentdatabytes.

  IF img_base64 IS NOT INITIAL.
    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = img_base64 "imagem
*       APPEND_TO_TABLE       = ' '
      IMPORTING
        output_length = comp
      TABLES
        binary_tab    = lt_tmp_content.

    CREATE OBJECT lo_docking_container
      EXPORTING
        repid     = sy-repid
        dynnr     = sy-dynnr
        side      = lo_docking_container->dock_at_right
        extension = 1200.

    CREATE OBJECT lo_html
      EXPORTING
        parent = lo_docking_container.

    lo_html->load_data(
      EXPORTING
        type         = `application`
        subtype      = `pdf`
      IMPORTING
        assigned_url         = v_url
      CHANGING
        data_table           = lt_tmp_content
      EXCEPTIONS
        dp_invalid_parameter = 1
        dp_error_general     = 2
        cntl_error           = 3
        OTHERS               = 4 ).

*  CHECK v_url NE space.
    IF v_url NE space.
      lo_html->show_url( url = v_url  in_place = ' ' ).
    ENDIF.
  ENDIF.

ENDFORM.                    " SEND_2_SAPHETY
*&---------------------------------------------------------------------*
*&      Form  LER_IMAGEM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_LS_XML  text
*      -->P_COMP  text
*      -->P_HEADER_UUID  text
*      -->P_I_ATTACHX  text
*----------------------------------------------------------------------*
FORM ler_imagem  USING xml TYPE /sbxc/st_img
                    comp TYPE i
                    uuid TYPE /sbxc/zckp_invh-uuid
                    i_attachx   TYPE  solix_tab
                    barcode TYPE /sbxc/zckp_invh-barcode
                    ref_doc_no_orig TYPE /sbxc/zckp_invh-ref_doc_no_orig.

*data: lr_saphety type ref to zcl_im_outbound_saphety.
  DATA: lr_saphety TYPE REF TO /sbxc/co_icnddocument_service.
*  DATA: lv_reason  TYPE string.
*  DATA: lv_content TYPE string.
*  DATA: lv_code    TYPE i.
*  DATA: lv_coden(3) TYPE n.
*  DATA: lv_messid  TYPE string.
  DATA: lv_process TYPE /sbxc/zckp_processo.
*  DATA: lv_chave   TYPE keyfield.
*  DATA: lv_user    TYPE uname.
*  DATA: lv_info    TYPE /bev1/rpmsgline.
*  DATA: lv_uuid    TYPE /sbxc/uuido.
  DATA: "lv_dummy TYPE /sbxc/uuido,
* CCF Ini 26.01.2023 15:48:52
    base64 TYPE /sbxc/z_data_response,
*    base64 TYPE /sbxc/icnddocument_service_g26,
* CCF Fim 26.01.2023 15:48:52

    t_uuid TYPE /sbxc/icnddocument_service_g27.

  DATA: lv_data TYPE REF TO data.
  DATA: lt_system_fault TYPE REF TO cx_ai_system_fault,
        l_msg           TYPE string.
  CONSTANTS: lc_error TYPE symsgty VALUE 'E'.

  SELECT SINGLE * INTO @DATA(ls_tab30)                  "#EC CI_NOORDER
    FROM /sbxc/zckp_tab30 .

* CCF Ini 26.01.2023 15:35:29



*  CREATE DATA lv_data TYPE /sbxc/st_img.
*
*  CLEAR lr_saphety.
*  lv_process = 'WSIMG'.
*
*  TRY .
*      CREATE OBJECT lr_saphety
*        EXPORTING
*          logical_port_name = 'BASICHTTPBINDING_ICNDDOCUMENTSERVICE'.
*    CATCH cx_ai_system_fault INTO lt_system_fault.
*      l_msg = lt_system_fault->get_text( ).
*      MESSAGE l_msg TYPE lc_error.
*  ENDTRY.
*
*  IF sy-subrc <> 0.
*    RETURN.
*  ENDIF.
*  FIELD-SYMBOLS: <fs> TYPE any.
*  ASSIGN lv_data->* TO <fs>.
*  MOVE-CORRESPONDING xml TO <fs>.
** Criar XML em BASE64
*
* [REDACTED: sensitive line omitted from public export]
* [REDACTED: sensitive line omitted from public export]
*  t_uuid-in_transport_document_id = uuid.
*  t_uuid-doc_type = '2'.
*  TRY.
*      lr_saphety->get_document_data_on_in_transp(
*          EXPORTING
*            input                  = t_uuid
*             IMPORTING
*                output = base64 ).
*    CATCH cx_ai_system_fault INTO lt_system_fault.
*      l_msg = lt_system_fault->get_text( ).
*      MESSAGE l_msg TYPE lc_error.
*  ENDTRY.

  DATA: go_saphety  TYPE REF TO /sbxc/zcl_read_doc_saphety,
        gtp_doc_b64 TYPE string.

  TRY.
      CREATE OBJECT go_saphety.
      IF go_saphety IS NOT BOUND. "Objecto não se encontra instanciado.
      ELSE.
        base64 = go_saphety->read_doc( i_barcode = barcode i_ref_doc = ref_doc_no_orig ).
      ENDIF.
    CATCH zcx_config_not_found.

  ENDTRY.
  DATA: img_base64 TYPE /sbxc/z_data_response-contentdatabytes.
*  DATA: img_base64 TYPE /sbxc/icnddocument_service_g26-get_document_data_on_in_transp-doc_data.
* CCF Fim 26.01.2023 15:36:20

  DATA:
    v_url(255)        TYPE c.
*    v_application(11) TYPE c VALUE 'application'.

*  DATA: lt_tmp_content                TYPE bapidoccontentab.
*  DATA: lo_dialog_container TYPE REF TO cl_gui_dialogbox_container.
*  DATA: lo_docking_container TYPE REF TO cl_gui_docking_container.
*  DATA: lo_html    TYPE REF TO cl_gui_html_viewer.

*  DATA: objbin    LIKE solisti1   OCCURS 10 WITH HEADER LINE.
*  DATA: doc_chng  LIKE sodocchgi1.
*  DATA: objpack   LIKE sopcklsti1 OCCURS 2  WITH HEADER LINE.
*  DATA: tab_lines LIKE sy-tabix.
*  DATA email TYPE comm_id_long.
*  DATA: reclist   LIKE somlreci1  OCCURS 5  WITH HEADER LINE.
*-DOC_DATA
  CLEAR: img_base64,  v_url .
* CCF Ini 26.01.2023 15:37:58
  img_base64 = base64-contentdatabytes.
*  img_base64 = base64-get_document_data_on_in_transp-doc_data.

* CCF Fim 26.01.2023 15:37:58

*  CHECK img_base64  NE space.
  IF img_base64  NE space.
    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = img_base64 "imagem
      IMPORTING
        output_length = comp
      TABLES
        binary_tab    = i_attachx.
  ENDIF.
ENDFORM.                    " LER_IMAGEM


*&---------------------------------------------------------------------*
*&      Form  status_sub
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->RT_EXTAB   text
*----------------------------------------------------------------------*
FORM status_sub USING rt_extab TYPE slis_t_extab.
  SET PF-STATUS 'STATUS_P'.
ENDFORM.                    "status_sub
*&---------------------------------------------------------------------*
*&      Form  status_classm
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->RT_EXTAB   text
*----------------------------------------------------------------------*
FORM status_classm USING rt_extab TYPE slis_t_extab.
  DATA: it_exfcode LIKE TABLE OF rsmpe-func,
        wa_exfcode LIKE  rsmpe-func, status TYPE /sbxc/zckp_status1.
  SELECT SINGLE status1 FROM /sbxc/zckp_ctrl INTO status
                       WHERE processo = w_item_sub_ref-processo
                         AND ano      = w_item_sub_ref-ano
                         AND seqno    = w_item_sub_ref-seqno.

  IF status = '3' OR status = '4'.
    MOVE 'REM' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
    MOVE 'ADD' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
    MOVE 'SAVE' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
  ENDIF.
  SET PF-STATUS 'STATUS_P' EXCLUDING it_exfcode.
ENDFORM.                    "status_classm
*&---------------------------------------------------------------------*
*&      Form  user_command_sub
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->L_UCOMM      text
*      -->LS_SELFIELD  text
*----------------------------------------------------------------------*
FORM user_command_sub                                       "#EC CALLED
  USING  l_ucomm        LIKE sy-ucomm
         ls_selfield    TYPE slis_selfield.
  DATA: item       LIKE LINE OF t_item_sub,
*        w_item_sub_tot LIKE LINE OF t_item_sub_tot,
        w_item_sub LIKE LINE OF t_item_sub,
*        l_doc_item     TYPE /sbxc/zckp_invi-invoice_doc_item,
        tot_amount LIKE item-item_amount,
        tot_quant  LIKE item-quant_fact.
  CLEAR: ls_selfield-exit.
  CASE l_ucomm.
    WHEN 'ADD'.
      READ TABLE t_item_sub INTO item INDEX 1.
      IF sy-subrc = 0.
        CLEAR: item-gl_account,
               item-costcenter,
               item-item_amount,
               item-quantity.
        ls_selfield-refresh = 'X'.
        item-invoice_doc_item = 0.
        APPEND item TO t_item_sub.
      ENDIF.
    WHEN 'REM'.
      DELETE t_item_sub WHERE box = 'X'.
      ls_selfield-refresh = 'X'.
    WHEN 'SAVE'.
      ls_selfield-refresh = 'X'.
      CLEAR: tot_amount, tot_quant.
      IF w_header_ref-barcode = ''.
        LOOP AT t_item_sub INTO w_item_sub.
          tot_amount = tot_amount + w_item_sub-item_amount.
          tot_quant = tot_quant + w_item_sub-quantity.
        ENDLOOP.
        CLEAR ls_selfield-exit.
        IF NOT w_item_sub_ref-po_number IS INITIAL.
          IF  tot_amount = w_item_sub_ref-item_amount AND tot_quant = w_item_sub_ref-quantity.
            PERFORM gravar_divisao.
            ls_selfield-exit     = 'X'.
          ELSEIF  tot_amount <> w_item_sub_ref-item_amount AND tot_quant = w_item_sub_ref-quantity.
            MESSAGE i046(/sbxc/zckp_cockpit) WITH tot_amount w_item_sub_ref-item_amount.
          ELSEIF  tot_amount = w_item_sub_ref-item_amount AND tot_quant <> w_item_sub_ref-quantity.
            MESSAGE i045(/sbxc/zckp_cockpit) WITH tot_quant w_item_sub_ref-quantity.
          ELSEIF  tot_amount <> w_item_sub_ref-item_amount AND tot_quant <> w_item_sub_ref-quantity.
            MESSAGE i047(/sbxc/zckp_cockpit) WITH w_item_sub_ref-item_amount w_item_sub_ref-quantity.
          ENDIF.
        ELSE.
          IF  tot_amount = w_item_sub_ref-item_amount.
            PERFORM gravar_divisao.
            ls_selfield-exit     = 'X'.
          ELSE.
            MESSAGE i046(/sbxc/zckp_cockpit) WITH tot_amount w_item_sub_ref-item_amount.
          ENDIF.
        ENDIF.
      ELSE.
        PERFORM gravar_divisao.
        ls_selfield-exit     = 'X'.
      ENDIF.
    WHEN 'CANCELAR'.
      ls_selfield-exit     = 'X'.
  ENDCASE.

ENDFORM.                "set_status
*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND_CLASSM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->L_UCOMM      text
*      -->LS_SELFIELD  text
*----------------------------------------------------------------------*
FORM user_command_classm                                    "#EC CALLED
  USING  l_ucomm        LIKE sy-ucomm
         ls_selfield    TYPE slis_selfield.
  DATA: "w_item_sub     LIKE LINE OF t_item_sub,
    "w_item_sub_tot LIKE LINE OF t_item_sub_tot,
    tot_amount LIKE w_item_sub_ref-item_amount,
    tot_quant  LIKE w_item_sub_ref-quant_fact,
    l_doc_item LIKE ekkn-zekkn.
  CLEAR: ls_selfield-exit.
  CASE l_ucomm.
    WHEN 'ADD'.
      CLEAR: l_doc_item, itab_ekkn.
      LOOP AT itab_ekkn.
        IF l_doc_item LE itab_ekkn-zekkn.
          l_doc_item = itab_ekkn-zekkn.
        ENDIF.
      ENDLOOP.
      l_doc_item = l_doc_item + 1.
      CLEAR itab_ekkn.
      itab_ekkn-processo = w_item_sub_ref-processo.
      itab_ekkn-ano      = w_item_sub_ref-ano.
      itab_ekkn-seqno    = w_item_sub_ref-seqno.
      itab_ekkn-ebeln    = w_item_sub_ref-po_number.
      itab_ekkn-ebelp    = w_item_sub_ref-po_item.
      itab_ekkn-zekkn    = l_doc_item.
      APPEND itab_ekkn.
      ls_selfield-refresh = 'X'.
    WHEN 'REM'.
      DELETE itab_ekkn WHERE box = 'X'.
      ls_selfield-refresh = 'X'.
    WHEN 'SAVE'.
      ls_selfield-refresh = 'X'.
      CLEAR: tot_amount, tot_quant.
      IF w_header_ref-barcode = ''.
        LOOP AT itab_ekkn.
          tot_amount = tot_amount + itab_ekkn-wrbtr.
          tot_quant = tot_quant + itab_ekkn-menge.
        ENDLOOP.
        CLEAR ls_selfield-exit.
        IF  tot_amount = w_item_sub_ref-item_amount AND tot_quant = w_item_sub_ref-quantity.
          PERFORM gravar_divisaomm.
          ls_selfield-exit     = 'X'.
        ELSEIF  tot_amount <> w_item_sub_ref-item_amount AND tot_quant = w_item_sub_ref-quantity.
          MESSAGE i046(/sbxc/zckp_cockpit) WITH tot_amount w_item_sub_ref-item_amount.
        ELSEIF  tot_amount = w_item_sub_ref-item_amount AND tot_quant <> w_item_sub_ref-quantity.
          MESSAGE i045(/sbxc/zckp_cockpit) WITH tot_quant w_item_sub_ref-quantity.
        ELSEIF  tot_amount <> w_item_sub_ref-item_amount AND tot_quant <> w_item_sub_ref-quantity.
          MESSAGE i047(/sbxc/zckp_cockpit) WITH w_item_sub_ref-item_amount w_item_sub_ref-quantity.
        ENDIF.
      ELSE.
        PERFORM gravar_divisaomm.
        ls_selfield-exit     = 'X'.
      ENDIF.
    WHEN 'CANCELAR'.
      ls_selfield-exit     = 'X'.
  ENDCASE.

ENDFORM.                "set_status
*&---------------------------------------------------------------------*
*&      Form  gravar_divisao
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM gravar_divisao .
  DATA: w_item_sub     LIKE LINE OF t_item_sub,
        w_item_sub_tot LIKE LINE OF t_item_sub_tot,
        l_doc_item     LIKE w_item_sub_tot-invoice_doc_item.
  CLEAR l_doc_item.
  LOOP AT t_item_sub_tot INTO w_item_sub_tot.
    IF l_doc_item LE w_item_sub_tot-invoice_doc_item.
      l_doc_item = w_item_sub_tot-invoice_doc_item.
    ENDIF.
  ENDLOOP.
  DELETE t_item_sub_tot INDEX g_idx_lin_ref.
  LOOP AT t_item_sub INTO w_item_sub.
    IF sy-tabix = 1.
      w_item_sub-invoice_doc_item = w_item_sub_ref-invoice_doc_item.
    ELSE.
      l_doc_item = l_doc_item + 1.
      w_item_sub-invoice_doc_item = l_doc_item.
    ENDIF.
    "Recalcular montante de imposto e taxa de imposto
    SELECT SINGLE comp_code INTO /sbxc/zckp_invh-comp_code
      FROM /sbxc/zckp_invh
      WHERE processo EQ w_item_sub_ref-processo
      AND ano EQ w_item_sub_ref-ano
      AND seqno EQ w_item_sub_ref-seqno.

    "Det país da empresa
    SELECT SINGLE land1 INTO t001-land1
      FROM t001 WHERE bukrs = /sbxc/zckp_invh-comp_code.
*           Det. taxa iva
    SELECT SINGLE knumh INTO a003-knumh                 "#EC CI_NOORDER
      FROM a003
      WHERE kappl = 'TX'
      AND aland = t001-land1
      AND mwskz = w_item_sub-tax_code_sap.

    SELECT SINGLE kbetr INTO konp-kbetr
      FROM konp
      WHERE knumh = a003-knumh
      AND kopos = 1.
    w_item_sub-tax_imposto_sap = konp-kbetr / 10.

    w_item_sub-tax_amount = w_item_sub-item_amount * ( konp-kbetr / 1000 ).
    "Fim
    MODIFY t_item_sub FROM w_item_sub.
    MOVE-CORRESPONDING w_item_sub TO w_item_sub_tot.
    APPEND w_item_sub_tot TO t_item_sub_tot.
  ENDLOOP.
ENDFORM.                    " GRAVAR_DIVISAO
*&---------------------------------------------------------------------*
*&      Form  GRAVAR_DIVISAOMM
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM gravar_divisaomm .
  DELETE FROM /sbxc/zckp_ekkn WHERE
                   processo = w_item_sub_ref-processo  AND
                        ano = w_item_sub_ref-ano       AND
                      seqno = w_item_sub_ref-seqno     AND
                      ebeln = w_item_sub_ref-po_number AND
                      ebelp = w_item_sub_ref-po_item.
  LOOP AT itab_ekkn.
    MOVE-CORRESPONDING itab_ekkn TO /sbxc/zckp_ekkn.
    INSERT /sbxc/zckp_ekkn. CLEAR /sbxc/zckp_ekkn.
    COMMIT WORK.
  ENDLOOP.
ENDFORM.                    " GRAVAR_DIVISAOMM

*&---------------------------------------------------------------------*
*&      Form  grava_alteracoes
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_ITEMDATA  text
*      -->P_HEADERDATA  text
*----------------------------------------------------------------------*
FORM grava_alteracoes TABLES itemdata STRUCTURE /sbxc/zckp_invi
                      USING headerdata LIKE /sbxc/zckp_invh.

  DATA: item  TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE.
  DATA invoice_doc_item LIKE /sbxc/zckp_invi-invoice_doc_item.

  "Dados de cabeçalho
  MOVE-CORRESPONDING headerdata TO /sbxc/zckp_invh.
  IF /sbxc/zckp_invh-em_tratamento IS NOT INITIAL AND /sbxc/zckp_invh-user_tratamento IS INITIAL.
    /sbxc/zckp_invh-user_tratamento = sy-uname.
    /sbxc/zckp_invh-data_tratamento = sy-datum.
  ENDIF.
  UPDATE /sbxc/zckp_invh.
  COMMIT WORK AND WAIT.

  "Dados de item
  SORT itemdata  DESCENDING.


  LOOP AT itemdata.

    MOVE-CORRESPONDING itemdata TO  /sbxc/zckp_invi.
    IF /sbxc/zckp_invi-processo =  /sbxc/zckp_invh-processo AND
           /sbxc/zckp_invi-ano = /sbxc/zckp_invh-ano AND
            /sbxc/zckp_invi-seqno = /sbxc/zckp_invh-seqno.

      SELECT MAX( invoice_doc_item ) INTO  invoice_doc_item
      FROM   /sbxc/zckp_invi
      WHERE
            processo =  /sbxc/zckp_invi-processo AND
            ano =  /sbxc/zckp_invi-ano AND
            seqno =   /sbxc/zckp_invi-seqno.
      IF /sbxc/zckp_invi-invoice_doc_item IS INITIAL.
        ADD 1 TO invoice_doc_item.
        /sbxc/zckp_invi-invoice_doc_item = invoice_doc_item.
      ENDIF.

      MODIFY /sbxc/zckp_invi.
      COMMIT WORK.
    ENDIF.
    MOVE-CORRESPONDING itemdata TO  item.
    item-invoice_doc_item = /sbxc/zckp_invi-invoice_doc_item.
    APPEND item.
  ENDLOOP.

  COMMIT WORK.
*
  IF item[] IS NOT INITIAL.
** elimina da tabela de itens linhas entretanto apagadas
    SELECT * FROM  /sbxc/zckp_invi WHERE
         processo =  /sbxc/zckp_invi-processo AND
         ano =  /sbxc/zckp_invi-ano AND
        seqno =   /sbxc/zckp_invi-seqno.
*
      READ TABLE item WITH KEY
            processo = /sbxc/zckp_invi-processo
            ano = /sbxc/zckp_invi-ano
            seqno = /sbxc/zckp_invi-seqno
            invoice_doc_item = /sbxc/zckp_invi-invoice_doc_item.
*
      IF sy-subrc <> 0.
        DELETE /sbxc/zckp_invi FROM /sbxc/zckp_invi.
      ENDIF.
    ENDSELECT.

    REFRESH itemdata.
    LOOP AT item.
      MOVE-CORRESPONDING item TO itemdata.
      APPEND itemdata.
    ENDLOOP.

  ENDIF.
  SORT itemdata ASCENDING.

  COMMIT WORK AND WAIT.
ENDFORM.                    "grava_alteracoes
*&---------------------------------------------------------------------*
*& Form bi_miro
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM bi_miro USING header LIKE /sbxc/zckp_invh..
  DATA emp LIKE rbkp-bukrs.
  CLEAR messtab. REFRESH messtab.
  "
  CLEAR opt.
  opt-dismode  = 'E'.
  opt-updmode  = 'S'.
  opt-racommit = 'X'.
  opt-nobinpt = 'X'.
  REFRESH messtab.

*  GET PARAMETER ID 'BUK' FIELD emp. "ODC - 15_06_2020
  SET PARAMETER ID 'BUK' FIELD header-comp_code. "ODC - 15_06_2020

*  IF emp IS INITIAL. "ODC - 15_06_2020
  IF header-comp_code IS INITIAL. "ODC - 15_06_2020

    PERFORM dynpro USING:  'X' 'SAPLACHD'    '3010',
                           ' ' 'BDC_OKCODE'  '=ENTR'.
  ENDIF.

  PERFORM dynpro USING:  'X' 'SAPLMR1M'    '6000',
                         ' ' 'BDC_OKCODE'  '=DUMMY',
                         ' ' 'INVFO-BLDAT' header-doc_date,
                         ' ' 'INVFO-XBLNR' header-ref_doc_no,
                         ' ' 'INVFO-BUDAT' header-pstng_date,
                         ' ' 'INVFO-WRBTR' header-gross_amount,
                         ' ' 'INVFO-SGTXT' header-header_txt,
*                         ' ' 'INVFO-MWSKZ' header-MWSKZ,
                         ' ' 'INVFO-XMWST' 'X',
                         ' ' 'RM08M-REFERENZBELEGTYP' '1',
                         ' ' 'RM08M-EBELN' header-po_number.

  CALL TRANSACTION 'MIRO' USING bdcdata
                        OPTIONS FROM opt
                        MESSAGES INTO messtab.
ENDFORM.
