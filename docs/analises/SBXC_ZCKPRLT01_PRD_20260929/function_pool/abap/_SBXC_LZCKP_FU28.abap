FUNCTION /sbxc/zckp_mm_simula.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(UCOMM_SIM) TYPE  SY-UCOMM OPTIONAL
*"     REFERENCE(BLINE_DATE) TYPE  BAPI_INCINV_CREATE_HEADER-BLINE_DATE
*"       OPTIONAL
*"  EXPORTING
*"     VALUE(INVOICEDOCNUMBER) TYPE  BAPI_INCINV_FLD-INV_DOC_NO
*"     VALUE(FISCALYEAR) TYPE  BAPI_INCINV_FLD-FISC_YEAR
*"     VALUE(UCOMM) TYPE  SY-UCOMM
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI
*"      RETURN STRUCTURE  BAPIRET2
*"  CHANGING
*"     VALUE(HEADERDATA) TYPE  /SBXC/ZCKP_INVH
*"     VALUE(CTRL) TYPE  /SBXC/ZCKP_CTRL
*"----------------------------------------------------------------------

* Função para criar facturas
*********************************************************************

  DATA: it_accit TYPE accit_t,
        t_acccr TYPE  acccr_t,
        return2 TYPE  bapirettab,
        it_acccr TYPE acccr_t,
        f4_layout_alv TYPE slis_layout_alv,
        ls_vari TYPE disvariant,
        wa_accit LIKE accit,
        wa_acccr like acccr,
        it_fieldcat TYPE slis_t_fieldcat_alv,
        wa_fieldcat TYPE slis_fieldcat_alv,
        s_return TYPE bapiret2,
        zckp_accit_t TYPE /sbxc/simula_accit_t WITH HEADER LINE,
        wa_return TYPE bapiret2,
        f_ktopl TYPE ktopl.

  DATA: ls_itemdata TYPE /sbxc/zckp_invi,
        ok TYPE flag.
  data: lv_fin TYPE flag.


* Preenche estruturas BAPI de facturas
  REFRESH: it_itemdata,  it_return, it_glaccountdata, return,  it_accountingdata.
  CLEAR wa_header.


*"Verificar se existe alguma linha com informação de imobilizado para não permitir simulação
*  CLEAR ok.
*  LOOP AT itemdata WHERE anln1 IS NOT INITIAL.
*    ok = abap_true.
*  ENDLOOP.
*  IF ok IS NOT INITIAL.
*    wa_return-type = 'I'.
*    wa_return-id   = '/SBXC/ZCKP_COCKPIT'.
*    wa_return-number = '059'.
*    APPEND wa_return TO return.
*  ENDIF.

  CHECK ok IS INITIAL.


  "Verificar se é uma simulação financeira
  clear lv_fin.
  LOOP AT itemdata INTO ls_itemdata WHERE po_number is NOT INITIAL.
    exit.
  ENDLOOP.
  IF sy-subrc ne 0.
    lv_fin = 'X'.
  ENDIF.
  export lv_fin from lv_fin to MEMORY id 'LV_FIN'.

  READ TABLE  itemdata WITH KEY po_number = ' '.

  PERFORM preenche_estruturas_inv
            TABLES itemdata
            USING headerdata.


  SORT it_itemdata BY invoice_doc_item.
  SORT it_glaccountdata BY invoice_doc_item.
*  SORT it_accountingdata BY invoice_doc_item ASCENDING serial_no ASCENDING.

  CALL FUNCTION 'MRM_PROT_RESET'
    .

  CALL FUNCTION 'MRM_PROT_INIT'
    EXCEPTIONS
      OTHERS = 0.

  CALL FUNCTION 'MRM_DBTAB_REFRESH'.

  CALL FUNCTION 'MRM_PUFFER_REFRESH'.

  CALL FUNCTION 'MESSAGES_INITIALIZE'.


*"----------------------------------------------------------------
* BAPI
*"----------------------------------------------------------------
  REFRESH: it_return, msg_ckp .
  refresh return2.
  "CCF 21.01.2022
  wa_header-BLINE_DATE = BLINE_DATE.
  CALL FUNCTION 'MRM_SRM_INVOICE_SIMULATE'
    EXPORTING
      headerdata          = wa_header
*     ADDRESSDATA         =
    IMPORTING
      return              = return2
      t_accit             = it_accit
      t_acccr             = it_acccr
    TABLES
      itemdata            = it_itemdata
      accountingdata      = it_accountingdata "#EC CI_USAGE_OK[2438006]
      glaccountdata       = it_glaccountdata
*     MATERIALDATA        =
*     TAXDATA             =
*     WITHTAXDATA         =
*     VENDORITEMSPLITDATA =
    .

  LOOP AT return2 INTO wa_return.
    wa_return-type = 'I'.
    COLLECT wa_return INTO return.
  ENDLOOP.

  IF return[] IS NOT INITIAL.
* Exibir msg de erro
*    CALL FUNCTION 'C14ALD_BAPIRET2_SHOW'
*      TABLES
*        i_bapiret2_tab = return.
  ENDIF.


IF UCOMM_SIM ne 'AUT'. "ODC - 10_06_2020

*** Mostrar ALV em Popup *********************************************
  f4_layout_alv-window_titlebar   = text-036. "'Simulação'.
  f4_layout_alv-colwidth_optimize = 'X'.
  f4_layout_alv-zebra             = 'X'.
  f4_layout_alv-numc_sum         = 'X'.
  f4_layout_alv-totals_text = 'SALDO'.

  wa_fieldcat-do_sum = 'X'.
  wa_fieldcat-fieldname = 'PSWBT'.
  APPEND wa_fieldcat TO it_fieldcat.

  ls_vari-report =  sy-repid.
  LOOP AT it_accit INTO wa_accit.
    MOVE-CORRESPONDING wa_accit TO zckp_accit_t.
    IF zckp_accit_t-pswbt IS INITIAL.
      read TABLE it_acccr INTO wa_acccr with key posnr = wa_accit-posnr
                                                 curtp = '00'.
      IF sy-subrc eq 0.
       zckp_accit_t-pswbt = wa_acccr-wrbtr.
      ENDIF.
    ENDIF.
    SELECT SINGLE ktopl FROM t001 INTO f_ktopl WHERE bukrs = wa_accit-bukrs.
    SELECT SINGLE txt50 FROM skat INTO zckp_accit_t-txt50 WHERE spras = sy-langu AND ktopl = f_ktopl  AND saknr = wa_accit-hkont.
    IF  zckp_accit_t-lifnr IS NOT INITIAL.
      SELECT SINGLE name1 FROM lfa1 INTO zckp_accit_t-name1
        WHERE lifnr = zckp_accit_t-lifnr.
    ENDIF.
    APPEND zckp_accit_t.
  ENDLOOP.

*  LOOP AT return_s INTO return_swa .
*    MOVE-CORRESPONDING return_swa TO return.
*    return-type = 'S'.
*    APPEND return.
*  ENDLOOP.
*  IF return[] IS NOT INITIAL.
** Exibir msg de erro
*    CALL FUNCTION 'C14ALD_BAPIRET2_SHOW'
*      TABLES
*        i_bapiret2_tab = return.
*  ENDIF.


  IF zckp_accit_t[] IS NOT INITIAL.
    CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
      EXPORTING
        i_callback_program       = sy-repid
        i_callback_pf_status_set = 'STATUS'
        i_callback_user_command  = 'USER_COMMAND'
*        i_callback_top_of_page   = 'F4_TOP_OF_PAGE '
        i_structure_name         = '/SBXC/SIMULA_ACCIT'
        is_layout                = f4_layout_alv
        it_fieldcat              = it_fieldcat
*        i_save                   = l_save
*        is_variant               = ls_vari
        i_screen_start_column    = 1
        i_screen_start_line      = 1
        i_screen_end_column      = 170
        i_screen_end_line        = 10
*      IMPORTING
*        es_exit_caused_by_user   = es_exit_caused_by_user
      TABLES
        t_outtab                 = zckp_accit_t "t_accit #EC CI_FLDEXT_OK[2610650]
      EXCEPTIONS
        program_error            = 1
        OTHERS                   = 2.

  ENDIF.

CALL FUNCTION 'DEQUEUE_ALL'.
ENDIF.
  ucomm	=	sy-ucomm.
*********************************************************************
ENDFUNCTION.
