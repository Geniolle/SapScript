FUNCTION-POOL /sbxc/zckp_f MESSAGE-ID s4.

TYPE-POOLS: slis.

TABLES: /sbxc/zckp_ctrl,
        /sbxc/zckp_tab06,
        /sbxc/zckp_tab07,
        /sbxc/zckp_tab08,
        /sbxc/zckp_tab09,
        /sbxc/zckp_tab10,
        /sbxc/zckp_tab14,
        /sbxc/zckp_invh,
        /sbxc/zckp_invi,
        /sbxc/zckp_tab00,
        /sbxc/zckp_adnt,
        /sbxc/zckp_hinvh,
        /sbxc/zckp_hinvi,
        /sbxc/zckp_mfi,
        /sbxc/zckp_mped,
        /sbxc/zckp_img,
        /sbxc/zckp_ekkn,
        mara,
        dd03l,
        rbkp,
        lfbw,
        bkpf,
        ekkn,
        ekpo,
        ekko,
        bsip,
        bbkpf, bbseg, bbtax, bgr00,

        rseg,
        prps,
        jest.
FIELD-SYMBOLS:" <f1>,
               <field> type any.

DATA: BEGIN OF nametab OCCURS 120.
    INCLUDE STRUCTURE dntab.
DATA: END OF nametab.

DATA: BEGIN OF i_bgr00.
    INCLUDE STRUCTURE bgr00.
DATA: END OF i_bgr00.

DATA: BEGIN OF i_bbkpf.
    INCLUDE STRUCTURE bbkpf.
DATA: END OF i_bbkpf.

DATA: BEGIN OF i_bbseg.
    INCLUDE STRUCTURE bbseg.
DATA: END OF i_bbseg.

DATA: BEGIN OF it_glaccountdata OCCURS 0.
    INCLUDE STRUCTURE  bapi_incinv_create_gl_account.
DATA: END OF it_glaccountdata.

DATA it_accountingdata LIKE bapi_incinv_create_account
                        OCCURS 0 WITH HEADER LINE.

DATA: BEGIN OF bapi_account OCCURS 0.
    INCLUDE STRUCTURE bapimepoaccount.
DATA: END OF bapi_account.

DATA: BEGIN OF it_itemdata OCCURS 0.
    INCLUDE STRUCTURE  bapi_incinv_create_item.
DATA: END OF it_itemdata.

DATA: BEGIN OF item2 OCCURS 0.
    INCLUDE STRUCTURE /sbxc/zckp_invi.
DATA: END OF item2.

DATA: BEGIN OF bapi_item OCCURS 0.
    INCLUDE STRUCTURE bapimepoitem.
DATA: END OF bapi_item.

DATA: BEGIN OF poitemx OCCURS 0.
    INCLUDE STRUCTURE bapimepoitemx.
DATA: END OF poitemx.

DATA: BEGIN OF poaccountx OCCURS 0.
    INCLUDE STRUCTURE bapimepoaccountx.
DATA: END OF poaccountx.

DATA: BEGIN OF bapi_schedule OCCURS 0.
    INCLUDE STRUCTURE bapimeposchedule.
DATA: END OF bapi_schedule.

DATA: BEGIN OF poschedulex OCCURS 0.
    INCLUDE STRUCTURE bapimeposchedulx.
DATA: END OF poschedulex.

DATA: BEGIN OF potextitem OCCURS 0 .
    INCLUDE STRUCTURE  bapimepotext.
DATA: END OF   potextitem .

DATA: BEGIN OF bapi_header.
    INCLUDE STRUCTURE bapimepoheader.
DATA: END OF bapi_header.

DATA: BEGIN OF poheaderx.
    INCLUDE STRUCTURE bapimepoheaderx.
DATA: END OF poheaderx.

DATA: BEGIN OF t_field OCCURS 0,
        tabname   TYPE dd03l-tabname,
        fieldname TYPE dd03l-fieldname,
      END OF t_field.


DATA: BEGIN OF  wa_/sbxc/zckp_invi,
        po_number    LIKE /sbxc/zckp_invi-po_number,
        po_item      LIKE /sbxc/zckp_invi-po_item,
        item_text    LIKE /sbxc/zckp_invi-item_text,
        item_amount  LIKE /sbxc/zckp_invi-item_amount,
        tax_code_sap LIKE  /sbxc/zckp_invi-tax_code_sap,
        quantity     LIKE /sbxc/zckp_invi-quantity,
        po_unit      LIKE /sbxc/zckp_invi-po_unit,
        dc_posterior LIKE /sbxc/zckp_invi-dc_posterior,
        costcenter   LIKE /sbxc/zckp_invi-costcenter,
        gl_account   LIKE /sbxc/zckp_invi-gl_account,
        db_cr_ind    LIKE /sbxc/zckp_invi-db_cr_ind,
        ref_doc      LIKE /sbxc/zckp_invi-ref_doc,
        ref_doc_year LIKE /sbxc/zckp_invi-ref_doc_year,
        ref_doc_item LIKE /sbxc/zckp_invi-ref_doc_item,
      END OF     wa_/sbxc/zckp_invi.

DATA: BEGIN OF t_land OCCURS 0,
        mandant TYPE cdhdr-mandant,
        land1   TYPE lfa1-land1,
      END OF t_land.

DATA: BEGIN OF t_cdhdr OCCURS 0,
        objectid TYPE cdhdr-objectid,
      END OF t_cdhdr.


DATA: BEGIN OF t_linhas_sel OCCURS 0,
        indice TYPE sy-tabix,
      END OF t_linhas_sel.
****************************************************************
* Declaração de variaveis para BAPI de Movimentos de material
****************************************************************

DATA: gm_header TYPE bapi2017_gm_head_01.
DATA: gm_items TYPE STANDARD TABLE OF bapi2017_gm_item_create.
DATA gm_return LIKE STANDARD TABLE OF bapiret2.
DATA: w_gm_item TYPE bapi2017_gm_item_create.
DATA w_return LIKE bapiret2.
*DATA: return_s   TYPE bapirettab, return_swa TYPE bapirettab WITH HEADER LINE.
****************************************************************

* Declaração de variaveis para rotina de calculo de qtd a facturar
DATA: lt_xekbe        TYPE TABLE OF ekbe,
*      ls_xekbe        TYPE ekbe,
      lt_xekbes       TYPE TABLE OF ekbes,
      ls_xekbes       TYPE ekbes,
      lt_accchg       TYPE TABLE OF  accchg,
      ls_accchg       TYPE accchg,
      qtd_entrada     LIKE ls_xekbes-wemng,
      val_entrada     LIKE ls_xekbes-wewwr,
      qtd_facturada   LIKE ls_xekbes-remng,
      val_facturado   LIKE ls_xekbes-rewwr,
      qtd_ped         LIKE ls_xekbes-wemng,
      valor_ped       LIKE ls_xekbes-wewwr,
      qtd_remanesc    LIKE ls_xekbes-remng,
      dc_posterior(1).


DATA: l_returncode,
      cor_tab           TYPE TABLE OF /sbxc/zckp_color WITH HEADER LINE,
      t_fimsg           TYPE TABLE OF fimsg WITH HEADER LINE,
      c_nodata(1)       TYPE c VALUE '/',
      char(61)          TYPE c,
      ds_name(40)       VALUE 'COCKPIT_DOC',
      forn              LIKE lfa1-lifnr,
      pedido            LIKE w_gm_item-po_number,
      indice            LIKE sy-tabix,
      flag_tax,
      ret,
      obj               TYPE swotobjid-objkey,
      it_return         LIKE bapiret2 OCCURS 0 WITH HEADER LINE,
      t_return          LIKE bapiret2 OCCURS 0 WITH HEADER LINE,
      lt_accountgl      LIKE bapiacgl09 OCCURS 0 WITH HEADER LINE,
      lt_currencyamount LIKE bapiaccr09 OCCURS 0 WITH HEADER LINE,
      lt_accountpayable LIKE bapiacap09 OCCURS 0 WITH HEADER LINE,
      lt_extension1     LIKE bapiacextc OCCURS 0 WITH HEADER LINE,
      ls_documentheader LIKE bapiache09,
      sinal(1),
      moeda             LIKE lt_currencyamount-currency,
      l_awkey           TYPE  awkey,
      wa_header         TYPE  bapi_incinv_create_header,
      lin               TYPE i,
      ld_aworg          LIKE accdn-aworg,
      lt_accdn          TYPE TABLE OF accdn,
      ls_accdn          TYPE accdn,
      stcd1             LIKE lfa1-stcd1,
      stcd2             LIKE lfa1-stcd2,
      stenr             LIKE lfa1-stenr,
      num_end           LIKE  t_cdhdr-objectid,
      lt_poitem         LIKE STANDARD TABLE OF bapimepoitem,
      ls_poitem         LIKE bapimepoitem,
      lt_poitemx        LIKE STANDARD TABLE OF bapimepoitemx,
      ls_poitemx        LIKE bapimepoitemx,
      l_fieldname       TYPE c LENGTH 60,


* Estruturas de cabeçalho e linha
      header            TYPE TABLE OF /sbxc/zckp_hinvh WITH HEADER LINE,
      opcao             TYPE i.

*******************************************************************************
* VARIAVEIS GLOBAIS A VARIAIS FUNÇÕES
*******************************************************************************
*      Tipos para mensagens
TYPES: BEGIN OF message_type_wa,
         msgid  LIKE sy-msgid,
         msgty  LIKE sy-msgty,
         msgno  LIKE sy-msgno,
         msgv1  LIKE sy-msgv1,
         msgv2  LIKE sy-msgv2,
         msgv3  LIKE sy-msgv3,
         msgv4  LIKE sy-msgv4,
         lineno LIKE mesg-zeile,
       END OF message_type_wa.
TYPES: message_tab_type TYPE message_type_wa OCCURS 0.

DATA: msg_cockpit TYPE message_tab_type,
      wa_msg      LIKE LINE OF msg_cockpit,
      msg_ckp     TYPE message_tab_type,
      status_ok. " Validação status nas funções


*DATA: BEGIN OF  msg_email OCCURS 0 .
*    INCLUDE STRUCTURE  /sbxc/zckp_hmail.
*DATA: END OF  msg_email .

DATA:  uuid_orig     TYPE /sbxc/zuuido.
DATA: chave    TYPE  keyfield,                             "c 132
      inform   TYPE char80,
      processo TYPE  /sbxc/zws_process.

FIELD-SYMBOLS <fs> TYPE any.
*******************************************************************************
*  Declaração de variaveis para função de PRÉ-EDIÇÃO de  facturas
*******************************************************************************
DATA BEGIN OF bdcdata OCCURS 100.
INCLUDE STRUCTURE bdcdata.
DATA END OF bdcdata.

DATA: save_sy_tabix LIKE sy-tabix,
      messtab       LIKE bdcmsgcoll OCCURS 0 WITH HEADER LINE,
      opt           LIKE ctu_params.


*******************************************************************************
*  Declaração de variaveis para função de ESTORNO de
*  facturas
*******************************************************************************
DATA: lt_sval          LIKE sval OCCURS 0 WITH HEADER LINE,
      t_messtab        TYPE TABLE OF bdcmsgcoll,
      wa_messtab       LIKE LINE OF t_messtab,
      invoicedocnumber TYPE bapi_incinv_fld-inv_doc_no,
      fiscalyear       TYPE bapi_incinv_fld-fisc_year,
      reason_rev       LIKE bapi_incinv_fld-reason_rev,
      data_lanc        LIKE bapi_incinv_fld-pstng_date,
      return           TYPE TABLE OF bapiret2 WITH HEADER LINE,
      it_fimsg         TYPE TABLE OF fimsg WITH HEADER LINE.


TYPES: BEGIN OF ty_irf,
         lifnr TYPE lifnr,
         bukrs TYPE bukrs,
         irf   TYPE char1,
       END OF ty_irf.

DATA: gt_dd03l TYPE TABLE OF dd03l WITH HEADER LINE.
DATA: gt_tab10 TYPE TABLE OF /sbxc/zckp_tab10 WITH HEADER LINE.
DATA: gt_irf TYPE TABLE OF ty_irf WITH HEADER LINE.
DATA: tipo_doc LIKE /sbxc/zckp_invh-doctypesaphety.
DATA: BEGIN OF t_item_sub OCCURS 0,
        box(1).
    INCLUDE STRUCTURE /sbxc/zckp_invi.
DATA: END OF t_item_sub.
DATA: t_item_sub_tot LIKE TABLE OF /sbxc/zckp_invi,
      w_item_sub_ref LIKE /sbxc/zckp_invi,
      w_header_ref   LIKE /sbxc/zckp_invh,
      g_idx_lin_ref.
*INCLUDE mrm_types_basis.
INCLUDE mrm_types_nast.                " message determination

TYPE-POOLS: mmcr, mrm, cxtab.

TYPES t_drseg TYPE mmcr_drseg.
*------- Rechnungspositionen Dialog (angezeigte Positionen) ----------*
DATA: ydrseg TYPE mmcr_drseg OCCURS 1 WITH HEADER LINE.

DATA wa_ydrseg LIKE ydrseg.
*------- Rechnungspositionen Dialog (restliche Positionen) -----------*
DATA: ydrsegr TYPE mmcr_drseg OCCURS 1 WITH HEADER LINE.

*------- Mehrfachselektion Einkaufsbeleg -----------------------------*
DATA: BEGIN OF ymsel_best OCCURS 1.
    INCLUDE STRUCTURE rbselbest.
DATA: END OF ymsel_best.
DATA: BEGIN OF xmsel_best OCCURS 1.
    INCLUDE STRUCTURE rbselbest.
DATA: END OF xmsel_best.

*------- Mehrfachselektion Lieferschein ------------------------------*
DATA: BEGIN OF ymsel_lifs OCCURS 1.
    INCLUDE STRUCTURE rbsellifs.
DATA: END OF ymsel_lifs.
DATA: BEGIN OF xmsel_lifs OCCURS 1.
    INCLUDE STRUCTURE rbsellifs.
DATA: END OF xmsel_lifs.

*------- Mehrfachselektion Erfassungsblätter -------------------------*
DATA: BEGIN OF ymsel_erfb OCCURS 1.
    INCLUDE STRUCTURE rbselerfb.
DATA: END OF ymsel_erfb.
DATA: BEGIN OF xmsel_erfb OCCURS 1.
    INCLUDE STRUCTURE rbselerfb.
DATA: END OF xmsel_erfb.

*------- Mehrfachselektion Frachtbrief -------------------------------*
DATA: BEGIN OF ymsel_frbr OCCURS 1.
    INCLUDE STRUCTURE rbselfrbr.
DATA: END OF ymsel_frbr.
DATA: BEGIN OF xmsel_frbr OCCURS 1.
    INCLUDE STRUCTURE rbselfrbr.
DATA: END OF xmsel_frbr.

*------- Mehrfachselektion Werk --------------------------------------*
DATA: BEGIN OF ymsel_werk OCCURS 1.
    INCLUDE STRUCTURE rbselwerk.
DATA: END OF ymsel_werk.
DATA: BEGIN OF xmsel_werk OCCURS 1.
    INCLUDE STRUCTURE rbselwerk.
DATA: END OF xmsel_werk.

*------- Mehrfachselektion Transport zum Lieferant -------------------*
DATA: BEGIN OF ymsel_tran OCCURS 1.
    INCLUDE STRUCTURE eksel.
DATA: END OF ymsel_tran.
DATA: BEGIN OF xmsel_tran OCCURS 1.
    INCLUDE STRUCTURE letra_iv_fields.
DATA: END OF xmsel_tran.

**------- Fehlerprotokoll ---------------------------------------------*
DATA t_errprot TYPE mrm_errprot OCCURS 100 WITH HEADER LINE.


*------- Docking/Tree für Bestell-Historie ----------------------------*
DATA: t_ebelntab      TYPE TABLE OF ebelntab WITH HEADER LINE,  " PO's
      f_text_icon(20) TYPE c,          " dynamic Tree-Button
      dock_visible.

DATA:    xlimit TYPE mmcr_tlimit.

DATA: BEGIN OF rbkpv  OCCURS 1.
    INCLUDE TYPE  mrm_rbkpv.
DATA: END OF  rbkpv.

DATA wa_rbkpv LIKE rbkpv.
DATA wa_xmsel_best LIKE xmsel_best.



DATA: BEGIN OF itab_ekkn_alv OCCURS 0,
        processo   LIKE /sbxc/zckp_ekkn-processo,
        ano        LIKE /sbxc/zckp_ekkn-ano,
        seqno      LIKE /sbxc/zckp_ekkn-seqno,
        ebeln      LIKE ekkn-ebeln,
        ebelp      LIKE ekkn-ebelp,
        zekkn      LIKE ekkn-zekkn,
        vproz      LIKE ekkn-vproz,
        sakto      LIKE ekkn-sakto,
        kostl      LIKE ekkn-kostl,
        ps_psp_pnr LIKE ekkn-ps_psp_pnr,
        aufnr      LIKE ekkn-aufnr,
        wrbtr      TYPE wrbtr,
        menge      LIKE cobl_mrm_d-menge,
        box(1),
      END OF itab_ekkn_alv.

DATA: itab_ekkn LIKE itab_ekkn_alv OCCURS 0 WITH HEADER LINE.


DATA: item_aux  TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.

DATA: it_mwdat LIKE TABLE OF rtax1u15,
      wa_mwdat LIKE rtax1u15.

"Variaveis para função de rejeição
DATA: v_cod_mes  TYPE /sbxc/zckp_tab07-cod_mes,
      v_mensagem TYPE /sbxc/zckp_tab07-mensagem.

CONTROLS: sc_tc TYPE TABLEVIEW USING SCREEN 0001.

TYPES: BEGIN OF ty_ped,
         ebeln TYPE ebeln,
         ebelp TYPE ebelp,
       END OF ty_ped.
DATA: lt_ped TYPE TABLE OF ty_ped,
      ls_ped TYPE ty_ped.
