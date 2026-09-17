*&---------------------------------------------------------------------*
*& Include          ZBIM_MONITOR_FRMS
*&---------------------------------------------------------------------*

FORM f_check_adm_block.

*  IF sy-uname = 'OGOMES' OR sy-uname = '15345' OR sy-uname = 'SAP_WFRT' OR
*   sy-uname = '107962' OR sy-uname = '246525' OR  sy-uname = '294923' OR
*   sy-uname = '17319' or sy-uname = '10148' or sy-uname = '299962'.
*
*
*
*  ELSE.
  SELECT SINGLE *
      FROM zbim_param_t
      INTO @DATA(ls_param_t)
      WHERE zprocess = 'BIM'
        AND fname = 'ADM_BLOCK'.

  IF sy-subrc = 0.
    IF ls_param_t-low = 'X'.
      CALL FUNCTION 'POPUP_TO_INFORM'
        EXPORTING
          titel = 'MONITOR BLOCKED'
          txt1  = 'Currently the monitor is blocked'
          txt2  = 'for maintenance purposes.'
*         TXT3  = ' '
*         TXT4  = ' '
        .
      gv_adm_block = abap_true.
    ELSE.
    ENDIF.
  ENDIF.
*  ENDIF.
ENDFORM.
FORM f_update_before_run.

  DATA: lt_invoices   TYPE STANDARD TABLE OF zbim_blk_invoice,
        lt_inv_fi     TYPE STANDARD TABLE OF zbim_blk_inv_fi,
        ls_workitem   TYPE swlwp1,
        lv_agent      TYPE sww_aagent,
        lv_objtype    TYPE sibftypeid,
        lv_event      TYPE sibfevent,
        lv_objkey     TYPE sibfinstid,
        lv_param_name TYPE swfdname,
        lv_buzei      TYPE p DECIMALS 2,
        lv_buzei_orig TYPE rblgp.

  IF p_log IS NOT INITIAL.
    SELECT *
      FROM zbim_blk_invoice
      INTO CORRESPONDING FIELDS OF TABLE lt_invoices
      WHERE zstatus <> 'CLOSED'
        AND ( wi_stat <> 'ENCERRADO'
         OR wi_stat <> 'ELIM.LOGICAM' )
        AND bukrs IN s_bukrs
            AND belnr IN s_belnr
            AND gjahr IN s_gjahr
            AND ebeln IN s_ebeln
            AND ( spgrp = p_spgrp OR
                  spgrm =  p_spgrm OR
                  zblck_razao = p_razao )
            AND wi_cd IN s_wi_cd
            AND wi_id IN s_wiid
            AND zmblnr_inv IN s_docinv
            AND zmjahr_inv IN s_invdat
            AND zmblnr_em IN s_docem
            AND zmjahr_em IN s_emdat
            AND zbelnr_nc IN s_doc_nc
            AND zgjahr_nc IN s_ncdate
            AND zeop_date IN s_zeop
            AND zdoc_estorno IN s_docest
            AND lifnr IN  s_lifnr
            AND zdebit_note IN s_nc_nd.

    IF sy-subrc = 0.

      SELECT a~xblnr, b~belnr, b~gjahr,b~buzei, b~ebeln, b~ebelp, b~spgrp, b~spgrm, a~zlspr
        FROM rbkp AS a INNER JOIN rseg AS b ON a~belnr = b~belnr
                                           AND a~gjahr = b~gjahr
        INTO TABLE @DATA(lt_rseg)
        FOR ALL ENTRIES IN @lt_invoices
        WHERE b~belnr = @lt_invoices-belnr
          AND b~gjahr = @lt_invoices-gjahr.


      LOOP AT lt_invoices ASSIGNING FIELD-SYMBOL(<fs>).
        lv_objtype = 'ZCL_WF_BIM_REQ'.
        lv_event = 'NO_DIV'.

*        CONCATENATE <fs>-belnr <fs>-bukrs <fs>-gjahr INTO DATA(lv_awkey).
* CCF 26.10.2022 tiha sido colocado a empresa neste concatenate, mas se uma fatura é logistica
* o sistena na referencia nao coloca a empresa
        CONCATENATE <fs>-belnr <fs>-gjahr INTO DATA(lv_awkey).
        SELECT SINGLE b~zlspr
          FROM bkpf AS a INNER JOIN bseg AS b ON
            a~bukrs = b~bukrs AND
            a~belnr = b~belnr AND
            a~gjahr = b~gjahr
          INTO <fs>-zlspr
          WHERE a~awkey = lv_awkey
            AND b~koart = 'K'.


        IF <fs>-zbuzei_orig IS INITIAL.
          lv_buzei_orig = <fs>-buzei.
        ELSE.
          lv_buzei_orig = <fs>-zbuzei_orig.
        ENDIF.

        "Verificar se diferença de quantidade ainda existe
        IF <fs>-spgrm IS NOT INITIAL.
          CONCATENATE <fs>-bukrs <fs>-belnr <fs>-gjahr 'QTD' INTO lv_objkey.

          READ TABLE lt_rseg INTO DATA(ls_rseg) WITH KEY belnr = <fs>-belnr
                                                         gjahr = <fs>-gjahr
                                                         buzei = lv_buzei_orig
                                                         spgrm = 'X'.
          IF sy-subrc <> 0.
            <fs>-zstatus = 'CLOSED'.
            CALL FUNCTION 'ZBIM_CLOSE_WORKITEM'
              EXPORTING
                i_objtype    = lv_objtype
                i_event      = lv_event
                i_objkey     = lv_objkey
                i_param_name = 'BUZEI'
                i_buzei      = <fs>-buzei
                i_wiid       = <fs>-wi_id
                i_status     = <fs>-zstatus.

            <fs>-zdiv_closed = abap_true.
            <fs>-zeop_date = sy-datum.
          ENDIF.
          "Verificar se diferença de preço ainda existe
        ELSEIF <fs>-spgrp IS NOT INITIAL.
          CONCATENATE <fs>-bukrs <fs>-belnr <fs>-gjahr 'PRC' INTO lv_objkey.
          READ TABLE lt_rseg INTO ls_rseg WITH KEY belnr = <fs>-belnr
                                                   gjahr = <fs>-gjahr
                                                   buzei = lv_buzei_orig
                                                   spgrp = 'X'.
          IF sy-subrc <> 0.
            <fs>-zstatus = 'CLOSED'.
            CALL FUNCTION 'ZBIM_CLOSE_WORKITEM'
              EXPORTING
                i_objtype    = lv_objtype
                i_event      = lv_event
                i_objkey     = lv_objkey
                i_param_name = 'BUZEI'
                i_buzei      = <fs>-buzei
                i_wiid       = <fs>-wi_id
                i_status     = <fs>-zstatus
              IMPORTING
                o_wiid       = <fs>-wi_id
                o_wistat     = <fs>-wi_stat
                o_wicruser   = lv_agent.

            <fs>-zdiv_closed = 'X'.
            <fs>-zeop_date = sy-datum.
          ENDIF.
        ENDIF.

        "obter referencia da fatura do fornecedor
        READ TABLE lt_rseg INTO ls_rseg WITH KEY belnr = <fs>-belnr
                                                 gjahr = <fs>-gjahr.

        IF sy-subrc = 0.
          <fs>-xblnr = ls_rseg-xblnr.
        ENDIF.
        ls_workitem-wi_id = <fs>-wi_id.

        CALL FUNCTION 'ZBIM_GET_WIID'
          EXPORTING
            i_workitem = ls_workitem
            i_status   = <fs>-zstatus
          IMPORTING
            o_wi_id    = <fs>-wi_id
            o_status   = <fs>-wi_stat
            o_agent    = lv_agent.

        IF lv_agent IS INITIAL.
*          <fs>-wi_cruser = lv_agent.
          SELECT * FROM
            swwuserwi
            INTO TABLE @DATA(lt_userwi)
            WHERE wi_id = @<fs>-wi_id.

          IF sy-subrc = 0.
            CLEAR <fs>-wi_cruser.
            DESCRIBE TABLE lt_userwi LINES DATA(lv_lines).
            LOOP AT lt_userwi INTO DATA(ls_userwi).
              IF lv_lines > 1.
                CONCATENATE ls_userwi-user_id <fs>-wi_cruser INTO <fs>-wi_cruser SEPARATED BY '|'.
              ELSE.
                <fs>-wi_cruser = ls_userwi-user_id.
              ENDIF.
            ENDLOOP.
          ENDIF.
        ELSE.
          <fs>-wi_cruser = lv_agent.
        ENDIF.

        TRANSLATE <fs>-wi_stat TO UPPER CASE.

        IF <fs>-ekgrp IS INITIAL OR
           <fs>-matnr IS INITIAL OR
           <fs>-maktx IS INITIAL.
          SELECT SINGLE a~ekgrp, b~matnr
              FROM ekko AS a INNER JOIN ekpo AS b
              ON a~ebeln = b~ebeln
              INTO (@<fs>-ekgrp, @<fs>-matnr)
              WHERE b~ebeln = @<fs>-ebeln
                AND b~ebelp = @<fs>-ebelp.

          SELECT SINGLE maktx
            FROM makt
            INTO <fs>-maktx
            WHERE matnr = <fs>-matnr.
        ENDIF.
      ENDLOOP.
    ENDIF.

    MODIFY zbim_blk_invoice FROM TABLE lt_invoices.
  ELSEIF p_fi IS NOT INITIAL.
    SELECT *
       FROM zbim_blk_inv_fi
       INTO CORRESPONDING FIELDS OF TABLE lt_inv_fi
       WHERE zstatus <> 'CLOSED'
         AND ( wi_stat <> 'ENCERRADO'
          OR wi_stat <> 'ELIM.LOGICAM' )
         AND bukrs IN s_bukrs
         AND gjahr IN s_gjahr
         AND lifnr IN s_lifnr
         AND wi_id IN s_wiid
         AND zeop_date IN s_zeop
         AND wi_cd IN s_wi_cd.

    LOOP AT lt_inv_fi[] ASSIGNING FIELD-SYMBOL(<fs1>).
      lv_objtype = 'ZCL_WF_BIM_REQ'.
      lv_event = 'NO_DIV'.

      ls_workitem-wi_id = <fs1>-wi_id.
      CLEAR lv_awkey.

      CONCATENATE <fs1>-belnr <fs1>-bukrs <fs1>-gjahr INTO lv_awkey.

      SELECT SINGLE b~zlspr
        FROM bkpf AS a INNER JOIN bseg AS b ON
          a~bukrs = b~bukrs AND
          a~belnr = b~belnr AND
          a~gjahr = b~gjahr
        INTO <fs1>-zlspr
        WHERE a~awkey = lv_awkey
          AND b~koart = 'K'.


      "beg - mbovo   IT-31919 - fechar documentos de fi
      IF <fs1>-zlspr IS INITIAL.
        <fs1>-zstatus = 'CLOSED'.
        DATA t_PARAM_VALUE LIKE rseg-buzei.
        t_PARAM_VALUE = <fs1>-buzei.
        CONCATENATE <fs1>-bukrs <fs1>-belnr <fs1>-gjahr 'FI' INTO lv_objkey.
        CALL FUNCTION 'ZBIM_RAISE_EVENT'
          EXPORTING
            i_objtype       = lv_objtype
            i_event         = lv_event
            i_objkey        = lv_objkey
            i_param_name    = 'BUZEI'
            i_param_value   = t_PARAM_VALUE.
            <fs1>-zeop_date = sy-datum.
      ENDIF.
      "end - mbovo   IT-31919 - fechar documentos de fi

      CALL FUNCTION 'ZBIM_GET_WIID'
        EXPORTING
          i_workitem = ls_workitem
          i_status   = <fs1>-zstatus
        IMPORTING
          o_wi_id    = <fs1>-wi_id
          o_status   = <fs1>-wi_stat
          o_agent    = lv_agent.

      IF lv_agent IS INITIAL.
*          <fs>-wi_cruser = lv_agent.
        SELECT * FROM
          swwuserwi
          INTO TABLE @DATA(lt_user_wi)
          WHERE wi_id = @<fs1>-wi_id
            AND no_sel = ' '.

        IF sy-subrc = 0.
          CLEAR <fs1>-wi_cruser.
          DESCRIBE TABLE lt_user_wi LINES DATA(lv_lines1).
          LOOP AT lt_user_wi INTO DATA(ls_user_wi).
            IF lv_lines1 > 1.
              CONCATENATE ls_user_wi-user_id <fs1>-wi_cruser INTO <fs1>-wi_cruser SEPARATED BY '|'.
            ELSE.
              <fs1>-wi_cruser = ls_user_wi-user_id.
            ENDIF.
          ENDLOOP.
        ENDIF.
      ENDIF.
    ENDLOOP.

    MODIFY zbim_blk_inv_fi FROM TABLE lt_inv_fi.
  ENDIF.


ENDFORM.
FORM f_select_data.

  DATA:
    lt_t001      TYPE STANDARD TABLE OF t001,
    ls_t001      LIKE LINE OF lt_t001,
    lt_taba      TYPE STANDARD TABLE OF dd07v,
    lt_tabb      TYPE STANDARD TABLE OF dd07v,
    ls_taba      TYPE dd07v,
    lv_uname     TYPE string,
    ls_workitem  TYPE swlwp1,
    lt_steps     TYPE STANDARD TABLE OF swl_pm_cvh,
    ls_steps     LIKE LINE OF lt_steps,
    lt_merge_t   TYPE swww_t_merge_table,
    ls_stat      LIKE LINE OF s_stat,
    ls_zstatus   LIKE LINE OF s_status,
    lv_diftot    TYPE z_remng,
    lv_end_of_gr TYPE abap_bool,
    lv_tabix     TYPE sy-tabix.

  DATA: lv_objtype TYPE sibftypeid,
        lv_event   TYPE sibfevent,
        lv_objkey  TYPE sibfinstid,
        lv_agent   TYPE sww_aagent.

  DATA: lt_xekbe    TYPE TABLE OF ekbe,
        ls_xekbe    TYPE ekbe,
        lt_xekbes   TYPE TABLE OF ekbes,
        ls_xekbes   TYPE ekbes,
        ls_log_data LIKE LINE OF gt_log_data,
        ls_fi_data  LIKE LINE OF gt_fi_data.

  DATA: lv_buzei TYPE rseg-buzei.

  CONCATENATE '%' sy-uname '%' INTO lv_uname.

  LOOP AT s_stat ASSIGNING FIELD-SYMBOL(<fs_stat>).
    CASE <fs_stat>-low.
      WHEN 'ELIM.LOGIC' OR 'DELETED LOGI'.
        <fs_stat>-option = 'CP'.
        CONCATENATE <fs_stat>-low '*' INTO <fs_stat>-low.
      WHEN 'EM PROCESS' OR 'IN PROCESS'.
        <fs_stat>-option = 'CP'.
        CONCATENATE <fs_stat>-low '*' INTO <fs_stat>-low.
      WHEN OTHERS.
    ENDCASE.
  ENDLOOP.

*  IF sy-uname = 'OGOMES' OR sy-uname = '15345' or sy-uname = 'SAP_WFRT' or
*    sy-uname = ' 107962' or sy-uname = '246525'.
*    teste_sbx = 'X'.
*  ENDIF.



  SET LANGUAGE sy-langu.
  IF p_log IS NOT INITIAL.
    "Todos os workitems por iniciar para o utilizador
    IF p_mywidt IS NOT INITIAL.
      ls_stat-option = 'EQ'.
      ls_stat-sign = 'I'.
      ls_stat-low = TEXT-011.
      APPEND ls_stat TO s_stat.

      ls_zstatus-option = 'EQ'.
      ls_zstatus-sign = 'E'.
      ls_zstatus-low = 'CLOSED'.
      APPEND ls_zstatus TO s_status.

      SELECT *
        FROM zbim_blk_invoice
        INTO TABLE t_alv
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND ebeln IN s_ebeln
          AND ( spgrp = p_spgrp OR
                spgrm =  p_spgrm OR
                zblck_razao = p_razao )
          AND wi_cd IN s_wi_cd
          AND wi_stat IN s_stat
          AND wi_id IN s_wiid
          AND lifnr IN s_lifnr
          AND zmblnr_inv IN s_docinv
          AND zmjahr_inv IN s_invdat
          AND zmblnr_em IN s_docem
          AND zmjahr_em IN s_emdat
          AND zbelnr_nc IN s_doc_nc
          AND zgjahr_nc IN s_ncdate
          AND zstatus IN s_status
          AND zeop_date IN s_zeop
          AND zdoc_estorno IN s_docest
          AND wi_cruser LIKE lv_uname
          AND zdebit_note IN s_nc_nd.

      REFRESH s_status.
      "Todos os workitems na caixa do utilizador
    ELSEIF p_mywid IS NOT INITIAL.
      ls_zstatus-option = 'EQ'.
      ls_zstatus-sign = 'E'.
      ls_zstatus-low = 'CLOSED'.
      APPEND ls_zstatus TO s_status.

      SELECT *
        FROM zbim_blk_invoice
        INTO TABLE t_alv
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND ebeln IN s_ebeln
          AND ( spgrp = p_spgrp OR
                spgrm =  p_spgrm OR
                zblck_razao = p_razao )
          AND wi_cd IN s_wi_cd
          AND wi_stat IN s_stat
          AND wi_id IN s_wiid
          AND lifnr IN s_lifnr
          AND zmblnr_inv IN s_docinv
          AND zmjahr_inv IN s_invdat
          AND zmblnr_em IN s_docem
          AND zmjahr_em IN s_emdat
          AND zbelnr_nc IN s_doc_nc
          AND zgjahr_nc IN s_ncdate
          AND zstatus IN s_status
          AND zeop_date IN s_zeop
          AND zdoc_estorno IN s_docest
          AND wi_cruser LIKE sy-uname
          AND zdebit_note IN s_nc_nd..

      REFRESH s_status.
      "Todos os workitems do utilizador
    ELSEIF p_all IS NOT INITIAL.
      SELECT *
        FROM zbim_blk_invoice
        INTO TABLE t_alv
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND ebeln IN s_ebeln
          AND ( spgrp = p_spgrp OR
                spgrm =  p_spgrm OR
                zblck_razao = p_razao )
          AND wi_cd IN s_wi_cd
          AND wi_stat IN s_stat
          AND wi_id IN s_wiid
          AND lifnr IN s_lifnr
          AND zmblnr_inv IN s_docinv
          AND zmjahr_inv IN s_invdat
          AND zmblnr_em IN s_docem
          AND zmjahr_em IN s_emdat
          AND zbelnr_nc IN s_doc_nc
          AND zgjahr_nc IN s_ncdate
          AND zstatus IN s_status
          AND zeop_date IN s_zeop
          AND zdoc_estorno IN s_docest
          AND wi_cruser LIKE sy-uname
          AND zdebit_note IN s_nc_nd.
    ELSEIF p_allwi IS NOT INITIAL.
      SELECT *
        FROM zbim_blk_invoice
        INTO TABLE t_alv
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND ebeln IN s_ebeln
          AND ( spgrp = p_spgrp OR
                spgrm =  p_spgrm OR
                zblck_razao = p_razao )
          AND wi_cd IN s_wi_cd
          AND wi_stat IN s_stat
          AND wi_id IN s_wiid
          AND lifnr IN s_lifnr
          AND wi_cruser IN s_wiuser
          AND zmblnr_inv IN s_docinv
          AND zmjahr_inv IN s_invdat
          AND zmblnr_em IN s_docem
          AND zmjahr_em IN s_emdat
          AND zbelnr_nc IN s_doc_nc
          AND zgjahr_nc IN s_ncdate
          AND zstatus IN s_status
          AND zeop_date IN s_zeop
          AND zdoc_estorno IN s_docest
          AND zdebit_note IN s_nc_nd.
    ENDIF.

    "Documentos de FI
  ELSEIF p_fi IS NOT INITIAL.
    "Todos os workitems possíveis para o utilizador
    IF p_mywidt IS NOT INITIAL.
      ls_stat-option = 'EQ'.
      ls_stat-sign = 'I'.
      ls_stat-low = TEXT-011.
      APPEND ls_stat TO s_stat.

      ls_zstatus-option = 'EQ'.
      ls_zstatus-sign = 'E'.
      ls_zstatus-low = 'CLOSED'.
      APPEND ls_zstatus TO s_status.

      SELECT *
        FROM zbim_blk_inv_fi
        INTO TABLE t_alv_fi
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND wi_cd IN s_wi_cd
          AND wi_id IN s_wiid
          AND wi_stat IN s_stat
          AND zstatus IN s_status
          AND zeop_date IN s_zeop
          AND wi_cruser LIKE lv_uname.
      "Todos os workitems na caixa do utilizador
    ELSEIF p_mywid IS NOT INITIAL.
      ls_zstatus-option = 'EQ'.
      ls_zstatus-sign = 'E'.
      ls_zstatus-low = 'CLOSED'.
      APPEND ls_zstatus TO s_status.

      SELECT *
      FROM zbim_blk_inv_fi
      INTO TABLE t_alv_fi
      WHERE bukrs IN s_bukrs
        AND belnr IN s_belnr
        AND gjahr IN s_gjahr
        AND wi_cd IN s_wi_cd
        AND wi_id IN s_wiid
        AND lifnr IN s_lifnr
        AND wi_stat IN s_stat
        AND zstatus IN s_status
        AND zeop_date IN s_zeop
        AND wi_cruser =  sy-uname.
      "Todos os workitems para o utilizador
    ELSEIF p_all IS NOT INITIAL.
      SELECT *
      FROM zbim_blk_inv_fi
      INTO TABLE t_alv_fi
      WHERE bukrs IN s_bukrs
        AND belnr IN s_belnr
        AND gjahr IN s_gjahr
        AND wi_cd IN s_wi_cd
        AND wi_id IN s_wiid
        AND lifnr IN s_lifnr
        AND wi_stat IN s_stat
        AND zstatus IN s_status
        AND zeop_date IN s_zeop
        AND wi_cruser LIKE sy-uname.
      "todos os workitems
    ELSEIF p_allwi IS NOT INITIAL.
      SELECT *
      FROM zbim_blk_inv_fi
      INTO TABLE t_alv_fi
        WHERE bukrs IN s_bukrs
          AND belnr IN s_belnr
          AND gjahr IN s_gjahr
          AND wi_cd IN s_wi_cd
          AND wi_id IN s_wiid
          AND lifnr IN s_lifnr
          AND wi_stat IN s_stat
          AND wi_cruser IN s_wiuser
          AND zstatus IN s_status
          AND zeop_date IN s_zeop.
    ENDIF.
  ENDIF.

  "Converter para decisão de utilizador extenso
  CALL FUNCTION 'DD_DOMA_GET'
    EXPORTING
      domain_name = 'Y_USER_DECISION_V'
    TABLES
      dd07v_tab_a = lt_taba
      dd07v_tab_n = lt_tabb.

  lv_objtype = 'ZCL_WF_BIM_REQ'.
  lv_event = 'NO_DIV'.

  "Documentos logistica
  IF p_log IS NOT INITIAL.
    LOOP AT t_alv ASSIGNING FIELD-SYMBOL(<fs>).
      READ TABLE lt_taba INTO ls_taba WITH KEY domvalue_l = <fs>-zuser_decision.

      IF sy-subrc = 0.
        <fs>-zuser_decision = ls_taba-ddtext.
      ENDIF.

      READ TABLE lt_t001 INTO ls_t001 WITH KEY bukrs = <fs>-bukrs.
      IF sy-subrc <> 0.
        SELECT *
          FROM t001
          APPENDING TABLE lt_t001
          WHERE bukrs = <fs>-bukrs.

        READ TABLE lt_t001 INTO ls_t001 WITH KEY bukrs = <fs>-bukrs.
      ENDIF.

      REFRESH: lt_xekbe[],
               lt_xekbes[].

      CALL FUNCTION 'ME_READ_HISTORY'
        EXPORTING
          ebeln  = <fs>-ebeln
          ebelp  = <fs>-ebelp
          webre  = 'X'
        TABLES
          xekbe  = lt_xekbe
          xekbes = lt_xekbes.

*** INETUM:SFSS:FI Beg S4DK926552
*      lv_buzei = <fs>-buzei.
*
*      SELECT a~belnr
*           , a~gjahr
*           , a~buzei
*           , b~bpmng
*           , c~webre
*        FROM rseg AS a
*        INNER JOIN @lt_xekbe AS b ON ( a~lfbnr EQ b~belnr AND a~lfgja EQ b~gjahr AND a~lfpos EQ b~buzei )
*        INNER JOIN ekpo AS c ON ( b~ebeln EQ c~ebeln AND b~ebelp EQ c~ebelp )
*        WHERE a~belnr EQ @<fs>-belnr
*          AND a~gjahr EQ @<fs>-gjahr
*          AND a~buzei EQ @lv_buzei
*        INTO TABLE @DATA(lt_rseg).

      READ TABLE lt_xekbes INTO ls_xekbes WITH KEY ebelp = <fs>-ebelp.
      IF sy-subrc = 0.
*        READ TABLE lt_rseg ASSIGNING FIELD-SYMBOL(<fs_rseg>) WITH KEY belnr = <fs>-belnr
*                                                                      gjahr = <fs>-gjahr
*                                                                      buzei = <fs>-buzei.
*        IF sy-subrc IS INITIAL AND <fs_rseg> IS ASSIGNED AND <fs_rseg>-webre IS NOT INITIAL.
*          <fs>-wemng = <fs_rseg>-bpmng.
*          <fs>-zdif_em_fat = <fs_rseg>-bpmng - <fs>-ztot_remng.
*        ENDIF.
*** INETUM:SFSS:FI End S4DK926552

        <fs>-wemng = ls_xekbes-bpmng. "ls_xekbes-wemng.
        <fs>-zdif_em_fat = ls_xekbes-bpmng - <fs>-ztot_remng."<fs>-wemng
        "Bloqueio de quantidade->atualiza div.em-fat
        IF <fs>-spgrm IS NOT INITIAL AND ( <fs>-wemng >= <fs>-ztot_remng ) AND <fs>-zstatus <> 'CLOSED'.
          CONCATENATE <fs>-bukrs <fs>-belnr <fs>-gjahr 'QTD' INTO lv_objkey.

          <fs>-zstatus = 'CLOSED'.
          CALL FUNCTION 'ZBIM_CLOSE_WORKITEM'
            EXPORTING
              i_objtype    = lv_objtype
              i_event      = lv_event
              i_objkey     = lv_objkey
              i_param_name = 'BUZEI'
              i_buzei      = <fs>-buzei
              i_wiid       = <fs>-wi_id
              i_status     = <fs>-zstatus
            IMPORTING
              o_wiid       = <fs>-wi_id
              o_wistat     = <fs>-wi_stat
              o_wicruser   = lv_agent.

          <fs>-wi_cruser = lv_agent.
          <fs>-zdiv_closed = 'X'.
          <fs>-zeop_date = sy-datum.
        ENDIF.

        "Bloqueio de preço
        IF <fs>-spgrp IS NOT INITIAL.
          "Moeda da empresa igual à da PO
          IF <fs>-waers = ls_t001-waers.
            <fs>-zwrbtr_dif = ls_xekbes-reewr - ls_xekbes-wewrt.

*CCF 24.10.2021
* Fatura ativos, o valor da ND terá de ser comparado com o pedido e não com a receção
*            IF teste_sbx = 'X'. "para testes apenas se funcionar retirar este if
            IF <fs>-zwrbtr_dif IS INITIAL AND <fs>-zdif_wrbtr IS NOT INITIAL.
              <fs>-zwrbtr_dif =  <fs>-zdif_wrbtr * <fs>-remng.
            ENDIF.
*            ENDIF.
* FI CCF

          ENDIF.
        ENDIF.
      ENDIF.
      <fs>-zexecute = icon_execute_object.

* CCF 24.10.2021 - atualiza status
* A Susana pediu para nao alterar os que tem 'ND' ou 'SNC'

      DATA: zlspr          LIKE zbim_blk_invoice-zlspr.
      CONCATENATE <fs>-belnr <fs>-bukrs  <fs>-gjahr INTO DATA(lv_awkey).
      SELECT SINGLE b~zlspr
         FROM bkpf AS a INNER JOIN bseg AS b ON
         a~bukrs = b~bukrs AND
         a~belnr = b~belnr AND
         a~gjahr = b~gjahr
         INTO zlspr
         WHERE a~awkey = lv_awkey
         AND b~koart = 'K'.
      IF sy-subrc = 0 AND ( zlspr = ' ' AND <fs>-zlspr <> ' ').
        <fs>-zlspr = zlspr.
        IF <fs>-zdebit_note NE 'SNC'.
          <fs>-zstatus = 'CLOSED'." a susana pediu para nao atulizar para cosed
          "pois podem ter desbloqueado a fatura mas querem gerar a ND
        ENDIF.
*        if zlspr is NOT initial.
*          clear <fs>-zstatus.
*        endif.
      ENDIF.

* Atualiza status WI
*      ls_workitem-wi_id = <fs>-wi_id.
*
*      CALL FUNCTION 'ZBIM_GET_WIID'
*        EXPORTING
*          i_workitem = ls_workitem
*          i_status   = <fs>-zstatus
*        IMPORTING
*          o_wi_id    = <fs>-wi_id
*          o_status   = <fs>-wi_stat
*          o_agent    = lv_agent.
*      IF lv_agent IS INITIAL.
*        SELECT * FROM
*          swwuserwi
*          INTO TABLE @DATA(lt_userwi)
*          WHERE wi_id = @<fs>-wi_id.
*
*        IF sy-subrc = 0.
*          CLEAR <fs>-wi_cruser.
*          DESCRIBE TABLE lt_userwi LINES DATA(lv_lines).
*          LOOP AT lt_userwi INTO DATA(ls_userwi).
*            IF lv_lines > 1.
*              CONCATENATE ls_userwi-user_id <fs>-wi_cruser INTO <fs>-wi_cruser SEPARATED BY '|'.
*            ELSE.
*              <fs>-wi_cruser = ls_userwi-user_id.
*            ENDIF.
*          ENDLOOP.
*        ELSE.
*          <fs>-wi_cruser = lv_agent.
*        ENDIF.
*      ENDIF.
* FIM CCF

      IF ( <fs>-zintern_doc IS NOT INITIAL
        OR <fs>-zbelnr_nc IS NOT INITIAL ) AND <fs>-zlspr  IS NOT INITIAL.
        <fs>-zstatus = 'CLOSED'.
      ENDIF.

      MOVE-CORRESPONDING <fs> TO ls_log_data.
      APPEND ls_log_data TO gt_log_data.
    ENDLOOP.


    MODIFY zbim_blk_invoice FROM TABLE t_alv.

    "Documentos financeiros
  ELSEIF p_fi IS NOT INITIAL.
    LOOP AT t_alv_fi ASSIGNING FIELD-SYMBOL(<fs1>).
      READ TABLE lt_taba INTO ls_taba WITH KEY domvalue_l = <fs1>-zuser_decision.

      IF sy-subrc = 0.
        <fs1>-zuser_decision = ls_taba-ddtext.
      ENDIF.
      <fs1>-zexecute = icon_execute_object.

*22.11.20220CCF atualiza bloqueio de pagamento dos doc de fi
      CONCATENATE <fs1>-belnr <fs1>-bukrs <fs1>-gjahr INTO lv_awkey.

      SELECT SINGLE b~zlspr
        FROM bkpf AS a INNER JOIN bseg AS b ON
          a~bukrs = b~bukrs AND
          a~belnr = b~belnr AND
          a~gjahr = b~gjahr
        INTO <fs1>-zlspr
        WHERE a~awkey = lv_awkey
          AND b~koart = 'K'.
* fim CCf

      MOVE-CORRESPONDING <fs1> TO ls_fi_data.
      APPEND ls_fi_data TO gt_fi_data.
    ENDLOOP.
  ENDIF.

*  SORT t_alv BY bukrs belnr gjahr buzei ASCENDING.
ENDFORM.
FORM f_display_alv.
  CALL SCREEN 100.
ENDFORM.
FORM f_init_fieldcat.

  DATA: ls_fieldcat TYPE lvc_s_fcat.

  IF p_log IS NOT INITIAL OR p_fi IS NOT INITIAL.
    ls_fieldcat-fieldname = 'BELNR'.
    ls_fieldcat-hotspot = 'X'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ZEXECUTE'.
    APPEND ls_fieldcat TO gt_fieldcat.
  ENDIF.
  IF p_log IS NOT INITIAL.
    ls_fieldcat-fieldname = 'EBELN'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ZMBLNR_INV'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ZMBLNR_EM'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ZBELNR_NC'.
    APPEND ls_fieldcat TO gt_fieldcat.
*>>>NBA -INI  01.11.2021
    ls_fieldcat-fieldname = 'ICON_LOG_NC'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ICON_LOG_MM'.
    APPEND ls_fieldcat TO gt_fieldcat.
    ls_fieldcat-fieldname = 'ICON_LOG_DESB'.
    APPEND ls_fieldcat TO gt_fieldcat.
*<<<NBA-FIM



  ENDIF.

ENDFORM.
FORM f_init_style.

  DATA: ls_celltab TYPE lvc_s_styl.

  "COMENTADO PARA POSSIBILITAR INTRODUÇÃO DE DOCUMENTO INTERNO
  IF p_log IS NOT INITIAL.
    LOOP AT gt_log_data ASSIGNING FIELD-SYMBOL(<fs>).
*      IF <fs>-zdebit_note = 'SNC' OR <fs>-zdebit_note = 'ND'. "AND <fs>-zintern_doc IS INITIAL.
      ls_celltab-fieldname = 'ZINTERN_DOC'.
      ls_celltab-style = cl_gui_alv_grid=>mc_style_enabled.
      INSERT ls_celltab INTO TABLE <fs>-celltab.
*      ENDIF.
    ENDLOOP.
  ELSE.
    LOOP AT gt_fi_data ASSIGNING FIELD-SYMBOL(<fs1>).
      ls_celltab-fieldname = 'ZINTERN_DOC'.
      ls_celltab-style = cl_gui_alv_grid=>mc_style_enabled.
      INSERT ls_celltab INTO TABLE <fs1>-celltab.
    ENDLOOP.
  ENDIF.
ENDFORM.
FORM f_save_data USING p_value TYPE c.

  DATA: l_valid     TYPE c,
        lv_answer   TYPE c,
        lt_log_data TYPE STANDARD TABLE OF ty_outtab_log,
        lt_fi_data  TYPE STANDARD TABLE OF ty_outtab_fi,
        ls_aux_data LIKE zbim_blk_invoice.

  CLEAR l_valid.

*  CALL METHOD g_grid->refresh_table_display.

  lt_log_data[] = gt_log_data[].
  lt_fi_data[] = gt_fi_data[].

  IF g_grid IS NOT INITIAL.
    CALL METHOD g_grid->check_changed_data
      IMPORTING
        e_valid = l_valid.
  ENDIF.
  "Saiu sem gravar
  IF p_value = 'V'.
    IF l_valid IS NOT INITIAL.
      PERFORM f_update_table USING lt_log_data
                                   lt_fi_data
                                   'V'.
    ENDIF.
  ELSE.
    "Clicou em salvar
    IF l_valid IS NOT INITIAL.
      PERFORM f_update_table USING lt_log_data
                                   lt_fi_data
                                   'S'.
    ENDIF.
  ENDIF.

ENDFORM.
FORM f_update_table USING t_log_data LIKE gt_log_data
                          t_fi_data LIKE gt_fi_data
                          l_action TYPE c.

  DATA: lt_log_data TYPE STANDARD TABLE OF zbim_blk_invoice,
        lt_fi_data  TYPE STANDARD TABLE OF zbim_blk_inv_fi,
        lv_answer   TYPE c,
        ls_aux_data LIKE zbim_blk_invoice,
        ls_aux_fi   LIKE zbim_blk_inv_fi.

  IF p_log IS NOT INITIAL.
    LOOP AT gt_log_data ASSIGNING FIELD-SYMBOL(<fs>).
      READ TABLE t_log_data ASSIGNING FIELD-SYMBOL(<fs1>) WITH KEY bukrs = <fs>-bukrs
                                                                   belnr = <fs>-belnr
                                                                   buzei = <fs>-buzei.

      IF sy-subrc = 0.
        IF <fs1> <> <fs>.
          DATA(lv_change) = abap_true.
          MOVE-CORRESPONDING <fs> TO ls_aux_data.
          APPEND ls_aux_data TO lt_log_data.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ELSE.
    LOOP AT gt_fi_data ASSIGNING FIELD-SYMBOL(<fs2>).
      READ TABLE t_fi_data ASSIGNING FIELD-SYMBOL(<fs3>) WITH KEY bukrs = <fs2>-bukrs
                                                                  belnr = <fs2>-belnr
                                                                  buzei = <fs2>-buzei.
      IF sy-subrc = 0.
        IF <fs2> <> <fs3>.
          lv_change = abap_true.
          MOVE-CORRESPONDING <fs2> TO ls_aux_fi.
          APPEND ls_aux_fi TO lt_fi_data.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDIF.

  IF lv_change = abap_true.
    IF l_action = 'S'.
      IF p_log IS NOT INITIAL.
        MODIFY zbim_blk_invoice FROM TABLE lt_log_data.
      ELSE.
        MODIFY zbim_blk_inv_fi FROM TABLE lt_fi_data.
      ENDIF.
    ELSEIF l_action = 'V'.
      CALL FUNCTION 'POPUP_TO_CONFIRM_STEP_2_BUTTON'
        EXPORTING
          textline1 = TEXT-013
          titel     = TEXT-014
        IMPORTING
          answer    = lv_answer.

      IF lv_answer <> 'A'.
        IF p_log IS NOT INITIAL.
          MODIFY zbim_blk_invoice FROM TABLE lt_log_data.
        ELSE.
          MODIFY zbim_blk_inv_fi FROM TABLE lt_fi_data.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form exibe_log
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*&      --> LS_ALV_BALOGNR
*&---------------------------------------------------------------------*
FORM exibe_log  USING    p_balognr.




  DATA: lit_msg_handle      TYPE  bal_t_msgh,
        lwa_display_profile TYPE bal_s_prof,
        lwa_log_filter      TYPE bal_s_lfil,
        lit_log_header      TYPE balhdr_t,
        lit_log_handle      TYPE bal_t_logh,
        lwa_log_header      TYPE balhdr,
        lwa_lognumber       TYPE LINE OF bal_r_logn,
        lwa_object          TYPE LINE OF bal_r_obj,
        lwa_subobject       TYPE LINE OF bal_r_sub.

  CALL FUNCTION 'BAL_GLB_MEMORY_REFRESH'
    EXPORTING
      i_refresh_all = 'X'.

  "Pesquisa log
  lwa_object-sign = 'I'.
  lwa_object-option = 'EQ'.
  lwa_object-low = 'ZBIM'.
  APPEND lwa_object TO lwa_log_filter-object.

  lwa_subobject-sign = 'I'.
  lwa_subobject-option = 'EQ'.
  lwa_subobject-low = 'ZBIM00'.
  APPEND lwa_subobject TO lwa_log_filter-subobject.


  lwa_lognumber-sign = 'I'.
  lwa_lognumber-option = 'EQ'.
  lwa_lognumber-low = p_balognr.
  APPEND lwa_lognumber TO lwa_log_filter-lognumber.


  CALL FUNCTION 'BAL_DB_SEARCH'
    EXPORTING
      i_s_log_filter     = lwa_log_filter
    IMPORTING
      e_t_log_header     = lit_log_header
    EXCEPTIONS
      log_not_found      = 1
      no_filter_criteria = 2
      OTHERS             = 3.
  IF sy-subrc <> 0.
    MESSAGE 'Log não encontrado' TYPE 'S' DISPLAY LIKE 'E'.
  ENDIF.


  CALL FUNCTION 'BAL_DB_LOAD'
    EXPORTING
      i_t_log_header     = lit_log_header
    IMPORTING
      e_t_log_handle     = lit_log_handle
      e_t_msg_handle     = lit_msg_handle
    EXCEPTIONS
      no_logs_specified  = 1
      log_not_found      = 2
      log_already_loaded = 3
      OTHERS             = 4.
  IF sy-subrc <> 0.
    MESSAGE 'Erro na exibição de log.' TYPE 'E'.
  ENDIF.


  CALL FUNCTION 'BAL_DSP_LOG_DISPLAY'
    EXPORTING
*     i_s_display_profile = lwa_display_profile
      i_t_log_handle = lit_log_handle
      i_t_msg_handle = lit_msg_handle
    EXCEPTIONS
      OTHERS         = 1.
  IF sy-subrc NE 0.
    MESSAGE 'Erro na exibição de log.' TYPE 'E'.
  ENDIF.







ENDFORM.