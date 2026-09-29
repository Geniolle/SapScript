FUNCTION /sbxc/zckp_baseline_date.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(PROCESSO) TYPE  /SBXC/ZCKP_INVH-PROCESSO
*"     REFERENCE(SEQNO) TYPE  /SBXC/ZCKP_INVH-SEQNO
*"  TABLES
*"      ITEMDATA STRUCTURE  /SBXC/ZCKP_INVI OPTIONAL
*"  CHANGING
*"     REFERENCE(BASELINE_DATE) TYPE  DATS
*"----------------------------------------------------------------------
**-------------------------------------------------------------------*
**  MODIFICAÇÕES
*---------------------------------------------------------------------*
*& Autor        : A. Garrido (SBX)
*& Data         : 24.07.2024 10:06:57
*& Referência   : 2500005
*& Transporte   : S4DK950559
*& Ch. pesquisa : AFG-ID-2500005
*& Objetivo     : Data Base é sempre igual à data documento
*
*&---------------------------------------------------------------------*
  DATA: lt_xekbe  TYPE TABLE OF ekbe,
        lt_xekbes TYPE TABLE OF ekbes.
  DATA: lv_bldat TYPE sy-datum.

  DATA: headerdata LIKE /sbxc/zckp_invh.
*  MOVE-CORRESPONDING cab TO headerdata.

* Se o item foi transferido de outro processo não deve recalcular
* a data base de pagamento
  SELECT SINGLE status1 AS status1, processo_ant AS processo_ant INTO  @DATA(zckp_ctrl)
    FROM /sbxc/zckp_ctrl
    WHERE
    processo =     @processo AND
    seqno =  @seqno.

  SELECT SINGLE * FROM /sbxc/zckp_invh INTO headerdata
WHERE processo =     processo AND
  seqno =  seqno.

**AFG-ID-2500005****
baseline_date = headerdata-doc_date.

*  IF ( zckp_ctrl-processo_ant IS INITIAL AND ( zckp_ctrl-status1 = ' ' OR zckp_ctrl-status1 = '0' ) )
*    OR headerdata-baseline_date IS INITIAL.
*
*    "determinar incoterms de excepção
*    SELECT * INTO TABLE @DATA(lt_inc)
*      FROM /sbxc/zckp_inc.
*
** Determina org. compras CORE
*    SELECT * INTO TABLE @DATA(lt_core)
*      FROM /sbxc/zckp_tab32.
*
*    "determinar incoterm do pedido
** CCF nesta altura ainda não temos os dados de item
** por isso temos de fazer a seleção à tabela de itens
*    SELECT SINGLE po_number INTO  @DATA(lv_po_number)
*      FROM /sbxc/zckp_invi
*      WHERE
*      processo =     @processo AND
*      seqno =  @seqno AND
*      po_number <> ' '.
*
**    SELECT SINGLE * FROM /sbxc/zckp_invh INTO headerdata
**      WHERE processo =     processo AND
**        seqno =  seqno.
*
*    IF lv_po_number IS NOT INITIAL.
*      SELECT SINGLE inco1 AS inc, ekorg AS org, lifnr AS vendor INTO @DATA(ls_struct)
*        FROM ekko
*        WHERE ebeln EQ  @lv_po_number.
**      WHERE ebeln EQ @<fs1>-po_number.
*
*      "Determinar itens do pedido
*      SELECT * INTO TABLE @DATA(lt_ekpo)
*        FROM ekpo
*        WHERE ebeln EQ  @lv_po_number.
*
*
*      " le os dados do 1º item do pedido de compra
*      READ TABLE lt_ekpo  ASSIGNING FIELD-SYMBOL(<fs3>) INDEX 1.
*
*
*      READ TABLE lt_core ASSIGNING FIELD-SYMBOL(<fs5>) WITH KEY ekorg = ls_struct-org
*      .
*      IF sy-subrc NE 0.
*        MESSAGE i085(/sbxc/zckp_cockpit) WITH ls_struct-org DISPLAY LIKE 'E'.
*      ELSE.
*        IF  <fs5>-naltdata = 'X'. "para a Salsa a data base não é manipulada
*          lv_bldat = headerdata-doc_date.
*        ELSE.
*          "empresas não Salsa
*          READ TABLE lt_inc ASSIGNING FIELD-SYMBOL(<fs2>) WITH KEY inco1 = ls_struct-inc.
*          IF sy-subrc EQ 0. "Se incoterm existe na tabela então determina o campo z do pedido de compra
*            lv_bldat = <fs3>-zzdat03.
*          ELSE.
*
*            IF <fs5>-core = ' '.
*              "Organização de compras NCORE e como tal é uma compra de serviços
*              "Se compra de serviços (não core), então a data base pagamento é a data de receção do documento fatura
*              lv_bldat = headerdata-data_in.
*            ELSE."Determina a maior data entre a data da fatura e a data de recepção de mercadoria
*
*              "Determinar no dado mestre de fornecedor a indicação de que a fatura é entregue com a mercadoria
*              "Campo LFM1-WEBRE
*
*              SELECT SINGLE webre INTO @DATA(lv_webre)
*                FROM lfm1
*                WHERE lifnr EQ @ls_struct-vendor
*                AND ekorg EQ @ls_struct-org.
*
*              "Data da fatura (headerdata-doc_date)
*              "Data da recepção da fatura (headerdata-data_in)
*              CLEAR lv_bldat.
*              "Verificar histórico do pedido
*              CALL FUNCTION 'ME_READ_HISTORY'
*                EXPORTING
*                  ebeln  = lv_po_number "<fs1>-po_number
*                  ebelp  = <fs3>-ebelp
*                  webre  = 'X'
*                TABLES
*                  xekbe  = lt_xekbe
*                  xekbes = lt_xekbes.
*              LOOP AT lt_xekbe ASSIGNING FIELD-SYMBOL(<fs4>) WHERE vgabe = '1' AND bwart = '101'. "Entradas de mercadoria
*                IF <fs4>-budat > lv_bldat.
*                  "IF <fs4>-bldat > lv_bldat.CCF
*                  "lv_bldat = <fs4>-bldat.CCF
*                  lv_bldat = <fs4>-budat.
*
*                ENDIF.
*              ENDLOOP.
*
*              "lv_bldat  --> Data da ultima receção
*              IF lv_webre EQ 'X'.
*                "Data base fica igual à data da fatura do fornecedor ou receção da mercadoria  (a ultima das quais
*                "se a data da fatura > data de receção
*                IF headerdata-doc_date > lv_bldat.
**          lv_bldat = headerdata-data_in.
*                  lv_bldat = headerdata-doc_date.
*                ELSE.
*                  lv_bldat = lv_bldat.
*                ENDIF.
*
*              ELSE.
*                "A data base de pagamento será a Data da receção da fatura ou data de receção da mercadoria /ultima das quais
*                IF headerdata-data_in > lv_bldat.
*                  lv_bldat = headerdata-data_in.
*                ELSE.
*                  lv_bldat = lv_bldat.
*                ENDIF.
*              ENDIF.
*
*
*            ENDIF.
*          ENDIF.
*        ENDIF.
*      ENDIF.
*
*    ENDIF.
*
*    baseline_date = lv_bldat.
*    IF  baseline_date IS INITIAL.
*      baseline_date = headerdata-data_in.
*    ENDIF.
*
*  ELSE.
*    baseline_date = headerdata-baseline_date. "Manter a data base que já estava calculada.
*  ENDIF.
**************************************



ENDFUNCTION.
