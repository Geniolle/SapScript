FUNCTION /SBXC/ZCKP_IMG.
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
  DATA: header TYPE TABLE OF /sbxc/zckp_invh WITH HEADER LINE,
        wa_header TYPE /sbxc/zckp_invh,
        item   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        item_f   TYPE TABLE OF /sbxc/zckp_invi   WITH HEADER LINE,
        l_returncode,
        lt_sval LIKE sval OCCURS 0 WITH HEADER LINE,
        lv_pro  TYPE  /sbxc/zckp_processo,
        lv_ano  TYPE  gjahr,
        lv_seq  TYPE  /sbxc/zckp_seqno,
        opcao TYPE i,
        est_mensagem TYPE /sbxc/zckp_tab06-est_mensagem,
        processo_pos TYPE /sbxc/zckp_processo,
        ano_pos TYPE gjahr,
        seqno_pos TYPE /sbxc/zckp_seqno,
        gt_outtab TYPE TABLE OF /sbxc/zckp_tab07 WITH HEADER LINE,
        gs_private TYPE slis_data_caller_exit,
        gs_selfield TYPE slis_selfield,
        g_exit(1) TYPE c.


data: ls_xml    type /sbxc/zsbx_st_img.
data: it_xml type table of /sbxc/zsbx_st_img with header line.


* Preencher estrutura cabecalho

  MOVE-CORRESPONDING cab TO header.

APPEND header.

*ls_xml-UUID     = header-UUID.
*ls_xml-docType = '2'.
    READ TABLE header INDEX 1.
 perform send_2_saphety using ls_xml  header-REF_DOC_NO_ORIG header-BARCODE.


ENDFUNCTION.
