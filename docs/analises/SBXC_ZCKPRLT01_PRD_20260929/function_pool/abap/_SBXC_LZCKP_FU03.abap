FUNCTION /sbxc/zckp_cancel_doc_pre_edit.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(E_UCOMM)
*"     REFERENCE(IDX_LIN) OPTIONAL
*"  EXPORTING
*"     REFERENCE(REFRESH)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR OPTIONAL
*"  CHANGING
*"     REFERENCE(CAB)
*"     REFERENCE(CTRL) TYPE  /SBXC/ZCKP_CTRL OPTIONAL
*"----------------------------------------------------------------------

  refresh = 'X'.

  MOVE-CORRESPONDING cab TO header.

  PERFORM mir4_el_pre_editada USING header-doc_lo header-ano_lanc.

  LOOP AT messtab.

    IF messtab-msgtyp = 'S'.
      UPDATE /sbxc/zckp_invh SET
      doc_fi = ' '
      ano_lanc = ' '
      doc_lo = ' '
      doc_estorno = ' '
      WHERE processo = header-processo AND
      ano = header-ano AND
      seqno = header-seqno.

      UPDATE /sbxc/zckp_ctrl SET
           status1 = '0'
           data_chg_st1 = ' '
           hora_chg_st1 = ' '
           user_chg_st1 = ' '
           WHERE processo = header-processo AND
           ano = header-ano AND
           seqno = header-seqno.

      CLEAR: header-doc_estorno, header-doc_fi, header-ano_lanc, header-doc_lo .

      ADD 1 TO wa_msg-lineno.
      wa_msg-msgid = messtab-msgid.
      wa_msg-msgno = messtab-msgnr.
      wa_msg-msgty = messtab-msgtyp.

      wa_msg-msgv1 =  messtab-msgv1.
      wa_msg-msgv2 =  messtab-msgv2.
      wa_msg-msgv3 =  messtab-msgv3.
      wa_msg-msgv4 =  messtab-msgv4.
      APPEND wa_msg TO msg_ckp.
*        delete from zckp_inv_item where
*             processo = header-processo and
*             ano = header-ano and
*            seqno = header-seqno.
*
*
*      select * from zckp_hist_item into zckp_inv_item  where
*                    processo = header-processo and
*                   ano = header-ano and
*                   seqno = header-seqno.
*        insert zckp_inv_item.
*      endselect.
    ENDIF.
  ENDLOOP.

  COMMIT WORK AND WAIT.

  MOVE-CORRESPONDING header TO cab.


ENDFUNCTION.
