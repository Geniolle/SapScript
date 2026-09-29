FUNCTION /sbxc/zckp_mm_regista_adiant.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(HEADERDATA) TYPE  /SBXC/ZCKP_ADIANTAMENTO
*"  EXPORTING
*"     VALUE(INVOICEDOCNUMBER) LIKE  BAPI_INCINV_FLD-INV_DOC_NO
*"     VALUE(FISCALYEAR) LIKE  BAPI_INCINV_FLD-FISC_YEAR
*"  TABLES
*"      RETURN STRUCTURE  BAPIRET2
*"----------------------------------------------------------------------
* Função para criar adiantamentos vindos Saphety
************************************************************************

** Preenche estruturas BAPI de adiantamento
  CLEAR wa_header.

  PERFORM preenche_estrs_adiant
            USING headerdata.

**"---------------------------------------------------------------------
** BAPI
**"---------------------------------------------------------------------

  PERFORM lanca_doc_fi USING  headerdata-comp_code
                       headerdata-pstng_date  headerdata-doc_date
                       text-022 headerdata-username headerdata-processo
                        headerdata-ano headerdata-seqno.


  CALL FUNCTION 'C14Z_MESSAGES_SHOW_AS_POPUP'
    TABLES
      i_message_tab = msg_ckp.

  return[] = t_return[].

* apaga tabela de log
  PERFORM log_delete TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .
* Guarda msg na tabela de log
  PERFORM log TABLES return USING headerdata-processo headerdata-ano headerdata-seqno .


ENDFUNCTION.
