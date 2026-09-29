FUNCTION /SBXC/ZCKP_BLOQ_FATURA1 .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(LISTA_FACTURAS) TYPE  /SBXC/ZCKP_BLOQ_FATURA
*"     VALUE(UUID) TYPE  /SBXC/ZUUIDO OPTIONAL
*"  EXPORTING
*"     VALUE(RETURN) TYPE  CHAR1
*"     VALUE(DETERRO) TYPE  CHAR50
*"     VALUE(ORIG_UUID) TYPE  /SBXC/ZUUIDO
*"----------------------------------------------------------------------

*"----------------------------------------------------------------------
* PROJECTO COCKPIT
* Criado por Cláudia Fernandes @ SBXC
* Função para alterar motivo de bloqueio de facturas enviadas pela origem
************************************************************************
  uuid_orig = uuid.  "guarda origem para registo LOG
  orig_uuid = uuid.  "guarda origem para devolver

* Preenche estruturas BAPI de facturas
  PERFORM preenche_estrut_func
             USING lista_facturas.


*"----------------------------------------------------------------------
* chamada à função de desbloqueio
*"----------------------------------------------------------------------
  MOVE lista_facturas-fisc_year TO ld_aworg.
  UNPACK  lista_facturas-lifnr TO  lista_facturas-lifnr.

  IF lista_facturas-inv_doc_no IS INITIAL.
    return = '2'.
    CONCATENATE text-037 lista_facturas-inv_doc_no
                          lista_facturas-fisc_year
                          lista_facturas-comp_code
                         text-038
    INTO deterro SEPARATED BY space.
  ELSE.

    CALL FUNCTION 'ENQUEUE_EFBKPF'
      EXPORTING
        belnr          = lista_facturas-inv_doc_no
        bukrs          = lista_facturas-comp_code
        gjahr          = lista_facturas-fisc_year
      EXCEPTIONS
        foreign_lock   = 1
        system_failure = 2.
    IF sy-subrc EQ 0.
      CALL FUNCTION 'DEQUEUE_EFBKPF'
        EXPORTING
          belnr = lista_facturas-inv_doc_no
          bukrs = lista_facturas-comp_code
          gjahr = lista_facturas-fisc_year.
    ELSE.
      deterro = text-039. "'ERRO DocFI bloqueado por outro processo'.
      return = '1'.
    ENDIF.

    CHECK return IS INITIAL.

    CALL FUNCTION 'FI_DOCUMENT_CHANGE'
      EXPORTING
        i_awtyp              = 'RMRP'
*        i_awref              =  ls_accdn-awref
        i_awref              = lista_facturas-inv_doc_no
        i_aworg              = ld_aworg
        i_lifnr              = lista_facturas-lifnr
      TABLES
        t_accchg             = lt_accchg
      EXCEPTIONS
        no_reference         = 1
        no_document          = 2
        many_documents       = 3
        wrong_input          = 4
        overwrite_creditcard = 5
        OTHERS               = 6.


    return = sy-subrc.

    CASE return.
      WHEN 0.
        deterro = text-040. "'OK - DocFi modificado'. "desbloqueado'.
      WHEN 1.
        deterro = text-041. "'ERRO- no reference'.
        return = '5'.
      WHEN 2.
        deterro = text-042. "'ERRO- no document'.
        return = '5'.
      WHEN 3.
        deterro = text-043. "'ERRO- many documents'.
        return = '5'.
      WHEN 4.
        deterro = text-044. "'ERRO- wrong_input'.
        return = '5'.
      WHEN 5.
        deterro = text-045. "'ERRO- overwrite_creditcard'.
        return = '5'.
      WHEN 6.
        deterro = text-046. "'ERRO- Others '.
        return = '5'.
    ENDCASE.
  ENDIF.


* obter chave do request / response para LOG
  ASSIGN  lista_facturas TO <fs> CASTING TYPE c.
  chave = <fs>.

  processo = '/SBXC/ZCKP_BLOQ_FATURA'.
  inform   = deterro.

  DATA: sap_uuid  TYPE  /SBXC/ZUUID. "/sbxc/zuuido.
CALL FUNCTION '/SBXC/ZCKP_REGISTA_LOG'
  EXPORTING
    processo         = processo
    chave            = chave
    utilizador       = sy-uname
    inform           = inform
    codretorno       = return
    UUID             = uuid_orig
    ISAP_UUID        = 'NA'
 IMPORTING
   SAP_UUID         = SAP_UUID.

ENDFUNCTION.
