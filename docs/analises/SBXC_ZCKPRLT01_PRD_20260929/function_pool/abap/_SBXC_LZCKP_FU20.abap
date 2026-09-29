FUNCTION /sbxc/zckp_mm_cria_adiant.
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
* Criado por Cláudia Fernandes @ SBXC
* Função para registar dados dos adiantamentos na tabela de suporte ao
* cockpit
************************************************************************

  MOVE-CORRESPONDING headerdata TO /sbxc/zckp_adnt.

  IF headerdata-bline_date IS INITIAL.
    headerdata-bline_date = headerdata-pstng_date.
  ENDIF.

  IF headerdata-username IS INITIAL.
    /sbxc/zckp_adnt-username = sy-uname.
  ENDIF.
  IF headerdata-processo IS INITIAL.
    /sbxc/zckp_adnt-processo = 'ADNT_GRP1'.
  ENDIF.

  CALL FUNCTION 'NUMBER_GET_NEXT'
     EXPORTING
       nr_range_nr                   = '01'
       object                        = 'ZCKP_COCKP'
      quantity                      = '1'
*   SUBOBJECT                     = ' '
      toyear                        = sy-datum(4)
*   IGNORE_BUFFER                 = ' '
    IMPORTING
      number                        =  /sbxc/zckp_adnt-seqno
*   QUANTITY                      =
*   RETURNCODE                    =
 EXCEPTIONS
   INTERVAL_NOT_FOUND            = 1
   NUMBER_RANGE_NOT_INTERN       = 2
   OBJECT_NOT_FOUND              = 3
   QUANTITY_IS_0                 = 4
   OTHERS                        = 5
             .
  IF sy-subrc <> 0.
 MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
  ENDIF.
  INSERT  /sbxc/zckp_adnt.

  MOVE-CORRESPONDING  /sbxc/zckp_adnt TO /sbxc/zckp_ctrl.
  /sbxc/zckp_ctrl-data_in = sy-datum.
  /sbxc/zckp_ctrl-hora_in = sy-uzeit.
  /sbxc/zckp_ctrl-user_in = sy-uname.
  /sbxc/zckp_ctrl-status1 = '0'.

  INSERT  /sbxc/zckp_ctrl.
  COMMIT WORK AND WAIT.

  return-message = text-031. "'Documento registado para posterior processamento em SAP'.
  return-type = 'S'.
  APPEND return.

ENDFUNCTION.
