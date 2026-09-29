FUNCTION /sbxc/zckp_valida_processo.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(PROCESSO) TYPE  /SBXC/ZCKP_PROCESSO
*"  EXPORTING
*"     REFERENCE(STR_PROC) LIKE  /SBXC/ZCKP_TAB00 STRUCTURE
*"       /SBXC/ZCKP_TAB00
*"  EXCEPTIONS
*"      SEM_AUTORIZACAO
*"      NAO_EXISTE
*"----------------------------------------------------------------------
*  SELECT SINGLE * FROM /sbxc/zckp_tab00
*   INTO CORRESPONDING FIELDS OF str_proc
*          WHERE processo = processo.
*
*  IF sy-subrc NE 0.
*    RAISE nao_existe.
*  ELSE.

    AUTHORITY-CHECK OBJECT 'ZCKP:PROCS'
             ID '/SBXC/CPRO' FIELD processo
             ID '/SBXC/CAUT' DUMMY.

    IF sy-subrc NE 0.
      RAISE sem_autorizacao.
    ENDIF.

*  ENDIF.

ENDFUNCTION.
