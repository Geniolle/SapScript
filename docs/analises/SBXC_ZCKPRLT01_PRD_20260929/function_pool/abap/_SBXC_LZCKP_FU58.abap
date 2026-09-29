FUNCTION /SBXC/ZCKP_PO_CREATE .
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(POHEADER) TYPE  /SBXC/ZBAPIMEPOHEADER
*"     VALUE(TESTRUN) TYPE  BAPIFLAG-BAPIFLAG
*"     VALUE(BASE64_STRING) TYPE  SFHTTPTYPE
*"  EXPORTING
*"     VALUE(EXPPURCHASEORDER) LIKE  BAPIMEPOHEADER-PO_NUMBER
*"     VALUE(EXPHEADER) LIKE  BAPIMEPOHEADER STRUCTURE  BAPIMEPOHEADER
*"     VALUE(EXPPOEXPIMPHEADER) LIKE  BAPIEIKP STRUCTURE  BAPIEIKP
*"  TABLES
*"      RETURN STRUCTURE  BAPIRET2
*"      POITEM STRUCTURE  /SBXC/ZBAPIMEPOITEM
*"      POSCHEDULE STRUCTURE  /SBXC/ZBAPIMEPOSCHEDUL
*"      POACCOUNT STRUCTURE  /SBXC/ZBAPIMEPOACCOUNT
*"----------------------------------------------------------------------

  TYPES: BEGIN OF ty64,
      tipo(60),
      ficheiro(60),
      base64 TYPE string,
      END OF ty64.

  DATA: it64 TYPE TABLE OF ty64 WITH HEADER LINE.
  DATA: BEGIN OF it_base64 ,
          base64_imp LIKE TABLE OF it64,
       END OF it_base64.

  DATA: BEGIN OF valida_fich OCCURS 0,
    ficheiro(60),
    END OF valida_fich.



  DATA: xml_out TYPE  string.
  DATA: imagem_out TYPE  string.

  DATA: objecttype LIKE toa01-sap_object.
  DATA conn_tab   LIKE toav0 OCCURS 1 WITH HEADER LINE.
  DATA: objectid LIKE toa01-object_id.
  DATA:  ar_obj LIKE toadv-ar_object.
  DATA: toa  LIKE toadt.
  DATA: archivobject TYPE TABLE OF tbl1024 WITH HEADER LINE.
  DATA: len TYPE i.
  DATA:  l_xstr TYPE xstring.


* descodificar mensagem de imagens
  PERFORM decode_base64 USING base64_string
          CHANGING xml_out
                  l_xstr .

* Validar se mensagem ok
  LOOP AT it_base64-base64_imp INTO it64.


*valida tipo
    CASE it64-tipo.
      WHEN 'Photos'.
*        ar_obj = 'ZACIC_PO03'.
      WHEN 'Email'.
*        ar_obj = 'ZACIC_PO04'.
      WHEN 'Proposal Subject for Approval'.
*        ar_obj = 'ZACIC_PO01'.  " ZACIC_PO01 ou ZACIC_PO02 !!
      WHEN 'Other Proposal'.
*         ar_obj = 'ZACIC_PO02'.
      WHEN OTHERS.
        w_return-message_v1 = it64-tipo.
        PERFORM error_processing TABLES return
          USING 'E' 'ZCKP' '000' w_return-message_v1 '' '' ''.
        "Tipo de objecto de imagem & inválido.

* valida ficheiro
        REFRESH valida_fich.
        SPLIT it64-ficheiro AT '.' INTO TABLE valida_fich.
        DESCRIBE TABLE valida_fich LINES sy-tfill.
        READ TABLE valida_fich INDEX sy-tfill.
        CASE valida_fich-ficheiro.
          WHEN 'jpg' OR 'JPG' OR 'jpeg' OR 'JPEG'.
            IF it64-tipo NE 'Photos'.
              w_return-message_v2 = it64-tipo.
              w_return-message_v1 = it64-ficheiro.
              PERFORM error_processing TABLES return
                      USING 'E' 'ZCKP' '002' w_return-message_v1 w_return-message_v2 '' ''.
              "Ficheiro & não esperado no tipo &.
            ENDIF.
          WHEN 'pdf' OR 'PDF'.
            IF it64-tipo NE 'Proposal Subject for Approval' AND
               it64-tipo NE 'Other Proposal'.
              w_return-message_v2 = it64-tipo.
              w_return-message_v1 = it64-ficheiro.
              PERFORM error_processing TABLES return
                      USING 'E' 'ZCKP' '002' w_return-message_v1 w_return-message_v2 '' ''.
              "Ficheiro & não esperado no tipo &.
            ENDIF.
          WHEN 'msg' OR 'MSG'.
            IF it64-tipo NE 'Email'.
              w_return-message_v2 = it64-tipo.
              w_return-message_v1 = it64-ficheiro.
              PERFORM error_processing TABLES return
                      USING 'E' 'ZCKP' '002' w_return-message_v1 w_return-message_v2 '' ''.
              "Ficheiro & não esperado no tipo &.
            ENDIF.

          WHEN OTHERS.
* erro - tipo inválido
            w_return-message_v1 = it64-ficheiro.
            PERFORM error_processing TABLES return
              USING 'E' 'ZCKP' '001' w_return-message_v1 '' '' ''.
            "Ficheiro & de tipo não esperado.
        ENDCASE.

    ENDCASE.

  ENDLOOP.

  IF NOT return[] IS INITIAL.
    return.
  ENDIF.

  CALL TRANSFORMATION /sbxc/zckp_base64
  SOURCE XML xml_out
  RESULT base64 = it_base64.

*  IF sy-subrc NE 0.
*
*  ENDIF.

  PERFORM deriva_campos
              TABLES poitem
                     poschedule
                     poaccount
               USING poheader.

  REFRESH: return, bapi_item, poitemx, bapi_schedule,
           poschedulex, bapi_account, poaccountx,
           potextitem.


  PERFORM preenche_estruturas

            TABLES poitem
                   poschedule
                   poaccount

             USING poheader.



*"----------------------------------------------------------------------
* BAPI
*"----------------------------------------------------------------------

  CALL FUNCTION 'BAPI_PO_CREATE1' "#EC CI_FLDEXT_OK[2522971]
    EXPORTING
      poheader                     =   bapi_header
      poheaderx                    =   poheaderx
*   POADDRVENDOR                 =
     testrun                      = testrun
*   MEMORY_UNCOMPLETE            =
*   MEMORY_COMPLETE              =
*   POEXPIMPHEADER               =
*   POEXPIMPHEADERX              =
*   VERSIONS                     =
*   NO_MESSAGING                 =
*   NO_MESSAGE_REQ               =
*   NO_AUTHORITY                 =
   NO_PRICE_FROM_PO             = 'X'
   IMPORTING
     exppurchaseorder             = exppurchaseorder
     expheader                    = expheader
     exppoexpimpheader            = exppoexpimpheader
   TABLES
     return                       = return
     poitem                       = bapi_item "#EC CI_USAGE_OK[2438131]
     poitemx                      = poitemx "#EC CI_USAGE_OK[2438131] #EC CI_FLDEXT_OK
*   POADDRDELIVERY               =
     poschedule                   = bapi_schedule
     poschedulex                  = poschedulex
     poaccount                    = bapi_account
*   POACCOUNTPROFITSEGMENT       =
     poaccountx                   = poaccountx
*   POCONDHEADER                 =
*   POCONDHEADERX                =
*   POCOND                       =
*   POCONDX                      =
*   POLIMITS                     =
*   POCONTRACTLIMITS             =
*   POSERVICES                   =
*   POSRVACCESSVALUES            =
*   POSERVICESTEXT               =
*   EXTENSIONIN                  =
*   EXTENSIONOUT                 =
*   POEXPIMPITEM                 =
*   POEXPIMPITEMX                =
*   POTEXTHEADER                 =
   potextitem                   = potextitem
*   ALLVERSIONS                  =
*   POPARTNER                    =
            .


  DATA: l_flag              TYPE c,
        l_off               TYPE i.

  IF exppurchaseorder NE space.

* Imagens

*   Imagem para Ixos
    objecttype = 'BUS2012'. " Pedidos

    LOOP AT it_base64-base64_imp INTO it64.

      PERFORM decode_base64 USING it64-base64
               CHANGING imagem_out l_xstr .


      len = STRLEN( imagem_out ).
      CASE it64-tipo.
        WHEN 'Photos'.
          ar_obj = 'ZACIC_PO03'.
        WHEN 'Email'.
          ar_obj = 'ZACIC_PO04'.
        WHEN 'Proposal Subject for Approval'.
          ar_obj = 'ZACIC_PO01'.  " ZACIC_PO01 ou ZACIC_PO02 !!
        WHEN 'Other Proposal'.
          ar_obj = 'ZACIC_PO02'.
        WHEN OTHERS.

      ENDCASE.

*len = strlen( l_xstr ).



      WHILE l_flag IS INITIAL.

        IF len LE 1024.
          archivobject-line = l_xstr+l_off(len).
          l_flag = 'X'.
        ELSE.
          archivobject-line = l_xstr+l_off(1024).
          l_off = l_off + 1024.
          len = len - 1024.
        ENDIF.

        APPEND archivobject." TO li_contents.

      ENDWHILE.


      DATA: lenght TYPE sapb-length.
      DATA objid LIKE sapb-sapobjid.

      objid = exppurchaseorder. "nº pedido

      lenght = len.

      CALL FUNCTION 'ARCHIV_CREATE_TABLE'
        EXPORTING
          ar_object                      = ar_obj
*   DEL_DATE                       =
          object_id                      = objid
          sap_object                     = objecttype
          flength                        = lenght
*   DOC_TYPE                       =
         document                       = l_xstr
*   MANDT                          = SY-MANDT
       IMPORTING
         outdoc                         = toa
* TABLES
*   ARCHIVOBJECT                   =
*   BINARCHIVOBJECT                = archivobject
       EXCEPTIONS
         error_archiv                   = 1
         error_communicationtable       = 2
         error_connectiontable          = 3
         error_kernel                   = 4
         error_parameter                = 5
         error_mandant                  = 6
         OTHERS                         = 7
                .



      IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
      ENDIF.


    ENDLOOP.



  ELSE.
*  Pedido não criado

  ENDIF.


  CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'.


ENDFUNCTION.
