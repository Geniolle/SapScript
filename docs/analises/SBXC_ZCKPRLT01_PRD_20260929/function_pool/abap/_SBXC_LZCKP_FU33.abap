FUNCTION /sbxc/zckp_dm_fornecedor.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     VALUE(DATA) TYPE  DATS
*"     VALUE(CARR_INICIAL) TYPE  CHAR1
*"  TABLES
*"      VENDORHEADER STRUCTURE  /SBXC/ZBAPIVENDOR_HEAD
*"      VENDORPURCHORG STRUCTURE  /SBXC/ZBAPIVENDOR_PURCH_ORG
*"      VENDORBANK STRUCTURE  /SBXC/ZBAPIVENDOR_BANK
*"      VENDORCOMPANY STRUCTURE  /SBXC/ZBAPIVENDOR_COMP
*"----------------------------------------------------------------------

* Criado por Cláudia Fernandes @ SBXC
* Função para obter dados mestres de fornecedores criados/alterados numa
* determinada data - Se data vazia, envia fornecedores alterados/criados
* no dia da execução
************************************************************************

  t_land-mandant = sy-mandt.

*se pisco de carregamento inicial estiver activo então selecciona todos
*os fornecedores independentemente
* da data de modificação
  IF carr_inicial = 'X'.

* testes - Apenas empresa SPCC
    SELECT lifnr bukrs zwels FROM lfb1 CLIENT SPECIFIED
      INTO (vendorcompany-vendor, vendorcompany-comp_code,
            vendorcompany-payment_methods)
      WHERE mandt = t_land-mandant AND
            bukrs = 'SPCC'.

      vendorcompany-mandante_pais = t_land-land1.

      SELECT SINGLE lifnr adrnr name1 name2 name3 land1 stceg  stcd1 stcd2
         stenr stras pstlz ort01
           FROM lfa1  CLIENT SPECIFIED
                INTO (vendorheader-vendor, num_end, vendorheader-name1,
                vendorheader-name2,
                       vendorheader-name3, vendorheader-land1,
                       vendorheader-stceg,
                        stcd1, stcd2,  stenr,
                       vendorheader-street, vendorheader-post_code1,
                       vendorheader-city2)
                 WHERE mandt = t_land-mandant AND
                        lifnr = vendorcompany-vendor AND
                       sperr = ' ' .

      CHECK sy-subrc = 0.
* ir à tabela adrc6 buscar endereço de email
      SELECT SINGLE smtp_addr INTO vendorheader-smtp_addr FROM adr6  "#EC CI_NOORDER
      CLIENT SPECIFIED
        WHERE  client = t_land-mandant AND
               addrnumber = num_end.


*se campo stceg vazio preenche com stcd1 se este estiver vazio preenche
*stcd2 se vazio preenche stenr
      IF vendorheader-stceg IS INITIAL AND stcd1 IS NOT INITIAL.
        vendorheader-stceg = stcd1.
      ELSEIF vendorheader-stceg IS INITIAL AND stcd2 IS NOT INITIAL.
        vendorheader-stceg = stcd2.
      ELSEIF vendorheader-stceg IS INITIAL AND stenr IS NOT INITIAL.
        vendorheader-stceg = stenr.
      ENDIF.
      vendorheader-mandante_pais = t_land-land1.
      APPEND vendorheader.

* ir à tabela LFM1 ler inf. de cond.pagamento por org.compras
      SELECT lifnr ekorg zterm FROM lfm1 CLIENT SPECIFIED
        INTO (vendorpurchorg-vendor,vendorpurchorg-purch_org,
        vendorpurchorg-pmnttrms)
             WHERE  mandt = t_land-mandant AND
                    lifnr =    vendorheader-vendor.
        vendorpurchorg-mandante_pais = t_land-land1.
        APPEND vendorpurchorg.
      ENDSELECT.

      SELECT lifnr banks bankl bankn  FROM lfbk CLIENT SPECIFIED
        INTO (vendorbank-vendor, vendorbank-banks, vendorbank-bankl,
        vendorbank-bankn)
      WHERE
        mandt = t_land-mandant AND
        lifnr =    vendorheader-vendor.
        vendorbank-mandante_pais = t_land-land1.
        APPEND  vendorbank.
      ENDSELECT.

      APPEND  vendorcompany.

    ENDSELECT.

  ELSE.
    SELECT objectid  FROM cdhdr  CLIENT SPECIFIED INTO CORRESPONDING
    FIELDS OF TABLE t_cdhdr WHERE
          mandant = t_land-mandant AND
           objectclas = 'ADRESSE' AND
          objectid LIKE 'BP%' AND
           udate = data.

    SELECT objectid  FROM cdhdr  CLIENT SPECIFIED INTO CORRESPONDING
    FIELDS OF  t_cdhdr WHERE
           mandant = t_land-mandant AND
            objectclas = 'KRED' AND
              udate = data.
      APPEND t_cdhdr.
    ENDSELECT.
    SELECT objectid  FROM pcdhdr  CLIENT SPECIFIED INTO CORRESPONDING
    FIELDS OF  t_cdhdr WHERE
           mandant = t_land-mandant AND
            objectclas = 'ZCKP_LIFNR' AND
              udate = data.
      APPEND t_cdhdr.
    ENDSELECT.
  ENDIF.


  DELETE ADJACENT DUPLICATES FROM t_cdhdr. "#EC CI_SORTED

  LOOP AT t_cdhdr.
    IF t_cdhdr-objectid+0(2) = 'BP'.
      num_end = t_cdhdr-objectid+4(10).
    ELSE.
      forn = t_cdhdr-objectid.
    ENDIF.

*para todos os regitos encontrados procurar na tabela LFA1 ler inf.
*relevante
    SELECT lifnr adrnr name1 name2 name3 land1 stceg  stcd1 stcd2
    stenr stras pstlz ort01
      FROM lfa1  CLIENT SPECIFIED
           INTO (vendorheader-vendor, num_end, vendorheader-name1,
           vendorheader-name2,
                  vendorheader-name3, vendorheader-land1,
                  vendorheader-stceg,
                   stcd1, stcd2,  stenr,
                  vendorheader-street, vendorheader-post_code1,
                  vendorheader-city2)
            WHERE mandt = t_land-mandant AND
                  ( adrnr = num_end OR lifnr = forn ) AND
                  sperr = ' ' .

* ir à tabela adrc6 buscar endereço de email
      SELECT SINGLE smtp_addr INTO vendorheader-smtp_addr FROM adr6  "#EC CI_NOORDER
      CLIENT SPECIFIED
        WHERE  client = t_land-mandant AND
               addrnumber = num_end.


*se campo stceg vazio preenche com stcd1 se este estiver vazio preenche
*stcd2 se vazio preenche stenr
      IF vendorheader-stceg IS INITIAL AND stcd1 IS NOT INITIAL.
        vendorheader-stceg = stcd1.
      ELSEIF vendorheader-stceg IS INITIAL AND stcd2 IS NOT INITIAL.
        vendorheader-stceg = stcd2.
      ELSEIF vendorheader-stceg IS INITIAL AND stenr IS NOT INITIAL.
        vendorheader-stceg = stenr.
      ENDIF.
      vendorheader-mandante_pais = t_land-land1.
      APPEND vendorheader.

* ir à tabela LFM1 ler inf. de cond.pagamento por org.compras
      SELECT lifnr ekorg zterm FROM lfm1 CLIENT SPECIFIED
        INTO (vendorpurchorg-vendor,vendorpurchorg-purch_org,
        vendorpurchorg-pmnttrms)
             WHERE  mandt = t_land-mandant AND
                    lifnr =    vendorheader-vendor.
        vendorpurchorg-mandante_pais = t_land-land1.
        APPEND vendorpurchorg.
      ENDSELECT.

      SELECT lifnr banks bankl bankn  FROM lfbk CLIENT SPECIFIED
        INTO (vendorbank-vendor, vendorbank-banks, vendorbank-bankl,
        vendorbank-bankn)
      WHERE
        mandt = t_land-mandant AND
        lifnr =    vendorheader-vendor.
        vendorbank-mandante_pais = t_land-land1.
        APPEND  vendorbank.
      ENDSELECT.

      SELECT lifnr bukrs zwels   FROM lfb1 CLIENT SPECIFIED
        INTO (vendorcompany-vendor, vendorcompany-comp_code,
        vendorcompany-payment_methods)
                WHERE    mandt = t_land-mandant AND
                         lifnr =    vendorheader-vendor.
        vendorcompany-mandante_pais = t_land-land1.
        APPEND  vendorcompany.
      ENDSELECT.

      CLEAR: stcd1, stcd2,  stenr.
    ENDSELECT.


  ENDLOOP.
  SORT: vendorheader, vendorpurchorg, vendorbank, vendorcompany.

  DELETE ADJACENT DUPLICATES FROM vendorheader.
  DELETE ADJACENT DUPLICATES FROM vendorpurchorg.
  DELETE ADJACENT DUPLICATES FROM vendorbank.
  DELETE ADJACENT DUPLICATES FROM vendorcompany.

*verifica NIF empresa compradora
  LOOP AT vendorcompany.
    IF vendorcompany-nif_empresa_comp IS INITIAL.

      PERFORM verifica_nif_emp USING vendorcompany-comp_code
               CHANGING vendorcompany-nif_empresa_comp.
      MODIFY vendorcompany TRANSPORTING nif_empresa_comp WHERE
             comp_code = vendorcompany-comp_code.
    ENDIF.
  ENDLOOP.

ENDFUNCTION.
