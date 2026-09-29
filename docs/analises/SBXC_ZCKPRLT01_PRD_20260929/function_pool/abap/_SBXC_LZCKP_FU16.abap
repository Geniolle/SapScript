FUNCTION /sbxc/zckp_guarda_alt.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(PROCESSO)
*"     REFERENCE(CAB)
*"  TABLES
*"      LINHA
*"      COR STRUCTURE  /SBXC/ZCKP_COLOR
*"----------------------------------------------------------------------

* Estruturas de cabeçalho e linha
  DATA: header LIKE /sbxc/zckp_invh.
  DATA: item TYPE TABLE OF /sbxc/zckp_invi WITH HEADER LINE.


  TYPE-POOLS : abap.
  FIELD-SYMBOLS: <header> TYPE STANDARD TABLE,
                 <dyn_wa> type any,
                 <dyn_field> type any.

  DATA: dy_table TYPE REF TO data,
        dy_line  TYPE REF TO data,
        xfc TYPE lvc_s_fcat,
        ifc TYPE lvc_t_fcat.

  DATA : idetails TYPE abap_compdescr_tab,
         xdetails TYPE abap_compdescr.
  DATA : ref_table_des TYPE REF TO cl_abap_structdescr.



*  TABLES: dd03l.
  MOVE-CORRESPONDING cab TO header.

  REFRESH: idetails.


  SELECT SINGLE * FROM /sbxc/zckp_tab00
         WHERE processo = processo.

  ref_table_des ?=
      cl_abap_typedescr=>describe_by_name( /sbxc/zckp_tab00-est_cab ).

  idetails[] = ref_table_des->components[].

  LOOP AT idetails INTO xdetails.
    CLEAR xfc.
    xfc-fieldname = xdetails-name .
    CASE xdetails-type_kind.
      WHEN 'C'.
        xfc-datatype = 'CHAR'.
      WHEN 'N'.
        xfc-datatype = 'NUMC'.
      WHEN 'D'.
        xfc-datatype = 'DATE'.
      WHEN 'P'.
        xfc-datatype = 'PACK'.
      WHEN OTHERS.
        xfc-datatype = xdetails-type_kind.
    ENDCASE.
    xfc-inttype = xdetails-type_kind.
    xfc-intlen = xdetails-length.
    xfc-decimals = xdetails-decimals.

    APPEND xfc TO ifc.
  ENDLOOP.

  CALL METHOD cl_alv_table_create=>create_dynamic_table
    EXPORTING
      it_fieldcatalog  = ifc
      i_length_in_byte = 'X'
    IMPORTING
      ep_table         = dy_table.
  ASSIGN dy_table->* TO <header>.

  CREATE DATA dy_line LIKE LINE OF <header>.
  ASSIGN dy_line->* TO <dyn_wa>.

  MOVE-CORRESPONDING <dyn_wa> TO <header>.
  APPEND <dyn_wa>  TO <header>.

  LOOP AT linha.
    MOVE-CORRESPONDING linha TO item.
    APPEND item.
  ENDLOOP.

  SELECT * FROM  dd03l
    WHERE  tabname  = /sbxc/zckp_tab00-est_cab.
  ENDSELECT.

ENDFUNCTION.
