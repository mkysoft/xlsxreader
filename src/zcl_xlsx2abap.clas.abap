class ZCL_XLSX2ABAP definition
  public
  create public .

public section.

  types:
    BEGIN OF ts_table,
        col   TYPE c LENGTH 3,
        row   TYPE i,
        type  TYPE c LENGTH 1,
        value TYPE string,
      END OF ts_table .
  types:
    BEGIN OF ts_sheet,
        name  TYPE string,
        id    TYPE string,
      END OF ts_sheet .
  types:
    tt_table TYPE STANDARD TABLE OF ts_table WITH KEY COL ROW .
  types:
    tt_sheet TYPE STANDARD TABLE OF ts_sheet WITH KEY NAME .

  methods GET_SHEET
    importing
      !IV_NAME type STRING
    returning
      value(RT_TABLE) type TT_TABLE
    raising
      CX_OPENXML_NOT_FOUND
      CX_OPENXML_FORMAT .
  methods CONSTRUCTOR
    importing
      !IV_FILE type XSTRING
    raising
      CX_OPENXML_FORMAT .
  methods GET_ITAB
    importing
      !IV_NAME type STRING .
  methods GET_SHEETS
    returning
      value(RT_SHEET) type TT_SHEET
    raising
      CX_OPENXML_FORMAT .
protected section.
private section.

  data M_WORKBOOK type ref to CL_XLSX_WORKBOOKPART .
  data M_SHEETS type TT_SHEET .
  data M_XLSX type ref to CL_XLSX_DOCUMENT .
  constants C_NS_R type STRING value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ##NO_TEXT.
  constants C_EXCLDT type DATS value '19000101' ##NO_TEXT.

  methods GET_XMLDOC
    importing
      !IV_XML type XSTRING
    returning
      value(RO_XMLDOC) type ref to IF_IXML_DOCUMENT .
ENDCLASS.



CLASS ZCL_XLSX2ABAP IMPLEMENTATION.


  METHOD constructor.
    m_xlsx = cl_xlsx_document=>load_document( iv_file ).
    m_workbook = m_xlsx->get_workbookpart( ).
  ENDMETHOD.


  method GET_ITAB.
  endmethod.


  METHOD get_sheet.
    DATA: lo_worksheet TYPE REF TO cl_xlsx_worksheetpart.
    DATA: lo_ixml_doc  TYPE REF TO if_ixml_document.
    DATA: ls_sheet TYPE ts_sheet,
          ls_cell  TYPE ts_table.

    READ TABLE m_sheets INTO ls_sheet WITH TABLE KEY name = iv_name.
    IF sy-subrc NE 0.
      RAISE EXCEPTION TYPE cx_openxml_not_found.
    ENDIF.
    lo_worksheet ?= m_workbook->get_part_by_id( ls_sheet-id ).
    lo_ixml_doc = get_xmldoc( lo_worksheet->get_data( ) ).

    " refactoring needed
    TYPES: BEGIN OF ls_table,
             index  TYPE i,
             type   TYPE c LENGTH 1,
             cell   TYPE string,
             value  TYPE string,
             column TYPE string,
             row    TYPE string,
           END OF ls_table.

    TYPES: BEGIN OF ls_string,
             index TYPE i,
             value TYPE string,
           END OF ls_string.

    DATA: ls_string TYPE ls_string,
          lt_string TYPE TABLE OF ls_string,
          ls_table  TYPE ls_table,
          lt_table  TYPE TABLE OF ls_table.

    DATA(lo_ixml_root) = lo_ixml_doc->get_root_element( ).
    DATA(lo_nodes)         = lo_ixml_root->get_elements_by_tag_name( name = 'row' ).
    DATA(lo_node_iterator) = lo_nodes->create_iterator( ).
    DATA(lo_node)          = lo_node_iterator->get_next( ).
    WHILE lo_node IS NOT INITIAL.
      CLEAR ls_table.
      DATA(lo_att) = lo_node->get_attributes( ).
      ls_table-row = lo_att->get_named_item( 'r' )->get_value( ).

      DATA(lo_node_iterator_r) = lo_node->get_children( )->create_iterator( ).
      DATA(lo_node_r)          = lo_node_iterator_r->get_next( ).
      WHILE lo_node_r IS NOT INITIAL.
        CLEAR: ls_table-cell,
               ls_table-type,
               ls_table-value,
               ls_table-index.

        lo_att            = lo_node_r->get_attributes( ).
        DATA(lo_att_child)      = lo_att->get_named_item( 'r' ).
        ls_table-cell = lo_att_child->get_value( ).

        lo_att_child = lo_att->get_named_item( 't' ).
        IF lo_att_child IS BOUND.
          ls_table-type = lo_att_child->get_value( ).
        ENDIF.

        IF ls_table-type IS INITIAL.
          ls_table-value = lo_node_r->get_value( ).
        ELSE.
          ls_table-index = lo_node_r->get_value( ).
        ENDIF.
        APPEND ls_table TO lt_table.
        lo_node_r = lo_node_iterator_r->get_next( ).
      ENDWHILE.
      lo_node          = lo_node_iterator->get_next( ).
    ENDWHILE.

    " string data
    DATA(lo_shared_st)  = m_workbook->get_sharedstringspart( ).
    lo_ixml_doc = get_xmldoc( lo_shared_st->get_data( ) ).
    lo_ixml_root = lo_ixml_doc->get_root_element( ).
    lo_nodes         = lo_ixml_root->get_elements_by_tag_name( name = 'si' ).
    lo_node_iterator = lo_nodes->create_iterator( ).

    lo_node = lo_node_iterator->get_next( ).
    WHILE lo_node IS NOT INITIAL.
      CLEAR: ls_string.
      ls_string-index = sy-tabix.
      ls_string-value = lo_node->get_value( ).
      APPEND ls_string TO lt_string.
      lo_node = lo_node_iterator->get_next( ).
    ENDWHILE.

    LOOP AT lt_table INTO ls_table.
      "get column
      ls_table-column = ls_table-cell.
      CONDENSE ls_table-row NO-GAPS.
      REPLACE ls_table-row IN ls_table-column WITH space.

      IF ls_table-type EQ 's'.
        READ TABLE lt_string INTO ls_string
          WITH KEY index = ls_table-index BINARY SEARCH.
        IF sy-subrc EQ 0.
          ls_table-value = ls_string-value.
        ENDIF.
      ENDIF.
      CONDENSE ls_table-value.
      CLEAR ls_cell.
      ls_cell-row = ls_table-row.
      ls_cell-col = ls_table-column.
      ls_cell-type = ls_table-type.
      ls_cell-value = ls_table-value.
      APPEND ls_cell TO rt_table.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_sheets.
    DATA: ls_sheet TYPE ts_sheet.

    IF m_sheets IS INITIAL.
      DATA(lo_ixml_doc) = get_xmldoc( m_workbook->get_data( ) ).
      DATA(lo_ixml_root)     = lo_ixml_doc->get_root_element( ).
      DATA(lo_nodes)         = lo_ixml_root->get_elements_by_tag_name( name = 'sheet' ).
      DATA(lo_node_iterator) = lo_nodes->create_iterator( ).
      DATA(lo_node)          = lo_node_iterator->get_next( ).
      WHILE lo_node IS NOT INITIAL.
        DATA(lo_att)  = lo_node->get_attributes( ).
        ls_sheet-name = lo_att->get_named_item( 'name' )->get_value( ).
        ls_sheet-id   = lo_att->get_named_item_ns( name = 'id' uri = c_ns_r )->get_value( ).
        APPEND ls_sheet TO me->m_sheets.
        lo_node = lo_node_iterator->get_next( ).
      ENDWHILE.
    ENDIF.
    rt_sheet = m_sheets.
  ENDMETHOD.


  METHOD get_xmldoc.
    DATA(lo_ixml) = cl_ixml=>create( ).
    DATA(lo_ixml_sf) = lo_ixml->create_stream_factory( ).
    DATA(lo_ixml_stream) = lo_ixml_sf->create_istream_xstring( iv_xml ).
    ro_xmldoc = lo_ixml->create_document( ).
    DATA(lo_ixml_parser) = lo_ixml->create_parser( document = ro_xmldoc  istream = lo_ixml_stream stream_factory = lo_ixml_sf ).
    lo_ixml_parser->parse( ).
  ENDMETHOD.
ENDCLASS.
