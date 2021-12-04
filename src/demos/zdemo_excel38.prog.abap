REPORT zdemo_excel38.


CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_demo_excel_generator.

  PUBLIC SECTION.
    METHODS zif_demo_excel_generator~checker_initialization REDEFINITION.
    METHODS zif_demo_excel_generator~get_information REDEFINITION.
    METHODS zif_demo_excel_generator~generate_excel REDEFINITION.
    METHODS zif_demo_excel_generator~cleanup_for_diff REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '38_SAP-Icons.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


TABLES: icon.
SELECT-OPTIONS: s_icon FOR icon-name DEFAULT 'ICON_LED_*' OPTION CP.

START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_demo_excel_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_demo_excel_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Read file and output'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~checker_initialization.

    CLEAR: s_icon, s_icon[].
    s_icon-sign = 'I'.
    s_icon-option = 'CP'.
    s_icon-low = 'ICON_LED_*'.
    APPEND s_icon TO s_icon.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~generate_excel.

    DATA: lo_excel     TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_column    TYPE REF TO zcl_excel_column,
          lo_drawing   TYPE REF TO zcl_excel_drawing.

    TYPES: BEGIN OF gty_icon,
*         name      TYPE icon_name, "Fix #228
             name  TYPE iconname,   "Fix #228
             objid TYPE w3objid,
           END OF gty_icon,
           gtyt_icon TYPE STANDARD TABLE OF gty_icon WITH NON-UNIQUE DEFAULT KEY.

    DATA: lt_icon       TYPE gtyt_icon,
          lv_row        TYPE sytabix,
          ls_wwwdatatab TYPE wwwdatatab,
          lt_mimedata   TYPE STANDARD TABLE OF w3mime WITH NON-UNIQUE DEFAULT KEY,
          lv_xstring    TYPE xstring.

    FIELD-SYMBOLS: <icon>     LIKE LINE OF lt_icon,
                   <mimedata> LIKE LINE OF lt_mimedata.

    " Creates active sheet
    CREATE OBJECT lo_excel.

    " Get active sheet
    lo_worksheet = lo_excel->get_active_worksheet( ).
    lo_worksheet->set_title( ip_title = 'Demo Iconls' ).
    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_auto_size( 'X' ).
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_auto_size( 'X' ).

* Get all icons
    SELECT name
      INTO TABLE lt_icon
      FROM icon
      WHERE name IN s_icon
      ORDER BY name.
    LOOP AT lt_icon ASSIGNING <icon>.

      lv_row = sy-tabix.
*--------------------------------------------------------------------*
* Set name of icon
*--------------------------------------------------------------------*
      lo_worksheet->set_cell( ip_row = lv_row
                              ip_column = 'A'
                              ip_value = <icon>-name ).
*--------------------------------------------------------------------*
* Check whether the mime-repository holds some icondata for us
*--------------------------------------------------------------------*

* Get key
      SELECT SINGLE objid
        INTO <icon>-objid
        FROM wwwdata
        WHERE text = <icon>-name.
      CHECK sy-subrc = 0.  " :o(
      lo_worksheet->set_cell( ip_row = lv_row
                              ip_column = 'B'
                              ip_value = <icon>-objid ).

* Load mimedata
      CLEAR lt_mimedata.
      CLEAR ls_wwwdatatab.
      ls_wwwdatatab-relid = 'MI' .
      ls_wwwdatatab-objid = <icon>-objid.
      CALL FUNCTION 'WWWDATA_IMPORT'
        EXPORTING
          key               = ls_wwwdatatab
        TABLES
          mime              = lt_mimedata
        EXCEPTIONS
          wrong_object_type = 1
          import_error      = 2
          OTHERS            = 3.
      CHECK sy-subrc = 0.  " :o(

      lo_drawing = lo_excel->add_new_drawing( ).
      lo_drawing->set_position( ip_from_row = lv_row
                                ip_from_col = 'C' ).
      CLEAR lv_xstring.
      LOOP AT lt_mimedata ASSIGNING <mimedata>.
        CONCATENATE lv_xstring <mimedata>-line INTO lv_xstring IN BYTE MODE.
      ENDLOOP.

      lo_drawing->set_media( ip_media      = lv_xstring
                             ip_media_type = zcl_excel_drawing=>c_media_type_jpg
                             ip_width      = 16
                             ip_height     = 14  ).
      lo_worksheet->add_drawing( lo_drawing ).

    ENDLOOP.

    result = lo_excel.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~cleanup_for_diff.

    DATA: zip     TYPE REF TO cl_abap_zip,
          content TYPE xstring.
    FIELD-SYMBOLS: <file> TYPE cl_abap_zip=>t_file.

    zip = super->zif_demo_excel_generator~cleanup_for_diff( xstring ).

    LOOP AT zip->files ASSIGNING <file>
        WHERE name CP 'xl/drawings/drawing*.xml'.
      zip->get(
        EXPORTING
          name                    = <file>-name
        IMPORTING
          content                 = content
        EXCEPTIONS
          zip_index_error         = 1
          zip_decompression_error = 2
          OTHERS                  = 3 ).
      IF sy-subrc <> 0.
*     MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.
      DATA lo_ixml TYPE REF TO if_ixml.
      DATA lo_streamfactory TYPE REF TO if_ixml_stream_factory.
      DATA lo_istream TYPE REF TO if_ixml_istream.
      DATA lo_parser TYPE REF TO if_ixml_parser.
      DATA lo_renderer TYPE REF TO if_ixml_renderer.
      DATA lo_ostream TYPE REF TO if_ixml_ostream.
      DATA lo_document TYPE REF TO if_ixml_document.
      DATA lo_element TYPE REF TO if_ixml_element.
      DATA: file_name TYPE cl_abap_zip=>t_file-name,
            name TYPE string.
      lo_ixml = cl_ixml=>create( ).
      lo_streamfactory = lo_ixml->create_stream_factory( ).
      lo_istream = lo_streamfactory->create_istream_xstring( content ).
      lo_document = lo_ixml->create_document( ).
      lo_parser = lo_ixml->create_parser(
                document       = lo_document
                istream        = lo_istream
                stream_factory = lo_streamfactory ).
      lo_parser->parse( ).
*      DATA(namespace_context) = lo_document->get_namespace_context( ).
*      DATA(prefix) = namespace_context->map_to_prefix( 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
      DATA(filter) = lo_document->create_filter_name_ns( name = 'cNvPr' namespace = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
      DATA(iterator) = lo_document->create_iterator_filtered( filter ).
      DO.
        lo_element ?= iterator->get_next( ).
        IF lo_element IS NOT BOUND.
          EXIT.
        ENDIF.
*      lo_element = lo_document->find_from_name_ns( name = 'cNvPr' uri = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
*      WHILE lo_element IS BOUND.
        name = lo_element->get_name( ).
        data(id) = lo_element->get_attribute_ns( name = 'id' )."uri = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
        name = lo_element->get_attribute_ns( name = 'name' )."uri = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
        lo_element->set_attribute_ns( name = 'name' value = '' ).
*        lo_element->set_attribute( name = 'name' value = '' ).
*        lo_element ?= lo_element->get_next( ).
*      ENDWHILE.
      ENDDO.

      CLEAR content.
      lo_ostream = lo_streamfactory->create_ostream_xstring( content ).
      lo_renderer = lo_ixml->create_renderer(
                  document = lo_document
                  ostream  = lo_ostream ).
      lo_renderer->render( ).

      TYPES: BEGIN OF ty_file,
               name    TYPE string,
               content TYPE xstring,
             END OF ty_file.
      DATA: ls_file TYPE ty_file,
            lt_file TYPE TABLE OF ty_file.
      ls_file-name = <file>-name.
      ls_file-content = content.
      APPEND ls_file TO lt_file.

    ENDLOOP.

    FIELD-SYMBOLS: <ls_file2> TYPE ty_file.
    LOOP AT lt_file ASSIGNING <ls_file2>.
      zip->delete( name = <ls_file2>-name ).
      zip->add( name = <ls_file2>-name content = <ls_file2>-content ).
    ENDLOOP.

    result = zip.

  ENDMETHOD.

ENDCLASS.
