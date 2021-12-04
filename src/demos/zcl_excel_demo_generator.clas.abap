CLASS zcl_excel_demo_generator DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES zif_excel_demo_generator.
    CLASS-METHODS class_constructor.
    CLASS-METHODS get_date_now
      RETURNING
        VALUE(result) TYPE d.
    CLASS-METHODS get_time_now
      RETURNING
        VALUE(result) TYPE t.
    CLASS-METHODS set_date_now
      IMPORTING
        date TYPE d.
    CLASS-METHODS set_time_now
      IMPORTING
        time TYPE t.

  PROTECTED SECTION.
  PRIVATE SECTION.
    CLASS-DATA: date_now TYPE d,
                time_now TYPE t.
ENDCLASS.



CLASS zcl_excel_demo_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~checker_initialization.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~get_information.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~get_next_generator.

  ENDMETHOD.

  METHOD get_date_now.

    result = date_now.

  ENDMETHOD.

  METHOD get_time_now.

    result = time_now.

  ENDMETHOD.

  METHOD set_date_now.

    date_now = date.

  ENDMETHOD.

  METHOD set_time_now.

    time_now = time.

  ENDMETHOD.

  METHOD class_constructor.

    date_now = sy-datum.
    time_now = sy-uzeit.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~cleanup_for_diff.

    DATA: zip     TYPE REF TO cl_abap_zip,
          content TYPE xstring.

    CREATE OBJECT zip.
    zip->load(
      EXPORTING
        zip             = xstring
      EXCEPTIONS
        zip_parse_error = 1
        OTHERS          = 2 ).
    IF sy-subrc <> 0.
*   MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
    zip->get(
      EXPORTING
        name                    = 'docProps/core.xml'
      IMPORTING
        content                 = content
      EXCEPTIONS
        zip_index_error         = 1
        zip_decompression_error = 2
        OTHERS                  = 3 ).
    IF sy-subrc <> 0.
      RETURN.
*   MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    TYPES: BEGIN OF ty_docprops_core,
             creator          TYPE string,
             description      TYPE string,
             last_modified_by TYPE string,
             created          TYPE string,
             modified         TYPE string,
           END OF ty_docprops_core.
    DATA: docprops_core TYPE ty_docprops_core.

    TRY.
        CALL TRANSFORMATION zexcel_tr_docprops_core SOURCE XML content RESULT root = docprops_core.
      CATCH cx_root INTO DATA(lx).
        RETURN.
    ENDTRY.

    CLEAR: docprops_core-creator,
           docprops_core-description,
           docprops_core-created,
           docprops_core-modified.

    CALL TRANSFORMATION zexcel_tr_docprops_core SOURCE root = docprops_core RESULT XML content.

    zip->delete(
      EXPORTING
        name            = 'docProps/core.xml'
      EXCEPTIONS
        zip_index_error = 1
        OTHERS          = 2 ).
    IF sy-subrc <> 0.
*   MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    zip->add(
        name    = 'docProps/core.xml'
        content = content ).

    FIELD-SYMBOLS: <file> TYPE cl_abap_zip=>t_file.

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
            name      TYPE string,
            filter    TYPE REF TO if_ixml_node_filter,
            iterator  TYPE REF TO if_ixml_node_iterator.
      lo_ixml = cl_ixml=>create( ).
      lo_streamfactory = lo_ixml->create_stream_factory( ).
      lo_istream = lo_streamfactory->create_istream_xstring( content ).
      lo_document = lo_ixml->create_document( ).
      lo_parser = lo_ixml->create_parser(
                document       = lo_document
                istream        = lo_istream
                stream_factory = lo_streamfactory ).
      lo_parser->parse( ).
      filter = lo_document->create_filter_name_ns( name = 'cNvPr' namespace = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
      iterator = lo_document->create_iterator_filtered( filter ).
      DO.
        lo_element ?= iterator->get_next( ).
        IF lo_element IS NOT BOUND.
          EXIT.
        ENDIF.
        lo_element->set_attribute_ns( name = 'name' value = '' ).
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
