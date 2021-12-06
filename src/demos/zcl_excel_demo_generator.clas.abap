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

    TYPES : BEGIN OF ty_zip_structure,
              ref_to_structure TYPE REF TO data,
              ref_to_x         TYPE REF TO data,
              length           TYPE i,
              view             TYPE REF TO cl_abap_view_offlen,
              charset_bit      TYPE i,
              conv_in_utf8     TYPE REF TO cl_abap_conv_in_ce,
              conv_in_ibm437   TYPE REF TO cl_abap_conv_in_ce,
              conv_out_utf8    TYPE REF TO cl_abap_conv_out_ce,
              conv_out_ibm437  TYPE REF TO cl_abap_conv_out_ce,
            END OF ty_zip_structure.

    CLASS-DATA: date_now TYPE d,
                time_now TYPE t.

    METHODS zip_cleanup_for_diff
      IMPORTING
        zip_xstring   TYPE xstring
      RETURNING
        VALUE(result) TYPE xstring
      RAISING
        zcx_excel.

    METHODS init_structure
      IMPORTING
        length        TYPE i
        charset_bit   TYPE i
        structure     TYPE any
      RETURNING
        VALUE(result) TYPE ty_zip_structure.

    METHODS write_zip
      IMPORTING
        offset        TYPE i
      CHANGING
        zip_structure TYPE zcl_excel_demo_generator=>ty_zip_structure
        zip_xstring   TYPE xstring.

    METHODS read_zip
      IMPORTING
        zip_xstring   TYPE xstring
        offset        TYPE i
      CHANGING
        zip_structure TYPE zcl_excel_demo_generator=>ty_zip_structure.

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

    TYPES: BEGIN OF ty_docprops_core,
             creator          TYPE string,
             description      TYPE string,
             last_modified_by TYPE string,
             created          TYPE string,
             modified         TYPE string,
           END OF ty_docprops_core.
    TYPES: BEGIN OF ty_file,
             name    TYPE string,
             content TYPE xstring,
           END OF ty_file.
    DATA: zip           TYPE REF TO cl_abap_zip,
          content       TYPE xstring,
          docprops_core TYPE ty_docprops_core,
          lx            TYPE REF TO cx_root.
    DATA: ls_file          TYPE ty_file,
          lt_file          TYPE TABLE OF ty_file,
          lo_ixml          TYPE REF TO if_ixml,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_istream       TYPE REF TO if_ixml_istream,
          lo_parser        TYPE REF TO if_ixml_parser,
          lo_renderer      TYPE REF TO if_ixml_renderer,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_document      TYPE REF TO if_ixml_document,
          lo_element       TYPE REF TO if_ixml_element,
          lo_filter        TYPE REF TO if_ixml_node_filter,
          lo_iterator      TYPE REF TO if_ixml_node_iterator.
    FIELD-SYMBOLS:
      <file>     TYPE cl_abap_zip=>t_file,
      <ls_file2> TYPE ty_file.

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


    TRY.
        CALL TRANSFORMATION zexcel_tr_docprops_core SOURCE XML content RESULT root = docprops_core.
      CATCH cx_root INTO lx.
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

      lo_ixml = cl_ixml=>create( ).
      lo_streamfactory = lo_ixml->create_stream_factory( ).
      lo_istream = lo_streamfactory->create_istream_xstring( content ).
      lo_document = lo_ixml->create_document( ).
      lo_parser = lo_ixml->create_parser(
                  document       = lo_document
                  istream        = lo_istream
                  stream_factory = lo_streamfactory ).
      lo_parser->parse( ).

      lo_filter = lo_document->create_filter_name_ns(
                  name      = 'cNvPr'
                  namespace = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ).
      lo_iterator = lo_document->create_iterator_filtered( lo_filter ).
      DO.
        lo_element ?= lo_iterator->get_next( ).
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

      ls_file-name = <file>-name.
      ls_file-content = content.
      APPEND ls_file TO lt_file.

    ENDLOOP.

    LOOP AT lt_file ASSIGNING <ls_file2>.
      zip->delete( name = <ls_file2>-name ).
      zip->add( name = <ls_file2>-name content = <ls_file2>-content ).
    ENDLOOP.

    result = zip->save( ).
    result = zip_cleanup_for_diff( result ).

  ENDMETHOD.


  METHOD zip_cleanup_for_diff.

    TYPES : BEGIN OF ty_local_file_header,
              local_file_header_signature TYPE x LENGTH 4,  " 04034b50
              version_needed_to_extract   TYPE x LENGTH 2,
              general_purpose_bit_flag    TYPE x LENGTH 2,
              compression_method          TYPE x LENGTH 2,
              last_mod_file_time          TYPE int2,
              last_mod_file_date          TYPE int2,
              crc_32                      TYPE x LENGTH 4,
              compressed_size             TYPE i,
              uncompressed_size           TYPE i,
              file_name_length            TYPE int2,
              extra_field_length          TYPE int2,
              " file name (variable size)
              " extra field (variable size)
            END OF ty_local_file_header,
            BEGIN OF ty_central_file_header,
              central_file_header_signature TYPE x LENGTH 4, " 02014b50
              version_made_by               TYPE x LENGTH 2,
              version_needed_to_extract     TYPE x LENGTH 2,
              general_purpose_bit_flag      TYPE x LENGTH 2,
              compression_method            TYPE x LENGTH 2,
              last_mod_file_time            TYPE int2,
              last_mod_file_date            TYPE int2,
              crc_32                        TYPE x LENGTH 4,
              compressed_size               TYPE i,
              uncompressed_size             TYPE i,
              file_name_length              TYPE int2, " field 12
              extra_field_length            TYPE int2, " field 13
              file_comment_length           TYPE int2, " field 14
              disk_number_start             TYPE int2,
              internal_file_attributes      TYPE x LENGTH 2,
              external_file_attributes      TYPE x LENGTH 4,
              rel_offset_of_local_header    TYPE x LENGTH 4,
              " file name                       (variable size defined in 12)
              " extra field                     (variable size defined in 13)
              " file comment                    (variable size defined in 14)
            END OF ty_central_file_header,
            BEGIN OF ty_end_of_central_dir,
              signature                      TYPE x LENGTH 4, " 0x06054b50
              number_of_this_disk            TYPE int2,
              disk_num_start_of_central_dir  TYPE int2,
              n_of_entries_in_central_dir_dk TYPE int2,
              n_of_entries_in_central_dir    TYPE int2,
              size_of_central_dir            TYPE i,
              offset_start_of_central_dir    TYPE i,
              file_comment_length            TYPE int2,
            END OF ty_end_of_central_dir.

    FIELD-SYMBOLS:
      <local_file_header_x>   TYPE x,
      <central_file_header_x> TYPE x,
      <end_of_central_dir_x>  TYPE x,
      <local_file_header>     TYPE ty_local_file_header,
      <central_file_header>   TYPE ty_central_file_header,
      <end_of_central_dir>    TYPE ty_end_of_central_dir.
    CONSTANTS:
      local_file_header_signature   TYPE x LENGTH 4 VALUE '504B0304',
      central_file_header_signature TYPE x LENGTH 4 VALUE '504B0102',
      end_of_central_dir_signature  TYPE x LENGTH 4 VALUE '504B0506'.
    DATA:
      dummy_local_file_header   TYPE ty_local_file_header,
      dummy_central_file_header TYPE ty_central_file_header,
      dummy_end_of_central_dir  TYPE ty_end_of_central_dir,
      local_file_header         TYPE zcl_excel_demo_generator=>ty_zip_structure,
      central_file_header       TYPE zcl_excel_demo_generator=>ty_zip_structure,
      end_of_central_dir        TYPE zcl_excel_demo_generator=>ty_zip_structure,
      offset                    TYPE i,
      max_offset                TYPE i.



    local_file_header = init_structure( length = 30 charset_bit = 60 structure = dummy_local_file_header ).
    ASSIGN local_file_header-ref_to_structure->* TO <local_file_header>.
    ASSIGN local_file_header-ref_to_x->* TO <local_file_header_x>.

    central_file_header = init_structure( length = 46 charset_bit = 76 structure = dummy_central_file_header ).
    ASSIGN central_file_header-ref_to_structure->* TO <central_file_header>.
    ASSIGN central_file_header-ref_to_x->* TO <central_file_header_x>.

    end_of_central_dir = init_structure( length = 22 charset_bit = 0 structure = dummy_end_of_central_dir ).
    ASSIGN end_of_central_dir-ref_to_structure->* TO <end_of_central_dir>.
    ASSIGN end_of_central_dir-ref_to_x->* TO <end_of_central_dir_x>.

    result = zip_xstring.

    offset = 0.
    max_offset = xstrlen( result ) - 4.
    WHILE offset <= max_offset.

      CASE result+offset(4).

        WHEN local_file_header_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = local_file_header ).

          CLEAR <local_file_header>-last_mod_file_date.
          CLEAR <local_file_header>-last_mod_file_time.

          write_zip( EXPORTING offset = offset CHANGING zip_structure = local_file_header zip_xstring = result ).

          offset = offset + local_file_header-length + <local_file_header>-file_name_length + <local_file_header>-extra_field_length + <local_file_header>-compressed_size.

        WHEN central_file_header_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = central_file_header ).

          CLEAR <central_file_header>-last_mod_file_date.
          CLEAR <central_file_header>-last_mod_file_time.

          write_zip( EXPORTING offset = offset CHANGING zip_structure = central_file_header zip_xstring = result ).

          offset = offset + central_file_header-length + <central_file_header>-file_name_length + <central_file_header>-extra_field_length + <central_file_header>-file_comment_length.

        WHEN end_of_central_dir_signature.

          read_zip( EXPORTING zip_xstring = result offset = offset CHANGING zip_structure = end_of_central_dir ).

          offset = offset + end_of_central_dir-length + <end_of_central_dir>-file_comment_length.

        WHEN OTHERS.
          RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Invalid ZIP file'.

      ENDCASE.

    ENDWHILE.

  ENDMETHOD.


  METHOD init_structure.

    DATA:
      offset      TYPE i,
      rtts_struct TYPE REF TO cl_abap_structdescr.
    FIELD-SYMBOLS:
      <component> TYPE abap_compdescr.

    CREATE DATA result-ref_to_structure LIKE structure.
    result-length = length.
    result-charset_bit = charset_bit.
    CREATE DATA result-ref_to_x TYPE x LENGTH length.

    result-view = cl_abap_view_offlen=>create( ).
    offset = 0.
    rtts_struct ?= cl_abap_typedescr=>describe_by_data( structure ).
    LOOP AT rtts_struct->components ASSIGNING <component>.
      result-view->append( off = offset len = <component>-length ).
      offset = offset + <component>-length.
    ENDLOOP.

  ENDMETHOD.


  METHOD read_zip.

    DATA:
      charset TYPE i.
    FIELD-SYMBOLS:
      <zip_structure_x> TYPE x,
      <zip_structure>   TYPE any.

    ASSIGN zip_structure-ref_to_x->* TO <zip_structure_x>.
    ASSIGN zip_structure-ref_to_structure->* TO <zip_structure>.

    <zip_structure_x> = zip_xstring+offset.

    IF zip_structure-charset_bit >= 1.
      GET BIT zip_structure-charset_bit OF <zip_structure_x> INTO charset.
    ENDIF.

    IF charset = 0.
      IF zip_structure-conv_in_ibm437 IS NOT BOUND.
        zip_structure-conv_in_ibm437 = cl_abap_conv_in_ce=>create(
                  encoding = '1107'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_in_ibm437->convert_struc(
            EXPORTING input = <zip_structure_x>
                      view = zip_structure-view
            IMPORTING data = <zip_structure> ).
    ELSE.
      IF zip_structure-conv_in_utf8 IS NOT BOUND.
        zip_structure-conv_in_utf8 = cl_abap_conv_in_ce=>create(
                  encoding = '4110'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_in_utf8->convert_struc(
            EXPORTING input = <zip_structure_x>
                      view = zip_structure-view
            IMPORTING data = <zip_structure> ).
    ENDIF.

  ENDMETHOD.


  METHOD write_zip.

    DATA:
      charset TYPE i.
    FIELD-SYMBOLS:
      <zip_structure_x> TYPE x,
      <zip_structure>   TYPE any.

    ASSIGN zip_structure-ref_to_x->* TO <zip_structure_x>.
    ASSIGN zip_structure-ref_to_structure->* TO <zip_structure>.

    IF zip_structure-charset_bit >= 1.
      GET BIT zip_structure-charset_bit OF <zip_structure_x> INTO charset.
    ENDIF.

    IF charset = 0.
      IF zip_structure-conv_out_ibm437 IS NOT BOUND.
        zip_structure-conv_out_ibm437 = cl_abap_conv_out_ce=>create(
                  encoding = '1107'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_out_ibm437->convert_struc(
            EXPORTING data = <zip_structure>
                      view = zip_structure-view
            IMPORTING buffer = <zip_structure_x> ).
    ELSE.
      IF zip_structure-conv_out_utf8 IS NOT BOUND.
        zip_structure-conv_out_utf8 = cl_abap_conv_out_ce=>create(
                  encoding = '4110'
                  endian = 'L' ).
      ENDIF.
      zip_structure-conv_out_utf8->convert_struc(
            EXPORTING data = <zip_structure>
                      view = zip_structure-view
            IMPORTING buffer = <zip_structure_x> ).
    ENDIF.

    REPLACE SECTION OFFSET offset LENGTH zip_structure-length OF zip_xstring WITH <zip_structure_x> IN BYTE MODE.

  ENDMETHOD.


ENDCLASS.
