*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL15
*&
*&---------------------------------------------------------------------*
*& 2010-10-30, Gregor Wolf:
*& Added the functionality to ouput the read table content
*& 2011-12-19, Shahrin Shahrulzaman:
*& Added the functionality to have multiple input and output files
*&---------------------------------------------------------------------*

REPORT zdemo_excel15.


CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.
    METHODS zif_excel_demo_generator~get_next_generator REDEFINITION.

    CLASS-METHODS class_constructor.

    METHODS get_sub_generator
      RETURNING
        VALUE(result) TYPE REF TO zif_excel_demo_generator
      RAISING
        zcx_excel.

  PRIVATE SECTION.
    TYPES: ty_class_name TYPE c LENGTH 100.
    CLASS-DATA: lt_class_names TYPE TABLE OF ty_class_name.
    DATA: index         TYPE i VALUE 1,
          sub_generator TYPE REF TO zif_excel_demo_generator.

ENDCLASS.


TYPE-POOLS: abap.

DATA: lv_workdir         TYPE string,
      go_excel_generator TYPE REF TO zif_excel_demo_generator.


PARAMETERS: p_path  TYPE zexcel_export_dir,
            p_noout TYPE abap_bool DEFAULT abap_true.


AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( EXPORTING initial_folder  = lv_workdir
                                              CHANGING  selected_folder = lv_workdir ).
  p_path = lv_workdir.


INITIALIZATION.
  cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.


START-OF-SELECTION.

  PERFORM main.



FORM main.

  CONSTANTS: sheet_with_date_formats TYPE string VALUE '24_Sheets_with_different_default_date_formats.xlsx'.

  DATA: excel           TYPE REF TO zcl_excel,
        lo_excel_writer TYPE REF TO zif_excel_writer,
        reader          TYPE REF TO zif_excel_reader.

  DATA: ex  TYPE REF TO zcx_excel,
        msg TYPE string.

  DATA: lv_file      TYPE xstring,
        lv_bytecount TYPE i,
        lt_file_tab  TYPE solix_tab.

  DATA: output_file_path  TYPE string,
        input_file_path   TYPE string,
        lv_file_separator TYPE c.

  DATA: worksheet      TYPE REF TO zcl_excel_worksheet,
        highest_column TYPE zexcel_cell_column,
        highest_row    TYPE int4,
        column         TYPE zexcel_cell_column VALUE 1,
        col_str        TYPE zexcel_cell_column_alpha,
        row            TYPE int4               VALUE 1,
        value          TYPE zexcel_cell_value,
        converted_date TYPE d,
        info           TYPE zif_excel_demo_generator=>ty_information.


  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.
  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = lv_file_separator ).

  CREATE OBJECT go_excel_generator TYPE lcl_excel_generator.

  DO.

    TRY.

        excel = go_excel_generator->generate_excel( ).
        info = go_excel_generator->get_information( ).
        CONCATENATE p_path lv_file_separator info-filename INTO output_file_path.

        IF p_noout EQ abap_false.
          worksheet = excel->get_active_worksheet( ).
          highest_column = worksheet->get_highest_column( ).
          highest_row    = worksheet->get_highest_row( ).

          WRITE: / 'Filename ', info-filename.
          WRITE: / 'Highest column: ', highest_column, 'Highest row: ', highest_row.
          WRITE: /.

          WHILE row <= highest_row.
            WHILE column <= highest_column.
              col_str = zcl_excel_common=>convert_column2alpha( column ).
              worksheet->get_cell(
                EXPORTING
                  ip_column = col_str
                  ip_row    = row
                IMPORTING
                  ep_value = value
              ).
              WRITE: value.
              column = column + 1.
            ENDWHILE.
            WRITE: /.
            column = 1.
            row = row + 1.
          ENDWHILE.
          IF info-filename = sheet_with_date_formats.
            worksheet->get_cell(
              EXPORTING
                ip_column = 'A'
                ip_row = 4
              IMPORTING
                ep_value = value
            ).
            WRITE: / 'Date value using get_cell: ', value.
            converted_date = zcl_excel_common=>excel_string_to_date( ip_value = value ).
            WRITE: / 'Converted date: ', converted_date.
          ENDIF.
        ENDIF.
        CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
        lv_file = lo_excel_writer->write_file( excel ).

        " Convert to binary
        CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
          EXPORTING
            buffer        = lv_file
          IMPORTING
            output_length = lv_bytecount
          TABLES
            binary_tab    = lt_file_tab.

        " Save the file
        cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                          filename     = output_file_path
                                                          filetype     = 'BIN'
                                                 CHANGING data_tab     = lt_file_tab ).


      CATCH zcx_excel INTO ex.    " Exceptions for ABAP2XLSX
        msg = ex->get_text( ).
        WRITE: / msg.
    ENDTRY.


    go_excel_generator = go_excel_generator->get_next_generator( ).
    IF go_excel_generator IS NOT BOUND.
      EXIT.
    ENDIF.

  ENDDO.

ENDFORM.


CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD class_constructor.

    CLEAR lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL1\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL2\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL3\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL4\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL5\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL7\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL8\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL13\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL24\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL31\CLASS=LCL_EXCEL_GENERATOR' TO lt_class_names.

  ENDMETHOD.

  METHOD get_sub_generator.

    DATA: lv_class_name TYPE ty_class_name,
          ex            TYPE REF TO cx_root.

    IF sub_generator IS NOT BOUND.
      READ TABLE lt_class_names INDEX index INTO lv_class_name.
      TRY.
          CREATE OBJECT sub_generator TYPE (lv_class_name).
        CATCH cx_root INTO ex.
          RAISE EXCEPTION TYPE zcx_excel EXPORTING previous = ex.
      ENDTRY.
    ENDIF.

    result = sub_generator.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~get_next_generator.

    DATA: lo_excel_generator TYPE REF TO lcl_excel_generator.

    IF index >= lines( lt_class_names ).

      CLEAR result.

    ELSE.

      CREATE OBJECT lo_excel_generator.
      lo_excel_generator->index = index + 1.

      result = lo_excel_generator.

    ENDIF.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~get_information.

    DATA: lo_excel_generator TYPE REF TO zif_excel_demo_generator.

    lo_excel_generator = get_sub_generator( ).

    result = lo_excel_generator->get_information( ).
    result-program = sy-repid.
    result-objid = |{ sy-repid }_{ result-filename(2) }|.
    result-filename = |15_{ replace( val = result-filename sub = '.xlsx' with = 'FromReader.xlsx' ) }|.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

    DATA: lo_excel_generator TYPE REF TO zif_excel_demo_generator.

    lo_excel_generator = get_sub_generator( ).

    result = lo_excel_generator->generate_excel( ).

  ENDMETHOD.

ENDCLASS.
