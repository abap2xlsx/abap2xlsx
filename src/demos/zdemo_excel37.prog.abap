REPORT zdemo_excel37.

TYPE-POOLS: vrm.

DATA: excel                 TYPE REF TO zcl_excel,
      reader                TYPE REF TO zif_excel_reader,
      go_error              TYPE REF TO cx_root,
      gv_memid_gr8          TYPE text255,
      gv_message            TYPE string,
      lv_extension          TYPE string,
      gv_error_program_name TYPE syrepid,
      gv_error_include_name TYPE syrepid,
      gv_error_line         TYPE i.

DATA: gc_save_file_name TYPE string VALUE '37- Read template and output.&'.

SELECTION-SCREEN BEGIN OF BLOCK blx WITH FRAME.
PARAMETERS: p_upfile TYPE string LOWER CASE MEMORY ID gr8.
SELECTION-SCREEN END OF BLOCK blx.

INCLUDE zdemo_excel_outputopt_incl.

SELECTION-SCREEN BEGIN OF BLOCK cls WITH FRAME TITLE text-cls.
PARAMETERS: lb_read   TYPE seoclsname AS LISTBOX VISIBLE LENGTH 40 LOWER CASE OBLIGATORY DEFAULT 'Autodetect'(001).
PARAMETERS: lb_write  TYPE seoclsname AS LISTBOX VISIBLE LENGTH 40 LOWER CASE OBLIGATORY DEFAULT 'Autodetect'(001).
SELECTION-SCREEN END OF BLOCK cls.

SELECTION-SCREEN BEGIN OF BLOCK bl_err WITH FRAME TITLE text-err.
PARAMETERS: cb_errl AS CHECKBOX DEFAULT 'X'.
SELECTION-SCREEN BEGIN OF LINE.
PARAMETERS: cb_dump AS CHECKBOX DEFAULT space.
SELECTION-SCREEN COMMENT (60) cmt_dump FOR FIELD cb_dump.
SELECTION-SCREEN END OF LINE.
SELECTION-SCREEN END OF BLOCK bl_err.

INITIALIZATION.
  PERFORM setup_listboxes.
  cmt_dump = text-dum.
  GET PARAMETER ID 'GR8' FIELD gv_memid_gr8.
  p_upfile = gv_memid_gr8.

  IF p_upfile IS INITIAL.
    p_upfile = 'c:\temp\whatever.xlsx'.
  ENDIF.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_upfile.
  PERFORM f4_p_upfile CHANGING p_upfile.


START-OF-SELECTION.
  IF cb_dump IS INITIAL.
    TRY.
        PERFORM read_template.
        PERFORM write_template.
*** Create output
      CATCH cx_root INTO go_error.
        MESSAGE 'Error reading excelfile' TYPE 'I'.
        gv_message = go_error->get_text( ).
        IF cb_errl = ' '.
          IF gv_message IS NOT INITIAL.
            MESSAGE gv_message TYPE 'I'.
          ENDIF.
        ELSE.
          go_error->get_source_position( IMPORTING program_name = gv_error_program_name
                                                   include_name = gv_error_include_name
                                                   source_line  = gv_error_line         ).
          WRITE:/ 'Errormessage:'       ,gv_message.
          WRITE:/ 'Errorposition:',
                AT /10 'Program:'       ,gv_error_program_name,
                AT /10 'include_name:'  ,gv_error_include_name,
                AT /10 'source_line:'   ,gv_error_line.
        ENDIF.
    ENDTRY.
  ELSE.  " This will dump if an error occurs.  In some cases the information given in cx_root is not helpful - this will show exactly where the problem is
    PERFORM read_template.
    PERFORM write_template.
  ENDIF.



*&---------------------------------------------------------------------*
*&      Form  F4_P_UPFILE
*&---------------------------------------------------------------------*
FORM f4_p_upfile  CHANGING p_upfile TYPE string.

  DATA: lv_repid       TYPE syrepid,
        lt_fields      TYPE dynpread_tabtype,
        ls_field       LIKE LINE OF lt_fields,
        lt_files       TYPE filetable,
        lv_file_filter TYPE string.

  lv_repid = sy-repid.

  CALL FUNCTION 'DYNP_VALUES_READ'
    EXPORTING
      dyname               = lv_repid
      dynumb               = '1000'
      request              = 'A'
    TABLES
      dynpfields           = lt_fields
    EXCEPTIONS
      invalid_abapworkarea = 01
      invalid_dynprofield  = 02
      invalid_dynproname   = 03
      invalid_dynpronummer = 04
      invalid_request      = 05
      no_fielddescription  = 06
      undefind_error       = 07.
  READ TABLE lt_fields INTO ls_field WITH KEY fieldname = 'P_UPFILE'.
  p_upfile = ls_field-fieldvalue.

  lv_file_filter = 'Excel Files (*.XLSX;*.XLSM)|*.XLSX;*.XLSM'.
  cl_gui_frontend_services=>file_open_dialog( EXPORTING
                                                default_filename        = p_upfile
                                                file_filter             = lv_file_filter
                                              CHANGING
                                                file_table              = lt_files
                                                rc                      = sy-tabix
                                              EXCEPTIONS
                                                OTHERS                  = 1 ).
  READ TABLE lt_files INDEX 1 INTO p_upfile.

ENDFORM.                    " F4_P_UPFILE


*&---------------------------------------------------------------------*
*&      Form  SETUP_LISTBOXES
*&---------------------------------------------------------------------*
FORM setup_listboxes .

  DATA: lv_id                   TYPE vrm_id,
        lt_values               TYPE vrm_values,
        lt_implementing_classes TYPE seo_relkeys.

  FIELD-SYMBOLS: <ls_implementing_class> LIKE LINE OF lt_implementing_classes,
                 <ls_value>              LIKE LINE OF lt_values.

*--------------------------------------------------------------------*
* Possible READER-Classes
*--------------------------------------------------------------------*
  lv_id = 'LB_READ'.
  APPEND INITIAL LINE TO lt_values ASSIGNING <ls_value>.
  <ls_value>-key  = 'Autodetect'(001).
  <ls_value>-text = 'Autodetect'(001).


  PERFORM get_implementing_classds USING    'ZIF_EXCEL_READER'
                                   CHANGING lt_implementing_classes.
  CLEAR lt_values.
  LOOP AT lt_implementing_classes ASSIGNING <ls_implementing_class>.

    APPEND INITIAL LINE TO lt_values ASSIGNING <ls_value>.
    <ls_value>-key  = <ls_implementing_class>-clsname.
    <ls_value>-text = <ls_implementing_class>-clsname.

  ENDLOOP.

  CALL FUNCTION 'VRM_SET_VALUES'
    EXPORTING
      id              = lv_id
      values          = lt_values
    EXCEPTIONS
      id_illegal_name = 1
      OTHERS          = 2.

*--------------------------------------------------------------------*
* Possible WRITER-Classes
*--------------------------------------------------------------------*
  lv_id = 'LB_WRITE'.
  APPEND INITIAL LINE TO lt_values ASSIGNING <ls_value>.
  <ls_value>-key  = 'Autodetect'(001).
  <ls_value>-text = 'Autodetect'(001).


  PERFORM get_implementing_classds USING    'ZIF_EXCEL_WRITER'
                                   CHANGING lt_implementing_classes.
  CLEAR lt_values.
  LOOP AT lt_implementing_classes ASSIGNING <ls_implementing_class>.

    APPEND INITIAL LINE TO lt_values ASSIGNING <ls_value>.
    <ls_value>-key  = <ls_implementing_class>-clsname.
    <ls_value>-text = <ls_implementing_class>-clsname.

  ENDLOOP.

  CALL FUNCTION 'VRM_SET_VALUES'
    EXPORTING
      id              = lv_id
      values          = lt_values
    EXCEPTIONS
      id_illegal_name = 1
      OTHERS          = 2.

ENDFORM.                    " SETUP_LISTBOXES


*&---------------------------------------------------------------------*
*&      Form  GET_IMPLEMENTING_CLASSDS
*&---------------------------------------------------------------------*
FORM get_implementing_classds  USING    iv_interface_name       TYPE clike
                               CHANGING ct_implementing_classes TYPE seo_relkeys.

  DATA: lo_oo_interface            TYPE REF TO cl_oo_interface,
        lo_oo_class                TYPE REF TO cl_oo_class,
        lt_implementing_subclasses TYPE seo_relkeys.

  FIELD-SYMBOLS: <ls_implementing_class> LIKE LINE OF ct_implementing_classes.

  TRY.
      lo_oo_interface ?= cl_oo_interface=>get_instance( iv_interface_name ).
    CATCH cx_class_not_existent.
      RETURN.
  ENDTRY.
  ct_implementing_classes = lo_oo_interface->get_implementing_classes( ).

  LOOP AT ct_implementing_classes ASSIGNING <ls_implementing_class>.
    TRY.
        lo_oo_class ?= cl_oo_class=>get_instance( <ls_implementing_class>-clsname ).
        lt_implementing_subclasses = lo_oo_class->get_subclasses( ).
        APPEND LINES OF lt_implementing_subclasses TO ct_implementing_classes.
      CATCH cx_class_not_existent.
    ENDTRY.
  ENDLOOP.


ENDFORM.                    " GET_IMPLEMENTING_CLASSDS


*&---------------------------------------------------------------------*
*&      Form  READ_TEMPLATE
*&---------------------------------------------------------------------*
FORM read_template RAISING zcx_excel .

  CASE lb_read.
    WHEN 'Autodetect'(001).
      FIND REGEX '(\.xlsx|\.xlsm)\s*$' IN p_upfile SUBMATCHES lv_extension.
      TRANSLATE lv_extension TO UPPER CASE.
      CASE lv_extension.

        WHEN '.XLSX'.
          CREATE OBJECT reader TYPE zcl_excel_reader_2007.
          excel = reader->load_file(  p_upfile ).
          "Use template for charts
          excel->use_template = abap_true.

        WHEN '.XLSM'.
          CREATE OBJECT reader TYPE zcl_excel_reader_xlsm.
          excel = reader->load_file(  p_upfile ).
          "Use template for charts
          excel->use_template = abap_true.

        WHEN OTHERS.
          MESSAGE 'Unsupported filetype' TYPE 'I'.
          RETURN.

      ENDCASE.

    WHEN OTHERS.
      CREATE OBJECT reader TYPE (lb_read).
      excel = reader->load_file(  p_upfile ).
      "Use template for charts
      excel->use_template = abap_true.

  ENDCASE.

ENDFORM.                    " READ_TEMPLATE


*&---------------------------------------------------------------------*
*&      Form  WRITE_TEMPLATE
*&---------------------------------------------------------------------*
FORM write_template  RAISING zcx_excel.

  CASE lb_write.

    WHEN 'Autodetect'(001).
      FIND REGEX '(\.xlsx|\.xlsm)\s*$' IN p_upfile SUBMATCHES lv_extension.
      TRANSLATE lv_extension TO UPPER CASE.
      CASE lv_extension.

        WHEN '.XLSX'.
          REPLACE '&' IN gc_save_file_name WITH 'xlsx'.  " Pass extension for standard writer
          lcl_output=>output( excel ).

        WHEN '.XLSM'.
          REPLACE '&' IN gc_save_file_name WITH 'xlsm'.  " Pass extension for macro-writer
          lcl_output=>output( cl_excel            = excel
                              iv_writerclass_name = 'ZCL_EXCEL_WRITER_XLSM' ).

        WHEN OTHERS.
          MESSAGE 'Unsupported filetype' TYPE 'I'.
          RETURN.

      ENDCASE.

    WHEN OTHERS.
      lcl_output=>output( cl_excel            = excel
                          iv_writerclass_name = lb_write ).
  ENDCASE.

ENDFORM.                    " WRITE_TEMPLATE
