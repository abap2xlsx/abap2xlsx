*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL25
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel25.

TYPES: BEGIN OF ty_f4_path,
         pathintern TYPE filepath-pathintern,
         pathname   TYPE pathtext-pathname,
         pathextern TYPE path-pathextern,
       END OF ty_f4_path.

DATA: lt_r_fldval TYPE RANGE OF filepath-pathintern,
      lt_value    TYPE TABLE OF ty_f4_path,
      ls_value    TYPE ty_f4_path.

PARAMETERS log_path TYPE filepath-pathintern DEFAULT 'LOCAL_TEMPORARY_FILES'.
SELECTION-SCREEN COMMENT /35(83) physpath.
PARAMETERS filename TYPE string LOWER CASE DEFAULT '25_HelloWorld.xlsx'.
PARAMETERS param_1 TYPE string LOWER CASE.
PARAMETERS param_2 TYPE string LOWER CASE.

AT SELECTION-SCREEN OUTPUT.

  PERFORM read_file_paths.
  READ TABLE lt_value WITH KEY pathintern = log_path INTO ls_value.
  physpath = ls_value-pathextern.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR log_path.

  DATA: lt_return    TYPE TABLE OF ddshretval,
        ls_return    TYPE ddshretval,
        lt_dynpfield TYPE TABLE OF dynpread,
        ls_dynpfield TYPE dynpread.

  CALL FUNCTION 'DYNP_VALUES_READ'
    EXPORTING
      dyname     = sy-repid
      dynumb     = sy-dynnr
      request    = 'A' " read all screen fields
    TABLES
      dynpfields = lt_dynpfield
    EXCEPTIONS
      OTHERS     = 9.

  READ TABLE lt_dynpfield WITH KEY fieldname = 'LOG_PATH' INTO ls_dynpfield.
  log_path = ls_dynpfield-fieldvalue.

  PERFORM read_file_paths.

  CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'
    EXPORTING
      value_org       = 'S'
      multiple_choice = ' '
      retfield        = 'PATHINTERN'
    TABLES
      value_tab       = lt_value
      return_tab      = lt_return
    EXCEPTIONS
      OTHERS          = 0.

  IF lt_return IS INITIAL.
    RETURN.
  ENDIF.

  READ TABLE lt_return INDEX 1 INTO ls_return.
  READ TABLE lt_value WITH KEY pathintern = ls_return-fieldval INTO ls_value.

  DELETE lt_dynpfield WHERE fieldname = 'LOG_PATH' OR fieldname = 'PHYSPATH'.
  ls_dynpfield-fieldname = 'LOG_PATH'.
  ls_dynpfield-fieldvalue = ls_value-pathintern.
  APPEND ls_dynpfield TO lt_dynpfield.
  ls_dynpfield-fieldname = 'PHYSPATH'.
  ls_dynpfield-fieldvalue = ls_value-pathextern.
  APPEND ls_dynpfield TO lt_dynpfield.

  CALL FUNCTION 'DYNP_VALUES_UPDATE'
    EXPORTING
      dyname     = sy-repid
      dynumb     = sy-dynnr
    TABLES
      dynpfields = lt_dynpfield
    EXCEPTIONS
      OTHERS     = 8.

FORM read_file_paths.

  DATA: ls_r_fldval LIKE LINE OF lt_r_fldval.

  CLEAR lt_r_fldval.
  IF log_path CA '*'.
    ls_r_fldval-sign = 'I'.
    ls_r_fldval-option = 'CP'.
    ls_r_fldval-low  = log_path.
    APPEND ls_r_fldval TO lt_r_fldval.
  ENDIF.

  SELECT filepath~pathintern pathtext~pathname path~pathextern
        FROM filepath
        INNER JOIN path ON path~pathintern = filepath~pathintern
        INNER JOIN opsystem ON opsystem~filesys = path~filesys AND opsystem~opsys = sy-opsys
        LEFT JOIN pathtext ON pathtext~pathintern = filepath~pathintern AND pathtext~language = sy-langu
    INTO TABLE lt_value
    WHERE filepath~pathintern IN lt_r_fldval.

ENDFORM.

START-OF-SELECTION.

DATA: lo_excel             TYPE REF TO zcl_excel.
DATA: lo_excel_writer      TYPE REF TO zif_excel_writer.
DATA: lo_worksheet         TYPE REF TO zcl_excel_worksheet.
DATA: lo_exception         TYPE REF TO cx_root.
DATA: lv_file              TYPE xstring.
DATA: lv_default_file_name TYPE string.
DATA: lv_default_file_name2 TYPE c LENGTH 255.
DATA: lv_error             TYPE string.

CALL FUNCTION 'FILE_GET_NAME_USING_PATH'
  EXPORTING
    logical_path        = log_path
    file_name           = filename
    parameter_1         = param_1
    parameter_2         = param_2
  IMPORTING
    file_name_with_path = lv_default_file_name    " File name with path
  EXCEPTIONS
    other               = 1.
IF sy-subrc <> 0.
  MESSAGE ID sy-msgid TYPE 'I' NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
ENDIF.

" Creates active sheet
CREATE OBJECT lo_excel.

" Get active sheet
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( ip_title = 'Sheet1' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).

CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
lv_file = lo_excel_writer->write_file( lo_excel ).

TRY.
    OPEN DATASET lv_default_file_name FOR OUTPUT IN BINARY MODE.
    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE cx_sy_file_open
        EXPORTING
          filename  = lv_default_file_name
          errorcode = sy-subrc
          errortext = |Cannot create or open file - Check Tx FILE Logical Path 'LOCAL_TEMPORARY_FILES'|.
    ENDIF.
    TRANSFER lv_file  TO lv_default_file_name.

    CLOSE DATASET lv_default_file_name.
  CATCH cx_root INTO lo_exception.
    lv_error = lo_exception->get_text( ).
    MESSAGE lv_error TYPE 'I'.
    STOP.
ENDTRY.

lv_default_file_name2 = lv_default_file_name.
SET PARAMETER ID 'GR8' FIELD lv_default_file_name2.
SUBMIT zdemo_excel37 VIA SELECTION-SCREEN WITH p_applse = abap_true AND RETURN.
