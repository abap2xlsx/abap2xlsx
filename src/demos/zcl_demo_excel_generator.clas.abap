CLASS zcl_demo_excel_generator DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES zif_demo_excel_generator.
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



CLASS zcl_demo_excel_generator IMPLEMENTATION.

  METHOD zif_demo_excel_generator~checker_initialization.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~generate_excel.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~get_information.

  ENDMETHOD.

  METHOD zif_demo_excel_generator~get_next_generator.

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

  METHOD zif_demo_excel_generator~cleanup_for_diff.

    DATA: zip TYPE REF TO cl_abap_zip.

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
        content                 = DATA(content)
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

    result = zip.

  ENDMETHOD.

ENDCLASS.
