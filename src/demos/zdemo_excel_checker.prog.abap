*&---------------------------------------------------------------------*
*& Report zdemo_excel_checker
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdemo_excel_checker.

CLASS lcl_app DEFINITION.

  PUBLIC SECTION.

    METHODS set_sscrfields
      CHANGING
        sscrfields TYPE sscrfields.

    METHODS at_selection_screen_output.

    METHODS at_selection_screen.

  PRIVATE SECTION.

    TYPES : BEGIN OF ty_check_result,
              objid                  TYPE wwwdata-objid,
              info                   TYPE zif_excel_demo_generator=>ty_information,
              diff                   TYPE abap_bool,
              xlsx_just_now          TYPE xstring,
              xlsx_reference         TYPE xstring,
              compare_xlsx_just_now  TYPE REF TO cl_abap_zip,
              compare_xlsx_reference TYPE REF TO cl_abap_zip,
            END OF ty_check_result,
            BEGIN OF ty_alv_line,
              status_icon            TYPE string,
              xlsx_diff              TYPE string,
              write_smw0             TYPE string,
              program                TYPE trdir-name,
              prog_text              TYPE trdirt-text,
              objid                  TYPE wwwdata-objid,
              obj_text               TYPE wwwdata-text,
              filename               TYPE string,
              xlsx_just_now          TYPE xstring,
              xlsx_reference         TYPE xstring,
              compare_xlsx_just_now  TYPE REF TO cl_abap_zip,
              compare_xlsx_reference TYPE REF TO cl_abap_zip,
              cell_types             TYPE salv_t_int4_column,
            END OF ty_alv_line,
            ty_alv_table TYPE STANDARD TABLE OF ty_alv_line WITH DEFAULT KEY.

    METHODS check_regression
      IMPORTING
        lo_excel_generator TYPE REF TO zif_excel_demo_generator
      RETURNING
        VALUE(result)      TYPE ty_check_result
      RAISING
        zcx_excel.

    METHODS read_xlsx_from_web_repository
      IMPORTING
        objid         TYPE wwwdata-objid
      RETURNING
        VALUE(result) TYPE xstring.

    METHODS write_xlsx_to_web_repository
      IMPORTING
        objid    TYPE wwwdata-objid
        text     TYPE wwwdata-text
        xstring  TYPE xstring
        filename TYPE clike.

    METHODS load_alv_table
      RAISING
        zcx_excel.

    METHODS on_link_clicked FOR EVENT link_click OF cl_salv_events_table IMPORTING column row.

    DATA: ref_sscrfields     TYPE REF TO sscrfields,
          splitter           TYPE REF TO cl_gui_splitter_container,
          alv_container      TYPE REF TO cl_gui_container,
          zip_diff_container TYPE REF TO cl_gui_container,
          viewer             TYPE REF TO object,
          salv               TYPE REF TO cl_salv_table,
          alv_table          TYPE ty_alv_table.

ENDCLASS.



CLASS lcl_app IMPLEMENTATION.


  METHOD set_sscrfields.

    ref_sscrfields = REF #( sscrfields ).

  ENDMETHOD.


  METHOD at_selection_screen_output.

    DATA: lt_itab TYPE ui_functions,
          columns TYPE REF TO cl_salv_columns_table,
          events  TYPE REF TO cl_salv_events_table,
          lx      TYPE REF TO cx_root.

    LOOP AT SCREEN.
      screen-active = '0'.
      MODIFY SCREEN.
    ENDLOOP.

    TRY.

        ref_sscrfields->functxt_01 = icon_refresh.
        APPEND 'ONLI' TO lt_itab.
        APPEND 'PRIN' TO lt_itab.
        APPEND 'SPOS' TO lt_itab.
        CALL FUNCTION 'RS_SET_SELSCREEN_STATUS'
          EXPORTING
            p_status  = sy-pfkey
          TABLES
            p_exclude = lt_itab.

        load_alv_table( ).

        IF alv_container IS NOT BOUND.

          CREATE OBJECT splitter
            EXPORTING
              parent            = cl_gui_container=>screen0
              rows              = 1
              columns           = 2
            EXCEPTIONS
              cntl_error        = 1
              cntl_system_error = 2
              OTHERS            = 3.
          IF sy-subrc <> 0.
* MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
          ENDIF.

          alv_container = splitter->get_container( row = 1 column = 1 ).
          zip_diff_container = splitter->get_container( row = 1 column = 2 ).

          cl_salv_table=>factory(
            EXPORTING
              r_container    = alv_container
            IMPORTING
              r_salv_table   = salv
            CHANGING
              t_table        = alv_table ).

          columns = salv->get_columns( ).
          columns->set_cell_type_column( value = 'CELL_TYPES' ).
          columns->get_column( 'STATUS_ICON' )->set_output_length( 2 ).
          columns->get_column( 'XLSX_DIFF' )->set_output_length( 5 ).
          columns->get_column( 'XLSX_DIFF' )->set_alignment( if_salv_c_alignment=>centered ).
          columns->get_column( 'WRITE_SMW0' )->set_output_length( 5 ).
          columns->get_column( 'WRITE_SMW0' )->set_alignment( if_salv_c_alignment=>centered ).
          columns->get_column( 'PROGRAM' )->set_output_length( 15 ).
          columns->get_column( 'PROG_TEXT' )->set_output_length( 30 ).
          columns->get_column( 'OBJID' )->set_output_length( 20 ).
          columns->get_column( 'OBJ_TEXT' )->set_output_length( 50 ).
          columns->get_column( 'XLSX_JUST_NOW' )->set_visible( if_salv_c_bool_sap=>false ).
          columns->get_column( 'XLSX_JUST_NOW' )->set_alignment( if_salv_c_alignment=>centered ).
          columns->get_column( 'XLSX_REFERENCE' )->set_visible( if_salv_c_bool_sap=>false ).


          events = salv->get_event( ).
          SET HANDLER on_link_clicked FOR events.

          salv->display( ).

        ELSE.



        ENDIF.

      CATCH cx_root INTO lx.
        MESSAGE lx TYPE 'I' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD at_selection_screen.

    DATA: lx TYPE REF TO cx_root.

    TRY.

        CASE ref_sscrfields->ucomm.

          WHEN 'FC01'. " REFRESH

            " restart the program completely so that to consider modification/recompile of any ZDEMO_EXCEL program
            SUBMIT (sy-repid) VIA SELECTION-SCREEN.

        ENDCASE.

      CATCH cx_root INTO lx.
        MESSAGE lx TYPE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD load_alv_table.

    DATA: class_names        TYPE string_table,
          class_name         TYPE string,
          lo_excel_generator TYPE REF TO zif_excel_demo_generator,
          alv_line           TYPE lcl_app=>ty_alv_line,
          cell_type          TYPE salv_s_int4_column,
          lx                 TYPE REF TO cx_root,
          check_result       TYPE lcl_app=>ty_check_result,
          message            TYPE string.

    CLEAR class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL1\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL2\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL3\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL4\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL5\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL6\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL7\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL8\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL9\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL10\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL12\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL13\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL14\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL15\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL16\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL17\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL18\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL19\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL21\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL22\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL23\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL24\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL27\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL30\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL31\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL33\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL34\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL35\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL36\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL38\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL39\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL40\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZDEMO_EXCEL_COMMENTS\CLASS=LCL_EXCEL_GENERATOR' TO class_names.
    APPEND '\PROGRAM=ZTEST_EXCEL_IMAGE_HEADER\CLASS=LCL_EXCEL_GENERATOR' TO class_names.

    zcl_excel_demo_generator=>set_date_now( date = '20211204' ).
    zcl_excel_demo_generator=>set_time_now( time = '112359' ).

    CLEAR alv_table.
    LOOP AT class_names INTO class_name.

      TRY.
          CREATE OBJECT lo_excel_generator TYPE (class_name).
        CATCH cx_root INTO lx.
          RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |Generator can't be instantiated|.
      ENDTRY.

      DO.

        TRY.

            TRY.
                lo_excel_generator->checker_initialization( ).
              CATCH cx_root INTO lx.
            ENDTRY.

            check_result = check_regression( lo_excel_generator ).

            CLEAR alv_line.
            CASE check_result-diff.
              WHEN abap_true.
                alv_line-status_icon = icon_cancel.
              WHEN abap_false.
                alv_line-status_icon = icon_okay.
            ENDCASE.
            alv_line-program = check_result-info-program.
            SELECT SINGLE text FROM trdirt INTO alv_line-prog_text
                WHERE sprsl = sy-langu
                  AND name  = alv_line-program.
            alv_line-objid = check_result-objid.
            alv_line-obj_text = |{ check_result-info-filename } ({ check_result-info-program })|.
            IF check_result-diff = abap_true.
              alv_line-xlsx_diff = '@46\QShow differences@'.
              alv_line-write_smw0 = '@2L\QSave XLSX to Web Repository@'.
              cell_type-columnname = 'XLSX_DIFF'.
              cell_type-value      = if_salv_c_cell_type=>hotspot.
              APPEND cell_type TO alv_line-cell_types.
              cell_type-columnname = 'WRITE_SMW0'.
              cell_type-value      = if_salv_c_cell_type=>hotspot.
              APPEND cell_type TO alv_line-cell_types.
            ENDIF.
            alv_line-xlsx_just_now = check_result-xlsx_just_now.
            alv_line-xlsx_reference = check_result-xlsx_reference.
            alv_line-compare_xlsx_just_now = check_result-compare_xlsx_just_now.
            alv_line-compare_xlsx_reference = check_result-compare_xlsx_reference.
            APPEND alv_line TO alv_table.

          CATCH cx_root INTO lx.
            message = |{ class_name }: { lx->get_text( ) }|.
            RAISE EXCEPTION TYPE zcx_excel EXPORTING error = message.
        ENDTRY.

        lo_excel_generator = lo_excel_generator->get_next_generator( ).
        IF lo_excel_generator IS NOT BOUND.
          EXIT.
        ENDIF.

      ENDDO.

    ENDLOOP.

  ENDMETHOD.


  METHOD check_regression.

    DATA: lo_excel  TYPE REF TO zcl_excel,
          lo_writer TYPE REF TO zcl_excel_writer_2007,
          diff      TYPE REF TO object.
    FIELD-SYMBOLS: <is_different> TYPE abap_bool.

    "=========================
    " ASK XLSX TO DEMO PROGRAM
    "=========================

    lo_excel = lo_excel_generator->generate_excel( ).
    result-info = lo_excel_generator->get_information( ).
    IF result-info-objid IS NOT INITIAL.
      result-objid = result-info-objid.
    ELSE.
      result-objid = result-info-program.
    ENDIF.

    CREATE OBJECT lo_writer.
    result-xlsx_just_now = lo_writer->zif_excel_writer~write_file( lo_excel ).

    "=========================
    " READ REFERENCE XLSX FROM WEB REPOSITORY
    "=========================

    result-xlsx_reference = read_xlsx_from_web_repository( objid = result-objid ).

    "=========================
    " COMPARE
    "=========================
    IF result-xlsx_reference IS INITIAL.

      result-diff = abap_true.

    ELSE.

      result-compare_xlsx_just_now = lo_excel_generator->cleanup_for_diff( result-xlsx_just_now ).
      result-compare_xlsx_reference = lo_excel_generator->cleanup_for_diff( result-xlsx_reference ).

      CALL METHOD ('ZCL_ZIP_DIFF_ITEM')=>('GET_DIFF')
        EXPORTING
          zip_1  = result-compare_xlsx_reference
          zip_2  = result-compare_xlsx_just_now
        RECEIVING
          result = diff.

      ASSIGN ('DIFF->IS_DIFFERENT') TO <is_different>.
      IF sy-subrc = 0.
        result-diff = <is_different>.
      ENDIF.

    ENDIF.

  ENDMETHOD.


  METHOD on_link_clicked.

    DATA: alv_line       TYPE ty_alv_line,
          message        TYPE string,
          lx             TYPE REF TO cx_root,
          refresh_stable TYPE lvc_s_stbl.

    READ TABLE alv_table INDEX row INTO alv_line.

    CASE column.

      WHEN 'XLSX_DIFF'.

        TRY.
            IF viewer IS NOT BOUND.
              CREATE OBJECT viewer TYPE ('ZCL_ZIP_DIFF_VIEWER2')
                  EXPORTING
                    io_container = zip_diff_container.
            ENDIF.

            CALL METHOD viewer->('DIFF_AND_VIEW')
              EXPORTING
                zip_old = alv_line-compare_xlsx_reference
                zip_new = alv_line-compare_xlsx_just_now.

          CATCH cx_root INTO lx.
            message = |Viewer error (https://github.com/sandraros/zip-diff): { lx->get_text( ) }|.
            MESSAGE message TYPE 'I' DISPLAY LIKE 'E'.
        ENDTRY.

      WHEN 'WRITE_SMW0'.

        write_xlsx_to_web_repository(
            objid    = alv_line-objid
            text     = alv_line-obj_text
            xstring  = alv_line-xlsx_just_now
            filename = alv_line-filename ).
        COMMIT WORK.

        alv_line-status_icon = icon_okay.
        CLEAR alv_line-xlsx_diff.
        CLEAR alv_line-write_smw0.
        CLEAR alv_line-cell_types.
        MODIFY alv_table INDEX row FROM alv_line.

        refresh_stable-row = abap_true.
        refresh_stable-col = abap_true.
        salv->refresh( s_stable = refresh_stable ).

    ENDCASE.

  ENDMETHOD.


  METHOD read_xlsx_from_web_repository.

    DATA: query_string   TYPE w3query,
          query_table    TYPE TABLE OF w3query,
          html_table     TYPE TABLE OF w3html,
          return_code    TYPE w3param-ret_code,
          content_type   TYPE w3param-cont_type,
          content_length TYPE w3param-cont_len,
          mime_table     TYPE TABLE OF w3mime.

    CLEAR: query_table, query_string.
    query_string-name = '_OBJECT_ID'.
    query_string-value = objid.
    APPEND query_string TO query_table.

    CALL FUNCTION 'WWW_GET_MIME_OBJECT'
      TABLES
        query_string        = query_table
        html                = html_table
        mime                = mime_table
      CHANGING
        return_code         = return_code
        content_type        = content_type
        content_length      = content_length
      EXCEPTIONS
        object_not_found    = 1
        parameter_not_found = 2
        OTHERS              = 3.
    IF sy-subrc <> 0.
      RETURN.
*   MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    result = cl_bcs_convert=>solix_to_xstring( it_solix = mime_table iv_size = content_length ).

  ENDMETHOD.


  METHOD write_xlsx_to_web_repository.

    DATA: key                TYPE wwwdatatab,
          mime               TYPE TABLE OF w3mime,
          devclass           TYPE devclass,
          wwwparam           TYPE wwwparams,
          table_of_wwwparams TYPE TABLE OF wwwparams.

    SELECT SINGLE devclass FROM tadir
        INTO devclass
        WHERE pgmid    = 'R3TR'
          AND object   = 'W3MI'
          AND obj_name = objid.

    IF sy-subrc <> 0.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = |Object must be first created manually|.
    ENDIF.

    mime = cl_bcs_convert=>xstring_to_solix( xstring ).
    key-relid = 'MI'.
    key-objid = objid.
    key-chname = sy-uname.
    key-text = text.
    key-tdate = sy-datum.
    key-ttime = sy-uzeit.

    CALL FUNCTION 'WWWDATA_EXPORT'
      EXPORTING
        key               = key
      TABLES
        mime              = mime
      EXCEPTIONS
        wrong_object_type = 1
        export_error      = 2
        OTHERS            = 3.
    IF sy-subrc <> 0.
* MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    wwwparam-relid = 'MI'.
    wwwparam-objid = objid.
    wwwparam-name = 'mimetype'.
    wwwparam-value = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'filename'.
    wwwparam-value = filename.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'filesize'.
    wwwparam-value = |{ xstrlen( xstring ) }|.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'version'.
    wwwparam-value = |00001|.
    APPEND wwwparam TO table_of_wwwparams.
    wwwparam-name = 'fileextension'.
    wwwparam-value = |.xlsx|.
    APPEND wwwparam TO table_of_wwwparams.

    CALL FUNCTION 'WWWPARAMS_UPDATE'
      TABLES
        params       = table_of_wwwparams
      EXCEPTIONS
        update_error = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
*   MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

  ENDMETHOD.


ENDCLASS.



TABLES sscrfields.
DATA: app TYPE REF TO lcl_app.

PARAMETERS dummy.
SELECTION-SCREEN FUNCTION KEY 1.

INITIALIZATION.
  CREATE OBJECT app.
  app->set_sscrfields( CHANGING sscrfields = sscrfields ).

AT SELECTION-SCREEN OUTPUT.
  app->at_selection_screen_output( ).

AT SELECTION-SCREEN.
  app->at_selection_screen( ).
