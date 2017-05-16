*&---------------------------------------------------------------------*
*& Report  ZABAP2XLSX_DEMO_SHOW
*&---------------------------------------------------------------------*
REPORT  zabap2xlsx_demo_like_se83.


*----------------------------------------------------------------------*
*       CLASS lcl_perform DEFINITION
*----------------------------------------------------------------------*
CLASS lcl_perform DEFINITION CREATE PRIVATE.
  PUBLIC SECTION.
    CLASS-METHODS: setup_objects,
                   collect_reports,

                   handle_nav FOR EVENT double_click OF cl_gui_alv_grid
                              IMPORTING e_row.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_reports,
             progname TYPE reposrc-progname,
             sort     TYPE reposrc-progname,
             filename TYPE string,
           END OF ty_reports.

    CLASS-DATA:
            lo_grid       TYPE REF TO cl_gui_alv_grid,
            lo_text       TYPE REF TO cl_gui_textedit,
            cl_document   TYPE REF TO i_oi_document_proxy,

            t_reports     TYPE STANDARD TABLE OF ty_reports WITH NON-UNIQUE DEFAULT KEY.
    CLASS-DATA:error         TYPE REF TO i_oi_error,
         t_errors      TYPE STANDARD TABLE OF REF TO i_oi_error WITH NON-UNIQUE DEFAULT KEY,
         cl_control    TYPE REF TO i_oi_container_control.   "Office Dokument

ENDCLASS.                    "lcl_perform DEFINITION


START-OF-SELECTION.
  lcl_perform=>collect_reports( ).
  lcl_perform=>setup_objects( ).

END-OF-SELECTION.

  WRITE '.'.  " Force output


*----------------------------------------------------------------------*
*       CLASS lcl_perform IMPLEMENTATION
*----------------------------------------------------------------------*
CLASS lcl_perform IMPLEMENTATION.
  METHOD setup_objects.
    DATA: lo_split      TYPE REF TO cl_gui_splitter_container,
          lo_container  TYPE REF TO cl_gui_container.

    DATA: it_fieldcat TYPE lvc_t_fcat,
          is_layout   TYPE lvc_s_layo,
          is_variant  TYPE disvariant.
    FIELD-SYMBOLS: <fc> LIKE LINE OF it_fieldcat.


    CREATE OBJECT lo_split
      EXPORTING
        parent                  = cl_gui_container=>screen0
        rows                    = 1
        columns                 = 3
        no_autodef_progid_dynnr = 'X'.
    lo_split->set_column_width(  EXPORTING id                = 1
                                           width             = 20 ).
    lo_split->set_column_width(  EXPORTING id                = 2
                                           width             = 40 ).

* Left:   List of reports
    lo_container = lo_split->get_container( row       = 1
                                            column    = 1 ).

    CREATE OBJECT lo_grid
      EXPORTING
        i_parent = lo_container.
    SET HANDLER lcl_perform=>handle_nav FOR lo_grid.

    is_variant-report = sy-repid.
    is_variant-handle = '0001'.

    is_layout-cwidth_opt = 'X'.

    APPEND INITIAL LINE TO it_fieldcat ASSIGNING <fc>.
    <fc>-fieldname = 'PROGNAME'.
    <fc>-tabname   = 'REPOSRC'.

    APPEND INITIAL LINE TO it_fieldcat ASSIGNING <fc>.
    <fc>-fieldname   = 'SORT'.
    <fc>-ref_field   = 'PROGNAME'.
    <fc>-ref_table   = 'REPOSRC'.


    lo_grid->set_table_for_first_display( EXPORTING
                                            is_variant                    = is_variant
                                            i_save                        = 'A'
                                            is_layout                     = is_layout
                                          CHANGING
                                            it_outtab                     = t_reports
                                            it_fieldcatalog               = it_fieldcat
                                          EXCEPTIONS
                                            invalid_parameter_combination = 1
                                            program_error                 = 2
                                            too_many_lines                = 3
                                            OTHERS                        = 4 ).

* Middle: Text with coding
    lo_container = lo_split->get_container( row       = 1
                                            column    = 2 ).
    CREATE OBJECT lo_text
      EXPORTING
        parent = lo_container.
    lo_text->set_readonly_mode( cl_gui_textedit=>true ).
    lo_text->set_font_fixed( ).



* right:  DemoOutput
    lo_container = lo_split->get_container( row       = 1
                                            column    = 3 ).

    c_oi_container_control_creator=>get_container_control( IMPORTING control = cl_control
                                                                     error   = error ).
    APPEND error TO t_errors.

    cl_control->init_control( EXPORTING  inplace_enabled     = 'X'
                                         no_flush            = 'X'
                                         r3_application_name = 'Demo Document Container'
                                         parent              = lo_container
                              IMPORTING  error               = error
                              EXCEPTIONS OTHERS              = 2 ).
    APPEND error TO t_errors.

    cl_control->get_document_proxy( EXPORTING document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    IMPORTING document_proxy = cl_document
                                              error          = error ).
    APPEND error TO t_errors.
* Errorhandling should be inserted here


  ENDMETHOD.                    "setup_objects

  "collect_reports
  METHOD collect_reports.
    FIELD-SYMBOLS:<report> LIKE LINE OF t_reports.
    DATA: t_source TYPE STANDARD TABLE OF text255 WITH NON-UNIQUE DEFAULT KEY.

* Get all demoreports
    SELECT progname
      INTO CORRESPONDING FIELDS OF TABLE t_reports
      FROM reposrc
      WHERE progname LIKE 'ZDEMO_EXCEL%'
        AND progname <> sy-repid
        AND subc     = '1'.

    LOOP AT t_reports ASSIGNING <report>.

* Check if already switched to new outputoptions
      READ REPORT <report>-progname INTO t_source.
      IF sy-subrc = 0.
        FIND 'INCLUDE zdemo_excel_outputopt_incl.' IN TABLE t_source IGNORING CASE.
      ENDIF.
      IF sy-subrc <> 0.
        DELETE t_reports.
        CONTINUE.
      ENDIF.


* Build half-numeric sort
      <report>-sort = <report>-progname.
      REPLACE REGEX '(ZDEMO_EXCEL)(\d\d)\s*$' IN <report>-sort WITH '$1\0$2'. "      REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE

      REPLACE REGEX '(ZDEMO_EXCEL)(\d)\s*$'      IN <report>-sort WITH '$1\0\0$2'.
    ENDLOOP.
    SORT t_reports BY sort progname.

  ENDMETHOD.  "collect_reports

  METHOD handle_nav.
    CONSTANTS: filename TYPE text80 VALUE 'ZABAP2XLSX_DEMO_SHOW.xlsx'.
    DATA: wa_report   LIKE LINE OF t_reports,
          t_source    TYPE STANDARD TABLE OF text255,
          t_rawdata   TYPE solix_tab,
          wa_rawdata  LIKE LINE OF t_rawdata,
          bytecount   TYPE i,
          length      TYPE i,
          add_selopt  TYPE flag.


    READ TABLE t_reports INTO wa_report INDEX e_row-index.
    CHECK sy-subrc = 0.

* Set new text into middle frame
    READ REPORT wa_report-progname INTO t_source.
    lo_text->set_text_as_r3table( EXPORTING table = t_source ).


* Unload old xls-file
    cl_document->close_document( ).

* Get the demo
* If additional parameters found on selection screen, start via selection screen , otherwise start w/o
    CLEAR add_selopt.
    FIND 'PARAMETERS' IN TABLE t_source.
    IF sy-subrc = 0.
      add_selopt = 'X'.
    ELSE.
      FIND 'SELECT-OPTIONS' IN TABLE t_source.
      IF sy-subrc = 0.
        add_selopt = 'X'.
      ENDIF.
    ENDIF.
    IF add_selopt IS INITIAL.
      SUBMIT (wa_report-progname) AND RETURN                        "#EC CI_SUBMIT
              WITH p_backfn = filename
              WITH rb_back  = 'X'
              WITH rb_down  = ' '
              WITH rb_send  = ' '
              WITH rb_show  = ' '.
    ELSE.
      SUBMIT (wa_report-progname) VIA SELECTION-SCREEN AND RETURN   "#EC CI_SUBMIT
              WITH p_backfn = filename
              WITH rb_back  = 'X'
              WITH rb_down  = ' '
              WITH rb_send  = ' '
              WITH rb_show  = ' '.
    ENDIF.

    OPEN DATASET filename FOR INPUT IN BINARY MODE.
    IF sy-subrc = 0.
      DO.
        CLEAR wa_rawdata.
        READ DATASET filename INTO wa_rawdata LENGTH length.
        IF sy-subrc <> 0.
          APPEND wa_rawdata TO t_rawdata.
          ADD length TO bytecount.
          EXIT.
        ENDIF.
        APPEND wa_rawdata TO t_rawdata.
        ADD length TO bytecount.
      ENDDO.
      CLOSE DATASET filename.
    ENDIF.

    cl_control->get_document_proxy( EXPORTING document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    IMPORTING document_proxy = cl_document
                                              error          = error ).

    cl_document->open_document_from_table( EXPORTING document_size    = bytecount
                                                     document_table   = t_rawdata
                                                     open_inplace     = 'X' ).

  ENDMETHOD.                    "handle_nav

ENDCLASS.                    "lcl_perform IMPLEMENTATION
