REPORT zdemo_excel38.


CLASS lcl_excel_generator DEFINITION INHERITING FROM zcl_excel_demo_generator.

  PUBLIC SECTION.
    METHODS zif_excel_demo_generator~checker_initialization REDEFINITION.
    METHODS zif_excel_demo_generator~get_information REDEFINITION.
    METHODS zif_excel_demo_generator~generate_excel REDEFINITION.

ENDCLASS.

DATA: lo_excel           TYPE REF TO zcl_excel,
      lo_excel_generator TYPE REF TO lcl_excel_generator.

CONSTANTS: gc_save_file_name TYPE string VALUE '38_SAP-Icons.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


TABLES: icon.
SELECT-OPTIONS: s_icon FOR icon-name DEFAULT 'ICON_LED_*' OPTION CP.

START-OF-SELECTION.

  CREATE OBJECT lo_excel_generator.
  lo_excel = lo_excel_generator->zif_excel_demo_generator~generate_excel( ).

*** Create output
  lcl_output=>output( lo_excel ).



CLASS lcl_excel_generator IMPLEMENTATION.

  METHOD zif_excel_demo_generator~get_information.

    result-objid = sy-repid.
    result-text = 'abap2xlsx Demo: Read file and output'.
    result-filename = gc_save_file_name.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~checker_initialization.

    CLEAR: s_icon, s_icon[].
    s_icon-sign = 'I'.
    s_icon-option = 'CP'.
    s_icon-low = 'ICON_LED_*'.
    APPEND s_icon TO s_icon.

  ENDMETHOD.

  METHOD zif_excel_demo_generator~generate_excel.

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

ENDCLASS.
