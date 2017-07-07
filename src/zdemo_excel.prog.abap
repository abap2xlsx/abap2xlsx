*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel.

DATA: lv_workdir TYPE string,
      lv_upfile  TYPE string.


PARAMETERS: p_path TYPE zexcel_export_dir.

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

  IF p_path IS INITIAL.
    p_path = lv_workdir.
  ENDIF.

  cl_gui_frontend_services=>get_file_separator( CHANGING file_separator = sy-lisel ).
  CONCATENATE p_path sy-lisel '01_HelloWorld.xlsx' INTO lv_upfile.

  SUBMIT zdemo_excel1 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Hello world
  SUBMIT zdemo_excel2 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Styles
  SUBMIT zdemo_excel3 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: iTab binding
  SUBMIT zdemo_excel4 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Multi sheets, page setup and sheet properties
  SUBMIT zdemo_excel5 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Conditional formatting
  SUBMIT zdemo_excel6 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Formulas
  SUBMIT zdemo_excel7 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Conditional formatting
  SUBMIT zdemo_excel8 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Ranges
  SUBMIT zdemo_excel9 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path     = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Data validation
  SUBMIT zdemo_excel10 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Bind table with field catalog
  " zdemo_excel11 is not added because it has a selection screen and
  " you also need to have business partners maintained in transaction BP
  SUBMIT zdemo_excel12 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Column size
  SUBMIT zdemo_excel13 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Merge cell
  SUBMIT zdemo_excel14 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Alignment
  " zdemo_excel15 added at the end
  SUBMIT zdemo_excel16 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Drawing
  SUBMIT zdemo_excel17 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Lock sheet
  SUBMIT zdemo_excel18 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Lock workbook
  SUBMIT zdemo_excel19 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Set active sheet
  " zdemo_excel20 is not added because it uses ALV and cannot be processed (OLE2)
  SUBMIT zdemo_excel21 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Color Picker
  SUBMIT zdemo_excel22 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Bind table with field catalog & sheet style
  SUBMIT zdemo_excel23 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Multiple sheets with and w/o grid lines, print options
  SUBMIT zdemo_excel24 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Multiple sheets with different default date formats
  SUBMIT zdemo_excel25 AND RETURN.                                                                             "#EC CI_SUBMIT abap2xlsx Demo: Create and xlsx on Application Server (could be executed in batch mode)
  " zdemo_excel26 is not added because it uses ALV and cannot be processed (Native)
  SUBMIT zdemo_excel27 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Conditional Formatting
  SUBMIT zdemo_excel28 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: CSV writer
  " SUBMIT zdemo_excel29 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path  = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Macro enabled workbook
  SUBMIT zdemo_excel30 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: ABAP Cell data types + leading blanks string
  SUBMIT zdemo_excel31 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Autosize Column with different Font sizes
  " zdemo_excel32 is not added because it uses ALV and cannot be processed (Native)
  SUBMIT zdemo_excel33 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Table autofilter
  SUBMIT zdemo_excel34 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Static Styles Chess
  SUBMIT zdemo_excel35 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Static Styles
  SUBMIT zdemo_excel36 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Style applied to sheet, column and single cell
  SUBMIT zdemo_excel37 WITH p_upfile  = lv_upfile
                       WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Simplest call of the reader and writer - passthrough data
  SUBMIT zdemo_excel38 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Show off integration of drawings ( here using the SAP-Icons )
  SUBMIT zdemo_excel39 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Charts
  SUBMIT zdemo_excel40 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Demo Printsettings
  SUBMIT zdemo_excel41 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path    = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Inheritance
  "
  " Reader/Writer Demo must always run at the end
  " to make sure all documents where created
  "
  SUBMIT zdemo_excel15 WITH  p_path = p_path AND RETURN. "#EC CI_SUBMIT Read Excel and write it back
