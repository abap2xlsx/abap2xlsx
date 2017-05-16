*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL10
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel10.

DATA: lo_excel                TYPE REF TO zcl_excel,
      lo_worksheet            TYPE REF TO zcl_excel_worksheet,
      lo_style_conditional2   TYPE REF TO zcl_excel_style_conditional,
      column_dimension        TYPE REF TO zcl_excel_worksheet_columndime.

DATA: lt_field_catalog        TYPE zexcel_t_fieldcatalog,
      ls_table_settings       TYPE zexcel_s_table_settings,
      ls_iconset              TYPE zexcel_conditional_iconset.

CONSTANTS: gc_save_file_name TYPE string VALUE '10_iTabFieldCatalog.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


START-OF-SELECTION.

  FIELD-SYMBOLS: <fs_field_catalog> TYPE zexcel_s_fieldcatalog.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).

  ls_iconset-iconset                  = zcl_excel_style_conditional=>c_iconset_5arrows.
  ls_iconset-cfvo1_type               = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo1_value              = '0'.
  ls_iconset-cfvo2_type               = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo2_value              = '20'.
  ls_iconset-cfvo3_type               = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo3_value              = '40'.
  ls_iconset-cfvo4_type               = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo4_value              = '60'.
  ls_iconset-cfvo5_type               = zcl_excel_style_conditional=>c_cfvo_type_percent.
  ls_iconset-cfvo5_value              = '80'.
  ls_iconset-showvalue                = zcl_excel_style_conditional=>c_showvalue_true.

  "Conditional style
  lo_style_conditional2 = lo_worksheet->add_new_conditional_style( ).
  lo_style_conditional2->rule         = zcl_excel_style_conditional=>c_rule_iconset.
  lo_style_conditional2->mode_iconset = ls_iconset.
  lo_style_conditional2->priority     = 1.

  DATA lt_test TYPE TABLE OF sflight.
  SELECT * FROM sflight INTO TABLE lt_test. "#EC CI_NOWHERE

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_test ).

  LOOP AT lt_field_catalog ASSIGNING <fs_field_catalog>.
    CASE <fs_field_catalog>-fieldname.
      WHEN 'CARRID'.
        <fs_field_catalog>-position   = 3.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-totals_function = zcl_excel_table=>totals_function_count.
      WHEN 'CONNID'.
        <fs_field_catalog>-position   = 4.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-abap_type  = cl_abap_typedescr=>typekind_int.
        "This avoid the excel warning that the number is formatted as a text: abap2xlsx is not able to recognize numc as a number so it formats the number as a text with
        "the related warning. You can force the type and the framework will correctly format the number as a number
      WHEN 'FLDATE'.
        <fs_field_catalog>-position   = 2.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'PRICE'.
        <fs_field_catalog>-position   = 1.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-totals_function = zcl_excel_table=>totals_function_sum.
        <fs_field_catalog>-cond_style = lo_style_conditional2.
      WHEN OTHERS.
        <fs_field_catalog>-dynpfld = abap_false.
    ENDCASE.
  ENDLOOP.

  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium5.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings
                            it_field_catalog  = lt_field_catalog ).

  column_dimension = lo_worksheet->get_column_dimension( ip_column = 'D' ). "make date field a bit wider
  column_dimension->set_width( ip_width = 13 ).


*** Create output
  lcl_output=>output( lo_excel ).
