CLASS ltc_converter_alv DEFINITION FOR TESTING INHERITING FROM zcl_excel_converter_alv.

  PUBLIC SECTION.

    TYPES: BEGIN OF ty_alv_line,
             field_1 TYPE string,
             field_2 TYPE string,
           END OF ty_alv_line,
           ty_alv_table TYPE STANDARD TABLE OF ty_alv_line WITH DEFAULT KEY.

    METHODS:
      zif_excel_converter~can_convert_object REDEFINITION,
      zif_excel_converter~create_fieldcatalog REDEFINITION,
      zif_excel_converter~get_supported_class REDEFINITION.

    METHODS set_alv_table
      IMPORTING
        alv_table TYPE ty_alv_table.

    METHODS set_alv_filter_table
      IMPORTING
        alv_filter_table TYPE lvc_t_filt.

    METHODS wrapper_of_get_filter
      EXPORTING
        et_filter             TYPE zexcel_t_converter_fil
        et_filtered_alv_table TYPE ty_alv_table.

  PRIVATE SECTION.

    DATA: alv_table TYPE ty_alv_table.

ENDCLASS.


CLASS ltc_converter_alv IMPLEMENTATION.

  METHOD zif_excel_converter~can_convert_object.

  ENDMETHOD.

  METHOD zif_excel_converter~create_fieldcatalog.

  ENDMETHOD.

  METHOD zif_excel_converter~get_supported_class.

  ENDMETHOD.

  METHOD wrapper_of_get_filter.
    DATA: ref_alv_table TYPE REF TO data.

    et_filtered_alv_table = alv_table.
    GET REFERENCE OF et_filtered_alv_table INTO ref_alv_table.

    " Get ET_FILTER
    ws_option-filter = abap_true.
    get_filter(
      IMPORTING
        et_filter = et_filter
      CHANGING
        xo_table  = ref_alv_table ).

    " Get filtered table
    ws_option-filter = abap_undefined.
    get_filter(
      CHANGING
        xo_table  = ref_alv_table ).

  ENDMETHOD.

  METHOD set_alv_table.

    me->alv_table = alv_table.

  ENDMETHOD.

  METHOD set_alv_filter_table.

    wt_filt = alv_filter_table.

  ENDMETHOD.

ENDCLASS.


CLASS ltc_get_filter DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.

    TYPES: BEGIN OF ty_output,
             filter_table       TYPE zexcel_t_converter_fil,
             filtered_alv_table TYPE ltc_converter_alv=>ty_alv_table,
           END OF ty_output.
    DATA:
      f_cut              TYPE REF TO ltc_converter_alv,  "class under test
      alv_line           TYPE ltc_converter_alv=>ty_alv_line,
      alv_table          TYPE ltc_converter_alv=>ty_alv_table,
      alv_filter_line    TYPE lvc_s_filt,
      alv_filter_table   TYPE lvc_t_filt,
      filter_line        TYPE zexcel_s_converter_fil,
      filter_table       TYPE zexcel_t_converter_fil,
      filtered_alv_line  TYPE ltc_converter_alv=>ty_alv_line,
      filtered_alv_table TYPE ltc_converter_alv=>ty_alv_table,
      act                TYPE ty_output,
      exp                TYPE ty_output.

    METHODS:
      none FOR TESTING RAISING cx_static_check,
      one_field_eq_or_eq FOR TESTING RAISING cx_static_check,
      simple FOR TESTING RAISING cx_static_check,
      two_fields_eq FOR TESTING RAISING cx_static_check.

    METHODS: setup.

ENDCLASS.


CLASS ltc_get_filter IMPLEMENTATION.


  METHOD none.

    f_cut->wrapper_of_get_filter(
      IMPORTING
        et_filter             = act-filter_table
        et_filtered_alv_table = act-filtered_alv_table ).

    exp-filtered_alv_table = alv_table.

    cl_abap_unit_assert=>assert_equals( act = act-filter_table exp = exp-filter_table ).
    cl_abap_unit_assert=>assert_equals( act = act-filtered_alv_table exp = exp-filtered_alv_table ).

  ENDMETHOD.


  METHOD one_field_eq_or_eq.

    alv_filter_line-fieldname = 'FIELD_1'.
    alv_filter_line-sign = 'I'.
    alv_filter_line-option = 'EQ'.
    alv_filter_line-low = 'A'.
    APPEND alv_filter_line TO alv_filter_table.
    alv_filter_line-low = 'B'.
    APPEND alv_filter_line TO alv_filter_table.
    f_cut->set_alv_filter_table( alv_filter_table ).

    f_cut->wrapper_of_get_filter(
      IMPORTING
        et_filter             = act-filter_table
        et_filtered_alv_table = act-filtered_alv_table ).

    filter_line-rownumber = 1.
    filter_line-columnname = 'FIELD_1'.
    INSERT filter_line INTO TABLE exp-filter_table.
    filter_line-rownumber = 2.
    filter_line-columnname = 'FIELD_1'.
    INSERT filter_line INTO TABLE exp-filter_table.

    exp-filtered_alv_table = alv_table.

    cl_abap_unit_assert=>assert_equals( act = act-filter_table exp = exp-filter_table ).
    cl_abap_unit_assert=>assert_equals( act = act-filtered_alv_table exp = exp-filtered_alv_table ).

  ENDMETHOD.


  METHOD setup.

    CREATE OBJECT f_cut.
    alv_line-field_1 = 'A'.
    alv_line-field_2 = 'C'.
    APPEND alv_line TO alv_table.
    alv_line-field_1 = 'B'.
    alv_line-field_2 = 'D'.
    APPEND alv_line TO alv_table.
    f_cut->set_alv_table( alv_table ).

  ENDMETHOD.


  METHOD simple.

    alv_filter_line-fieldname = 'FIELD_1'.
    alv_filter_line-sign = 'I'.
    alv_filter_line-option = 'EQ'.
    alv_filter_line-low = 'A'.
    alv_filter_line-high = ''.
    APPEND alv_filter_line TO alv_filter_table.
    f_cut->set_alv_filter_table( alv_filter_table ).

    f_cut->wrapper_of_get_filter(
      IMPORTING
        et_filter             = act-filter_table
        et_filtered_alv_table = act-filtered_alv_table ).

    filter_line-rownumber = 1.
    filter_line-columnname = 'FIELD_1'.
    INSERT filter_line INTO TABLE exp-filter_table.

    exp-filtered_alv_table = alv_table.
    DELETE exp-filtered_alv_table WHERE field_1 <> 'A'.

    cl_abap_unit_assert=>assert_equals( act = act-filter_table exp = exp-filter_table ).
    cl_abap_unit_assert=>assert_equals( act = act-filtered_alv_table exp = exp-filtered_alv_table ).

  ENDMETHOD.


  METHOD two_fields_eq.

    alv_filter_line-fieldname = 'FIELD_1'.
    alv_filter_line-sign = 'I'.
    alv_filter_line-option = 'EQ'.
    alv_filter_line-low = 'A'.
    APPEND alv_filter_line TO alv_filter_table.
    alv_filter_line-fieldname = 'FIELD_2'.
    alv_filter_line-low = 'D'.
    APPEND alv_filter_line TO alv_filter_table.
    f_cut->set_alv_filter_table( alv_filter_table ).

    f_cut->wrapper_of_get_filter(
      IMPORTING
        et_filter             = act-filter_table
        et_filtered_alv_table = act-filtered_alv_table ).

    cl_abap_unit_assert=>assert_equals( act = act-filter_table exp = exp-filter_table ).
    cl_abap_unit_assert=>assert_equals( act = act-filtered_alv_table exp = exp-filtered_alv_table ).

  ENDMETHOD.


ENDCLASS.
