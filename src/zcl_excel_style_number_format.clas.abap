CLASS zcl_excel_style_number_format DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF t_num_format,
        id     TYPE string,
        format TYPE REF TO zcl_excel_style_number_format,
      END OF t_num_format .
    TYPES:
      t_num_formats TYPE HASHED TABLE OF t_num_format WITH UNIQUE KEY id .

*"* public components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
    CONSTANTS c_format_numc_std TYPE zexcel_number_format VALUE 'STD_NDEC'. "#EC NOTEXT
    CONSTANTS c_format_date_std TYPE zexcel_number_format VALUE 'STD_DATE'. "#EC NOTEXT
    CONSTANTS c_format_currency_eur_simple TYPE zexcel_number_format VALUE '[$EUR ]#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_usd TYPE zexcel_number_format VALUE '$#,##0_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_usd_simple TYPE zexcel_number_format VALUE '"$"#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple TYPE zexcel_number_format VALUE '$#,##0_);($#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple_red TYPE zexcel_number_format VALUE '$#,##0_);[Red]($#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple2 TYPE zexcel_number_format VALUE '$#,##0.00_);($#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple_red2 TYPE zexcel_number_format VALUE '$#,##0.00_);[Red]($#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_date_datetime TYPE zexcel_number_format VALUE 'd/m/y h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_ddmmyyyy TYPE zexcel_number_format VALUE 'dd/mm/yy'. "#EC NOTEXT
    CONSTANTS c_format_date_ddmmyyyydot TYPE zexcel_number_format VALUE 'dd\.mm\.yyyy'. "#EC NOTEXT
    CONSTANTS c_format_date_dmminus TYPE zexcel_number_format VALUE 'd-m'. "#EC NOTEXT
    CONSTANTS c_format_date_dmyminus TYPE zexcel_number_format VALUE 'd-m-y'. "#EC NOTEXT
    CONSTANTS c_format_date_dmyslash TYPE zexcel_number_format VALUE 'd/m/y'. "#EC NOTEXT
    CONSTANTS c_format_date_myminus TYPE zexcel_number_format VALUE 'm-y'. "#EC NOTEXT
    CONSTANTS c_format_date_time1 TYPE zexcel_number_format VALUE 'h:mm AM/PM'. "#EC NOTEXT
    CONSTANTS c_format_date_time2 TYPE zexcel_number_format VALUE 'h:mm:ss AM/PM'. "#EC NOTEXT
    CONSTANTS c_format_date_time3 TYPE zexcel_number_format VALUE 'h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_time4 TYPE zexcel_number_format VALUE 'h:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time5 TYPE zexcel_number_format VALUE 'mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time6 TYPE zexcel_number_format VALUE 'h:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time7 TYPE zexcel_number_format VALUE 'i:s.S'. "#EC NOTEXT
    CONSTANTS c_format_date_time8 TYPE zexcel_number_format VALUE 'h:mm:ss@'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx14 TYPE zexcel_number_format VALUE 'mm-dd-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx15 TYPE zexcel_number_format VALUE 'd-mmm-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx16 TYPE zexcel_number_format VALUE 'd-mmm'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx17 TYPE zexcel_number_format VALUE 'mmm-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx22 TYPE zexcel_number_format VALUE 'm/d/yy h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmdd TYPE zexcel_number_format VALUE 'yymmdd'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmddminus TYPE zexcel_number_format VALUE 'yy-mm-dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmddslash TYPE zexcel_number_format VALUE 'yy/mm/dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmdd TYPE zexcel_number_format VALUE 'yyyymmdd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmddminus TYPE zexcel_number_format VALUE 'yyyy-mm-dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmddslash TYPE zexcel_number_format VALUE 'yyyy/mm/dd'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx45 TYPE zexcel_number_format VALUE 'mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx46 TYPE zexcel_number_format VALUE '[h]:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx47 TYPE zexcel_number_format VALUE 'mm:ss.0'. "#EC NOTEXT
    CONSTANTS c_format_general TYPE zexcel_number_format VALUE ''. "#EC NOTEXT
    CONSTANTS c_format_number TYPE zexcel_number_format VALUE '0'. "#EC NOTEXT
    CONSTANTS c_format_number_00 TYPE zexcel_number_format VALUE '0.00'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep0 TYPE zexcel_number_format VALUE '#,##0'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep1 TYPE zexcel_number_format VALUE '#,##0.00'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep2 TYPE zexcel_number_format VALUE '#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_percentage TYPE zexcel_number_format VALUE '0%'. "#EC NOTEXT
    CONSTANTS c_format_percentage_00 TYPE zexcel_number_format VALUE '0.00%'. "#EC NOTEXT
    CONSTANTS c_format_text TYPE zexcel_number_format VALUE '@'. "#EC NOTEXT
    CONSTANTS c_format_fraction_1 TYPE zexcel_number_format VALUE '# ?/?'. "#EC NOTEXT
    CONSTANTS c_format_fraction_2 TYPE zexcel_number_format VALUE '# ??/??'. "#EC NOTEXT
    CONSTANTS c_format_scientific TYPE zexcel_number_format VALUE '0.00E+00'. "#EC NOTEXT
    CONSTANTS c_format_special_01 TYPE zexcel_number_format VALUE '##0.0E+0'. "#EC NOTEXT
    DATA format_code TYPE zexcel_number_format .
    CLASS-DATA mt_built_in_num_formats TYPE t_num_formats READ-ONLY .
    CONSTANTS c_format_xlsx37 TYPE zexcel_number_format VALUE '#,##0_);(#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx38 TYPE zexcel_number_format VALUE '#,##0_);[Red](#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx39 TYPE zexcel_number_format VALUE '#,##0.00_);(#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx40 TYPE zexcel_number_format VALUE '#,##0.00_);[Red](#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx41 TYPE zexcel_number_format VALUE '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx42 TYPE zexcel_number_format VALUE '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx43 TYPE zexcel_number_format VALUE '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx44 TYPE zexcel_number_format VALUE '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_currency_gbp_simple TYPE zexcel_number_format VALUE '[$£-809]#,##0.00'. "#EC NOTEXT
    CONSTANTS c_format_currency_pln_simple TYPE zexcel_number_format VALUE '#,##0.00\ "zł"'. "#EC NOTEXT

    CLASS-METHODS class_constructor .
    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(ep_number_format) TYPE zexcel_s_style_numfmt .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_style_number_format IMPLEMENTATION.


  METHOD class_constructor.

    DATA: ls_num_format LIKE LINE OF mt_built_in_num_formats.

    DEFINE predefined_format.
      ls_num_format-id                  = &1.
      CREATE OBJECT ls_num_format-format.
      ls_num_format-format->format_code = &2.
      INSERT ls_num_format INTO TABLE mt_built_in_num_formats.
    END-OF-DEFINITION.

    CLEAR mt_built_in_num_formats.

    predefined_format    '1'        zcl_excel_style_number_format=>c_format_number.               " '0'.
    predefined_format    '2'        zcl_excel_style_number_format=>c_format_number_00.            " '0.00'.
    predefined_format    '3'        zcl_excel_style_number_format=>c_format_number_comma_sep0.    " '#,##0'.
    predefined_format    '4'        zcl_excel_style_number_format=>c_format_number_comma_sep1.    " '#,##0.00'.
    predefined_format    '5'        zcl_excel_style_number_format=>c_format_currency_simple.      " '$#,##0_);($#,##0)'.
    predefined_format    '6'        zcl_excel_style_number_format=>c_format_currency_simple_red.  " '$#,##0_);[Red]($#,##0)'.
    predefined_format    '7'        zcl_excel_style_number_format=>c_format_currency_simple2.     " '$#,##0.00_);($#,##0.00)'.
    predefined_format    '8'        zcl_excel_style_number_format=>c_format_currency_simple_red2. " '$#,##0.00_);[Red]($#,##0.00)'.
    predefined_format    '9'        zcl_excel_style_number_format=>c_format_percentage.           " '0%'.
    predefined_format   '10'        zcl_excel_style_number_format=>c_format_percentage_00.        " '0.00%'.
    predefined_format   '11'        zcl_excel_style_number_format=>c_format_scientific.           " '0.00E+00'.
    predefined_format   '12'        zcl_excel_style_number_format=>c_format_fraction_1.           " '# ?/?'.
    predefined_format   '13'        zcl_excel_style_number_format=>c_format_fraction_2.           " '# ??/??'.
    predefined_format   '14'        zcl_excel_style_number_format=>c_format_date_xlsx14.          "'m/d/yyyy'.  <--  should have been 'mm-dd-yy' like constant in zcl_excel_style_number_format
    predefined_format   '15'        zcl_excel_style_number_format=>c_format_date_xlsx15.          "'d-mmm-yy'.
    predefined_format   '16'        zcl_excel_style_number_format=>c_format_date_xlsx16.          "'d-mmm'.
    predefined_format   '17'        zcl_excel_style_number_format=>c_format_date_xlsx17.          "'mmm-yy'.
    predefined_format   '18'        zcl_excel_style_number_format=>c_format_date_time1.           " 'h:mm AM/PM'.
    predefined_format   '19'        zcl_excel_style_number_format=>c_format_date_time2.           " 'h:mm:ss AM/PM'.
    predefined_format   '20'        zcl_excel_style_number_format=>c_format_date_time3.           " 'h:mm'.
    predefined_format   '21'        zcl_excel_style_number_format=>c_format_date_time4.           " 'h:mm:ss'.
    predefined_format   '22'        zcl_excel_style_number_format=>c_format_date_xlsx22.          " 'm/d/yyyy h:mm'.


    predefined_format   '37'        zcl_excel_style_number_format=>c_format_xlsx37.               " '#,##0_);(#,##0)'.
    predefined_format   '38'        zcl_excel_style_number_format=>c_format_xlsx38.               " '#,##0_);[Red](#,##0)'.
    predefined_format   '39'        zcl_excel_style_number_format=>c_format_xlsx39.               " '#,##0.00_);(#,##0.00)'.
    predefined_format   '40'        zcl_excel_style_number_format=>c_format_xlsx40.               " '#,##0.00_);[Red](#,##0.00)'.
    predefined_format   '41'        zcl_excel_style_number_format=>c_format_xlsx41.               " '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'.
    predefined_format   '42'        zcl_excel_style_number_format=>c_format_xlsx42.               " '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'.
    predefined_format   '43'        zcl_excel_style_number_format=>c_format_xlsx43.               " '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'.
    predefined_format   '44'        zcl_excel_style_number_format=>c_format_xlsx44.               " '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'.
    predefined_format   '45'        zcl_excel_style_number_format=>c_format_date_xlsx45.          " 'mm:ss'.
    predefined_format   '46'        zcl_excel_style_number_format=>c_format_date_xlsx46.          " '[h]:mm:ss'.
    predefined_format   '47'        zcl_excel_style_number_format=>c_format_date_xlsx47.          "  'mm:ss.0'.
    predefined_format   '48'        zcl_excel_style_number_format=>c_format_special_01.           " '##0.0E+0'.
    predefined_format   '49'        zcl_excel_style_number_format=>c_format_text.                 " '@'.

  ENDMETHOD.


  METHOD constructor.
    format_code = me->c_format_general.
  ENDMETHOD.


  METHOD get_structure.
    ep_number_format-numfmt = me->format_code.
  ENDMETHOD.
ENDCLASS.
