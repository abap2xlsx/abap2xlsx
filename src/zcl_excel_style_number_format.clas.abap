class ZCL_EXCEL_STYLE_NUMBER_FORMAT definition
  public
  final
  create public .

public section.

  types:
    begin of t_num_format,
      id     type string,
      format type ref to zcl_excel_style_number_format,
    end of t_num_format .
  types:
    t_num_formats type hashed table of t_num_format with unique key id .

*"* public components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
  constants C_FORMAT_NUMC_STD type ZEXCEL_NUMBER_FORMAT value 'STD_NDEC'. "#EC NOTEXT
  constants C_FORMAT_DATE_STD type ZEXCEL_NUMBER_FORMAT value 'STD_DATE'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_EUR_SIMPLE type ZEXCEL_NUMBER_FORMAT value '[$EUR ]#,##0.00_-'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_USD type ZEXCEL_NUMBER_FORMAT value '$#,##0_-'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_USD_SIMPLE type ZEXCEL_NUMBER_FORMAT value '"$"#,##0.00_-'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_SIMPLE type ZEXCEL_NUMBER_FORMAT value '$#,##0_);($#,##0)'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_SIMPLE_RED type ZEXCEL_NUMBER_FORMAT value '$#,##0_);[Red]($#,##0)'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_SIMPLE2 type ZEXCEL_NUMBER_FORMAT value '$#,##0.00_);($#,##0.00)'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_SIMPLE_RED2 type ZEXCEL_NUMBER_FORMAT value '$#,##0.00_);[Red]($#,##0.00)'. "#EC NOTEXT
  constants C_FORMAT_DATE_DATETIME type ZEXCEL_NUMBER_FORMAT value 'd/m/y h:mm'. "#EC NOTEXT
  constants C_FORMAT_DATE_DDMMYYYY type ZEXCEL_NUMBER_FORMAT value 'dd/mm/yy'. "#EC NOTEXT
  constants C_FORMAT_DATE_DDMMYYYYDOT type ZEXCEL_NUMBER_FORMAT value 'dd\.mm\.yyyy'. "#EC NOTEXT
  constants C_FORMAT_DATE_DMMINUS type ZEXCEL_NUMBER_FORMAT value 'd-m'. "#EC NOTEXT
  constants C_FORMAT_DATE_DMYMINUS type ZEXCEL_NUMBER_FORMAT value 'd-m-y'. "#EC NOTEXT
  constants C_FORMAT_DATE_DMYSLASH type ZEXCEL_NUMBER_FORMAT value 'd/m/y'. "#EC NOTEXT
  constants C_FORMAT_DATE_MYMINUS type ZEXCEL_NUMBER_FORMAT value 'm-y'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME1 type ZEXCEL_NUMBER_FORMAT value 'h:mm AM/PM'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME2 type ZEXCEL_NUMBER_FORMAT value 'h:mm:ss AM/PM'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME3 type ZEXCEL_NUMBER_FORMAT value 'h:mm'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME4 type ZEXCEL_NUMBER_FORMAT value 'h:mm:ss'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME5 type ZEXCEL_NUMBER_FORMAT value 'mm:ss'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME6 type ZEXCEL_NUMBER_FORMAT value 'h:mm:ss'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME7 type ZEXCEL_NUMBER_FORMAT value 'i:s.S'. "#EC NOTEXT
  constants C_FORMAT_DATE_TIME8 type ZEXCEL_NUMBER_FORMAT value 'h:mm:ss@'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX14 type ZEXCEL_NUMBER_FORMAT value 'mm-dd-yy'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX15 type ZEXCEL_NUMBER_FORMAT value 'd-mmm-yy'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX16 type ZEXCEL_NUMBER_FORMAT value 'd-mmm'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX17 type ZEXCEL_NUMBER_FORMAT value 'mmm-yy'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX22 type ZEXCEL_NUMBER_FORMAT value 'm/d/yy h:mm'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYMMDD type ZEXCEL_NUMBER_FORMAT value 'yymmdd'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYMMDDMINUS type ZEXCEL_NUMBER_FORMAT value 'yy-mm-dd'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYMMDDSLASH type ZEXCEL_NUMBER_FORMAT value 'yy/mm/dd'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYYYMMDD type ZEXCEL_NUMBER_FORMAT value 'yyyymmdd'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYYYMMDDMINUS type ZEXCEL_NUMBER_FORMAT value 'yyyy-mm-dd'. "#EC NOTEXT
  constants C_FORMAT_DATE_YYYYMMDDSLASH type ZEXCEL_NUMBER_FORMAT value 'yyyy/mm/dd'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX45 type ZEXCEL_NUMBER_FORMAT value 'mm:ss'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX46 type ZEXCEL_NUMBER_FORMAT value '[h]:mm:ss'. "#EC NOTEXT
  constants C_FORMAT_DATE_XLSX47 type ZEXCEL_NUMBER_FORMAT value 'mm:ss.0'. "#EC NOTEXT
  constants C_FORMAT_GENERAL type ZEXCEL_NUMBER_FORMAT value ''. "#EC NOTEXT
  constants C_FORMAT_NUMBER type ZEXCEL_NUMBER_FORMAT value '0'. "#EC NOTEXT
  constants C_FORMAT_NUMBER_00 type ZEXCEL_NUMBER_FORMAT value '0.00'. "#EC NOTEXT
  constants C_FORMAT_NUMBER_COMMA_SEP0 type ZEXCEL_NUMBER_FORMAT value '#,##0'. "#EC NOTEXT
  constants C_FORMAT_NUMBER_COMMA_SEP1 type ZEXCEL_NUMBER_FORMAT value '#,##0.00'. "#EC NOTEXT
  constants C_FORMAT_NUMBER_COMMA_SEP2 type ZEXCEL_NUMBER_FORMAT value '#,##0.00_-'. "#EC NOTEXT
  constants C_FORMAT_PERCENTAGE type ZEXCEL_NUMBER_FORMAT value '0%'. "#EC NOTEXT
  constants C_FORMAT_PERCENTAGE_00 type ZEXCEL_NUMBER_FORMAT value '0.00%'. "#EC NOTEXT
  constants C_FORMAT_TEXT type ZEXCEL_NUMBER_FORMAT value '@'. "#EC NOTEXT
  constants C_FORMAT_FRACTION_1 type ZEXCEL_NUMBER_FORMAT value '# ?/?'. "#EC NOTEXT
  constants C_FORMAT_FRACTION_2 type ZEXCEL_NUMBER_FORMAT value '# ??/??'. "#EC NOTEXT
  constants C_FORMAT_SCIENTIFIC type ZEXCEL_NUMBER_FORMAT value '0.00E+00'. "#EC NOTEXT
  constants C_FORMAT_SPECIAL_01 type ZEXCEL_NUMBER_FORMAT value '##0.0E+0'. "#EC NOTEXT
  data FORMAT_CODE type ZEXCEL_NUMBER_FORMAT .
  class-data MT_BUILT_IN_NUM_FORMATS type T_NUM_FORMATS read-only .
  constants C_FORMAT_XLSX37 type ZEXCEL_NUMBER_FORMAT value '#,##0_);(#,##0)'. "#EC NOTEXT
  constants C_FORMAT_XLSX38 type ZEXCEL_NUMBER_FORMAT value '#,##0_);[Red](#,##0)'. "#EC NOTEXT
  constants C_FORMAT_XLSX39 type ZEXCEL_NUMBER_FORMAT value '#,##0.00_);(#,##0.00)'. "#EC NOTEXT
  constants C_FORMAT_XLSX40 type ZEXCEL_NUMBER_FORMAT value '#,##0.00_);[Red](#,##0.00)'. "#EC NOTEXT
  constants C_FORMAT_XLSX41 type ZEXCEL_NUMBER_FORMAT value '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'. "#EC NOTEXT
  constants C_FORMAT_XLSX42 type ZEXCEL_NUMBER_FORMAT value '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'. "#EC NOTEXT
  constants C_FORMAT_XLSX43 type ZEXCEL_NUMBER_FORMAT value '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'. "#EC NOTEXT
  constants C_FORMAT_XLSX44 type ZEXCEL_NUMBER_FORMAT value '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_GBP_SIMPLE type ZEXCEL_NUMBER_FORMAT value '[$£-809]#,##0.00'. "#EC NOTEXT
  constants C_FORMAT_CURRENCY_PLN_SIMPLE type ZEXCEL_NUMBER_FORMAT value '#,##0.00\ "zł"'. "#EC NOTEXT

  class-methods CLASS_CONSTRUCTOR .
  methods CONSTRUCTOR .
  methods GET_STRUCTURE
    returning
      value(EP_NUMBER_FORMAT) type ZEXCEL_S_STYLE_NUMFMT .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
protected section.
private section.
*"* private components of class ZCL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_STYLE_NUMBER_FORMAT IMPLEMENTATION.


METHOD class_constructor.

  DATA: ls_num_format LIKE LINE OF mt_built_in_num_formats.

  DEFINE predefined_format.
    ls_num_format-id                  = &1.
    create object ls_num_format-format.
    ls_num_format-format->format_code = &2.
    insert ls_num_format into table mt_built_in_num_formats.
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


method CONSTRUCTOR.
  format_code = me->c_format_general.
  endmethod.


method GET_STRUCTURE.
  ep_number_format-numfmt = me->format_code.
  endmethod.
ENDCLASS.
