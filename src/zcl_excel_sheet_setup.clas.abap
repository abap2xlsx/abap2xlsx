class ZCL_EXCEL_SHEET_SETUP definition
  public
  final
  create public .

public section.
*"* public components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
  type-pools ABAP .

  data BLACK_AND_WHITE type FLAG .
  data CELL_COMMENTS type STRINGVAL .
  data COPIES type INT2 .
  constants C_BREAK_COLUMN type ZEXCEL_BREAK value 2. "#EC NOTEXT
  constants C_BREAK_NONE type ZEXCEL_BREAK value 0. "#EC NOTEXT
  constants C_BREAK_ROW type ZEXCEL_BREAK value 1. "#EC NOTEXT
  constants C_CC_AS_DISPLAYED type STRING value 'asDisplayed'. "#EC NOTEXT
  constants C_CC_AT_END type STRING value 'atEnd'. "#EC NOTEXT
  constants C_CC_NONE type STRING value 'none'. "#EC NOTEXT
  constants C_ORD_DOWNTHENOVER type STRING value 'downThenOver'. "#EC NOTEXT
  constants C_ORD_OVERTHENDOWN type STRING value 'overThenDown'. "#EC NOTEXT
  constants C_ORIENTATION_DEFAULT type ZEXCEL_SHEET_ORIENATATION value 'default'. "#EC NOTEXT
  constants C_ORIENTATION_LANDSCAPE type ZEXCEL_SHEET_ORIENATATION value 'landscape'. "#EC NOTEXT
  constants C_ORIENTATION_PORTRAIT type ZEXCEL_SHEET_ORIENATATION value 'portrait'. "#EC NOTEXT
  constants C_PAPERSIZE_6_3_4_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 38. "#EC NOTEXT
  constants C_PAPERSIZE_A2_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 64. "#EC NOTEXT
  constants C_PAPERSIZE_A3 type ZEXCEL_SHEET_PAPER_SIZE value 8. "#EC NOTEXT
  constants C_PAPERSIZE_A3_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 61. "#EC NOTEXT
  constants C_PAPERSIZE_A3_EXTRA_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 66. "#EC NOTEXT
  constants C_PAPERSIZE_A3_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 65. "#EC NOTEXT
  constants C_PAPERSIZE_A4 type ZEXCEL_SHEET_PAPER_SIZE value 9. "#EC NOTEXT
  constants C_PAPERSIZE_A4_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 51. "#EC NOTEXT
  constants C_PAPERSIZE_A4_PLUS_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 58. "#EC NOTEXT
  constants C_PAPERSIZE_A4_SMALL type ZEXCEL_SHEET_PAPER_SIZE value 10. "#EC NOTEXT
  constants C_PAPERSIZE_A4_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 53. "#EC NOTEXT
  constants C_PAPERSIZE_A5 type ZEXCEL_SHEET_PAPER_SIZE value 11. "#EC NOTEXT
  constants C_PAPERSIZE_A5_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 62. "#EC NOTEXT
  constants C_PAPERSIZE_A5_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 59. "#EC NOTEXT
  constants C_PAPERSIZE_B4 type ZEXCEL_SHEET_PAPER_SIZE value 12. "#EC NOTEXT
  constants C_PAPERSIZE_B4_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 33. "#EC NOTEXT
  constants C_PAPERSIZE_B5 type ZEXCEL_SHEET_PAPER_SIZE value 13. "#EC NOTEXT
  constants C_PAPERSIZE_B5_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 34. "#EC NOTEXT
  constants C_PAPERSIZE_B6_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 35. "#EC NOTEXT
  constants C_PAPERSIZE_C type ZEXCEL_SHEET_PAPER_SIZE value 24. "#EC NOTEXT
  constants C_PAPERSIZE_C3_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 29. "#EC NOTEXT
  constants C_PAPERSIZE_C4_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 30. "#EC NOTEXT
  constants C_PAPERSIZE_C5_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 28. "#EC NOTEXT
  constants C_PAPERSIZE_C65_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 32. "#EC NOTEXT
  constants C_PAPERSIZE_C6_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 31. "#EC NOTEXT
  constants C_PAPERSIZE_D type ZEXCEL_SHEET_PAPER_SIZE value 25. "#EC NOTEXT
  constants C_PAPERSIZE_DE_LEG_FANFOLD type ZEXCEL_SHEET_PAPER_SIZE value 41. "#EC NOTEXT
  constants C_PAPERSIZE_DE_STD_FANFOLD type ZEXCEL_SHEET_PAPER_SIZE value 40. "#EC NOTEXT
  constants C_PAPERSIZE_DL_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 27. "#EC NOTEXT
  constants C_PAPERSIZE_E type ZEXCEL_SHEET_PAPER_SIZE value 26. "#EC NOTEXT
  constants C_PAPERSIZE_EXECUTIVE type ZEXCEL_SHEET_PAPER_SIZE value 7. "#EC NOTEXT
  constants C_PAPERSIZE_FOLIO type ZEXCEL_SHEET_PAPER_SIZE value 14. "#EC NOTEXT
  constants C_PAPERSIZE_INVITE_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 47. "#EC NOTEXT
  constants C_PAPERSIZE_ISO_B4 type ZEXCEL_SHEET_PAPER_SIZE value 42. "#EC NOTEXT
  constants C_PAPERSIZE_ISO_B5_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 63. "#EC NOTEXT
  constants C_PAPERSIZE_ITALY_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 36. "#EC NOTEXT
  constants C_PAPERSIZE_JIS_B5_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 60. "#EC NOTEXT
  constants C_PAPERSIZE_JPN_DBL_POSTCARD type ZEXCEL_SHEET_PAPER_SIZE value 43. "#EC NOTEXT
  constants C_PAPERSIZE_LEDGER type ZEXCEL_SHEET_PAPER_SIZE value 4. "#EC NOTEXT
  constants C_PAPERSIZE_LEGAL type ZEXCEL_SHEET_PAPER_SIZE value 5. "#EC NOTEXT
  constants C_PAPERSIZE_LEGAL_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 49. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER type ZEXCEL_SHEET_PAPER_SIZE value 1. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 48. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER_EXTV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 54. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER_PLUS_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 57. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER_SMALL type ZEXCEL_SHEET_PAPER_SIZE value 2. "#EC NOTEXT
  constants C_PAPERSIZE_LETTER_TV_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 52. "#EC NOTEXT
  constants C_PAPERSIZE_MONARCH_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 37. "#EC NOTEXT
  constants C_PAPERSIZE_NO10_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 20. "#EC NOTEXT
  constants C_PAPERSIZE_NO11_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 21. "#EC NOTEXT
  constants C_PAPERSIZE_NO12_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 22. "#EC NOTEXT
  constants C_PAPERSIZE_NO14_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 23. "#EC NOTEXT
  constants C_PAPERSIZE_NO9_ENVELOPE type ZEXCEL_SHEET_PAPER_SIZE value 19. "#EC NOTEXT
  constants C_PAPERSIZE_NOTE type ZEXCEL_SHEET_PAPER_SIZE value 18. "#EC NOTEXT
  constants C_PAPERSIZE_QUARTO type ZEXCEL_SHEET_PAPER_SIZE value 15. "#EC NOTEXT
  constants C_PAPERSIZE_STANDARD_1 type ZEXCEL_SHEET_PAPER_SIZE value 16. "#EC NOTEXT
  constants C_PAPERSIZE_STANDARD_2 type ZEXCEL_SHEET_PAPER_SIZE value 17. "#EC NOTEXT
  constants C_PAPERSIZE_STANDARD_PAPER_1 type ZEXCEL_SHEET_PAPER_SIZE value 44. "#EC NOTEXT
  constants C_PAPERSIZE_STANDARD_PAPER_2 type ZEXCEL_SHEET_PAPER_SIZE value 45. "#EC NOTEXT
  constants C_PAPERSIZE_STANDARD_PAPER_3 type ZEXCEL_SHEET_PAPER_SIZE value 46. "#EC NOTEXT
  constants C_PAPERSIZE_STATEMENT type ZEXCEL_SHEET_PAPER_SIZE value 6. "#EC NOTEXT
  constants C_PAPERSIZE_SUPERA_A4_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 55. "#EC NOTEXT
  constants C_PAPERSIZE_SUPERB_A3_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 56. "#EC NOTEXT
  constants C_PAPERSIZE_TABLOID type ZEXCEL_SHEET_PAPER_SIZE value 3. "#EC NOTEXT
  constants C_PAPERSIZE_TABL_EXTRA_PAPER type ZEXCEL_SHEET_PAPER_SIZE value 50. "#EC NOTEXT
  constants C_PAPERSIZE_US_STD_FANFOLD type ZEXCEL_SHEET_PAPER_SIZE value 39. "#EC NOTEXT
  constants C_PE_BLANK type STRING value 'blank'. "#EC NOTEXT
  constants C_PE_DASH type STRING value 'dash'. "#EC NOTEXT
  constants C_PE_DISPLAYED type STRING value 'displayed'. "#EC NOTEXT
  constants C_PE_NA type STRING value 'NA'. "#EC NOTEXT
  data DIFF_ODDEVEN_HEADERFOOTER type FLAG .
  data DRAFT type FLAG .
  data ERRORS type STRINGVAL .
  data EVEN_FOOTER type ZEXCEL_S_WORKSHEET_HEAD_FOOT .
  data EVEN_HEADER type ZEXCEL_S_WORKSHEET_HEAD_FOOT .
  data FIRST_PAGE_NUMBER type INT2 .
  data FIT_TO_HEIGHT type INT2 .
  data FIT_TO_PAGE type FLAG .
  data FIT_TO_WIDTH type INT2 .
  data HORIZONTAL_CENTERED type FLAG .
  data HORIZONTAL_DPI type INT2 .
  data MARGIN_BOTTOM type ZEXCEL_DEC_8_2 .
  data MARGIN_FOOTER type ZEXCEL_DEC_8_2 .
  data MARGIN_HEADER type ZEXCEL_DEC_8_2 .
  data MARGIN_LEFT type ZEXCEL_DEC_8_2 .
  data MARGIN_RIGHT type ZEXCEL_DEC_8_2 .
  data MARGIN_TOP type ZEXCEL_DEC_8_2 .
  data ODD_FOOTER type ZEXCEL_S_WORKSHEET_HEAD_FOOT .
  data ODD_HEADER type ZEXCEL_S_WORKSHEET_HEAD_FOOT .
  data ORIENTATION type ZEXCEL_SHEET_ORIENATATION .
  data PAGE_ORDER type STRING .
  data PAPER_HEIGHT type STRING .
  data PAPER_SIZE type INT2 .
  data PAPER_WIDTH type STRING .
  data SCALE type INT2 .
  data USE_FIRST_PAGE_NUM type FLAG .
  data USE_PRINTER_DEFAULTS type FLAG .
  data VERTICAL_CENTERED type FLAG .
  data VERTICAL_DPI type INT2 .

  methods CONSTRUCTOR .
  methods SET_PAGE_MARGINS
    importing
      !IP_BOTTOM type FLOAT optional
      !IP_FOOTER type FLOAT optional
      !IP_HEADER type FLOAT optional
      !IP_LEFT type FLOAT optional
      !IP_RIGHT type FLOAT optional
      !IP_TOP type FLOAT optional
      !IP_UNIT type CSEQUENCE default 'in' .
  methods SET_HEADER_FOOTER
    importing
      !IP_ODD_HEADER type ZEXCEL_S_WORKSHEET_HEAD_FOOT optional
      !IP_ODD_FOOTER type ZEXCEL_S_WORKSHEET_HEAD_FOOT optional
      !IP_EVEN_HEADER type ZEXCEL_S_WORKSHEET_HEAD_FOOT optional
      !IP_EVEN_FOOTER type ZEXCEL_S_WORKSHEET_HEAD_FOOT optional .
  methods GET_HEADER_FOOTER_STRING
    exporting
      !EP_ODD_HEADER type STRING
      !EP_ODD_FOOTER type STRING
      !EP_EVEN_HEADER type STRING
      !EP_EVEN_FOOTER type STRING .
protected section.

*"* protected components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
  methods PROCESS_HEADER_FOOTER
    importing
      !IP_HEADER type ZEXCEL_S_WORKSHEET_HEAD_FOOT
      !IP_SIDE type STRING
    returning
      value(RV_PROCESSED_STRING) type STRING .
private section.
*"* private components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_SHEET_SETUP IMPLEMENTATION.


method CONSTRUCTOR.
  orientation = me->c_orientation_default.

* default margins
  margin_bottom = '0.75'.
  margin_footer = '0.3'.
  margin_header = '0.3'.
  margin_left   = '0.7'.
  margin_right  = '0.7'.
  margin_top    = '0.75'.

* clear page settings
  CLEAR: black_and_white,
         cell_comments,
         copies,
         draft,
         errors,
         first_page_number,
         fit_to_page,
         fit_to_height,
         fit_to_width,
         horizontal_dpi,
         orientation,
         page_order,
         paper_height,
         paper_size,
         paper_width,
         scale,
         use_first_page_num,
         use_printer_defaults,
         vertical_dpi.
  endmethod.


method GET_HEADER_FOOTER_STRING.
* ----------------------------------------------------------------------
  DATA:   lc_marker_left(2)   TYPE c VALUE '&L'
        , lc_marker_right(2)  TYPE c VALUE '&R'
        , lc_marker_center(2) TYPE c VALUE '&C'
        , lv_value            TYPE string
        .
* ----------------------------------------------------------------------
  IF ep_odd_header IS SUPPLIED.

    IF me->odd_header-left_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'LEFT' ).
      CONCATENATE lc_marker_left lv_value INTO ep_odd_header.
    ENDIF.

    IF me->odd_header-center_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'CENTER' ).
      CONCATENATE ep_odd_header lc_marker_center lv_value INTO ep_odd_header.
    ENDIF.

    IF me->odd_header-right_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_header ip_side = 'RIGHT' ).
      CONCATENATE ep_odd_header lc_marker_right lv_value INTO ep_odd_header.
    ENDIF.

  ENDIF.
* ----------------------------------------------------------------------
  IF ep_odd_footer IS SUPPLIED.

    IF me->odd_footer-left_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'LEFT' ).
      CONCATENATE lc_marker_left lv_value INTO ep_odd_footer.
    ENDIF.

    IF me->odd_footer-center_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'CENTER' ).
      CONCATENATE ep_odd_footer lc_marker_center lv_value INTO ep_odd_footer.
    ENDIF.

    IF me->odd_footer-right_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->odd_footer ip_side = 'RIGHT' ).
      CONCATENATE ep_odd_footer lc_marker_right lv_value INTO ep_odd_footer.
    ENDIF.

  ENDIF.
* ----------------------------------------------------------------------
  IF ep_even_header IS SUPPLIED.

    IF me->even_header-left_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'LEFT' ).
      CONCATENATE lc_marker_left lv_value INTO ep_even_header.
    ENDIF.

    IF me->even_header-center_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'CENTER' ).
      CONCATENATE ep_even_header lc_marker_center lv_value INTO ep_even_header.
    ENDIF.

    IF me->even_header-right_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_header ip_side = 'RIGHT' ).
      CONCATENATE ep_even_header lc_marker_right lv_value INTO ep_even_header.
    ENDIF.

  ENDIF.
* ----------------------------------------------------------------------
  IF ep_even_footer IS SUPPLIED.

    IF me->even_footer-left_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'LEFT' ).
      CONCATENATE lc_marker_left lv_value INTO ep_even_footer.
    ENDIF.

    IF me->even_footer-center_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'CENTER' ).
      CONCATENATE ep_even_footer lc_marker_center lv_value INTO ep_even_footer.
    ENDIF.

    IF me->even_footer-right_value IS NOT INITIAL.
      lv_value = me->process_header_footer( ip_header = me->even_footer ip_side = 'RIGHT' ).
      CONCATENATE ep_even_footer lc_marker_right lv_value INTO ep_even_footer.
    ENDIF.

  ENDIF.
* ----------------------------------------------------------------------
  endmethod.


method PROCESS_HEADER_FOOTER.

* ----------------------------------------------------------------------
* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size

  DATA:   lv_fname(12) TYPE c
        , lv_string    TYPE string
        .

  FIELD-SYMBOLS:   <lv_value> TYPE string
                 , <ls_font>  TYPE zexcel_s_style_font
                 .

* ----------------------------------------------------------------------
  CONCATENATE ip_side '_VALUE' INTO lv_fname.
  ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <lv_value>.

  CONCATENATE ip_side '_FONT' INTO lv_fname.
  ASSIGN COMPONENT lv_fname OF STRUCTURE ip_header TO <ls_font>.

  IF <ls_font> IS ASSIGNED AND <lv_value> IS ASSIGNED.

    IF <ls_font>-name IS NOT INITIAL.
      CONCATENATE '&"' <ls_font>-name ',' INTO rv_processed_string.
    ELSE.
      rv_processed_string = '&"-,'.
    ENDIF.

    IF <ls_font>-bold = abap_true.
      CONCATENATE rv_processed_string 'Bold"' INTO rv_processed_string.
    ELSE.
      CONCATENATE rv_processed_string 'Standard"' INTO rv_processed_string.
    ENDIF.

    IF <ls_font>-size IS NOT INITIAL.
      lv_string = <ls_font>-size.
      CONCATENATE rv_processed_string '&' lv_string INTO rv_processed_string.
    ENDIF.

    CONCATENATE rv_processed_string <lv_value> INTO rv_processed_string.

  ENDIF.
* ----------------------------------------------------------------------

  endmethod.


method SET_HEADER_FOOTER.

* Only Basic font/text formatting possible:
* Bold (yes / no), Font Type, Font Size
*
* usefull placeholders, which can be used in header/footer value strings
* '&P' - page number
* '&N' - total number of pages
* '&D' - Date
* '&T' - Time
* '&F' - File Name
* '&Z' - Path
* '&A' - Sheet name
* new line via class constant CL_ABAP_CHAR_UTILITIES=>newline
*
* Example Value String 'page &P of &N'
*
* DO NOT USE &L , &C or &R which automatically created as position markers

  me->odd_header = ip_odd_header.
  me->odd_footer = ip_odd_footer.
  me->even_header = ip_even_header.
  me->even_footer = ip_even_footer.

  IF me->even_header IS NOT INITIAL OR me->even_footer IS NOT INITIAL.
    me->diff_oddeven_headerfooter = abap_true.
  ENDIF.


  endmethod.


method SET_PAGE_MARGINS.
  DATA: lv_coef TYPE f,
        lv_unit TYPE string.

  lv_unit = ip_unit.
  TRANSLATE lv_unit TO UPPER CASE.

  CASE lv_unit.
    WHEN 'IN'. lv_coef = 1.
    WHEN 'CM'. lv_coef = '0.393700787'.
    WHEN 'MM'. lv_coef = '0.0393700787'.
  ENDCASE.

  IF ip_bottom IS SUPPLIED. margin_bottom = lv_coef * ip_bottom. ENDIF.
  IF ip_footer IS SUPPLIED. margin_footer = lv_coef * ip_footer. ENDIF.
  IF ip_header IS SUPPLIED. margin_header = lv_coef * ip_header. ENDIF.
  IF ip_left IS SUPPLIED.   margin_left   = lv_coef * ip_left. ENDIF.
  IF ip_right IS SUPPLIED.  margin_right  = lv_coef * ip_right. ENDIF.
  IF ip_top IS SUPPLIED.    margin_top    = lv_coef * ip_top. ENDIF.

  endmethod.
ENDCLASS.
