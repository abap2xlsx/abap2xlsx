CLASS zcl_excel_style_changer DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES zif_excel_style_changer.

    CLASS-METHODS create
      IMPORTING
        excel         TYPE REF TO zcl_excel
      RETURNING
        VALUE(result) TYPE REF TO zif_excel_style_changer
      RAISING
        zcx_excel.

  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS clear_initial_colorxfields
      IMPORTING
        is_color  TYPE zexcel_s_style_color
      CHANGING
        cs_xcolor TYPE zexcel_s_cstylex_color.

    METHODS move_supplied_borders
      IMPORTING
        iv_border_supplied        TYPE abap_bool
        is_border                 TYPE zexcel_s_cstyle_border
        iv_xborder_supplied       TYPE abap_bool
        is_xborder                TYPE zexcel_s_cstylex_border
      CHANGING
        cs_complete_style_border  TYPE zexcel_s_cstyle_border
        cs_complete_stylex_border TYPE zexcel_s_cstylex_border.

    DATA: excel                   TYPE REF TO zcl_excel,
          lv_xborder_supplied     TYPE abap_bool,
          single_change_requested TYPE zexcel_s_cstylex_complete,
          BEGIN OF multiple_change_requested,
            complete   TYPE abap_bool,
            font       TYPE abap_bool,
            fill       TYPE abap_bool,
            BEGIN OF borders,
              complete   TYPE abap_bool,
              allborders TYPE abap_bool,
              diagonal   TYPE abap_bool,
              down       TYPE abap_bool,
              left       TYPE abap_bool,
              right      TYPE abap_bool,
              top        TYPE abap_bool,
            END OF borders,
            alignment  TYPE abap_bool,
            protection TYPE abap_bool,
          END OF multiple_change_requested.
    CONSTANTS:
          lv_border_supplied  TYPE abap_bool VALUE abap_true.
    ALIASES:
          complete_style   FOR zif_excel_style_changer~complete_style,
          complete_stylex  FOR zif_excel_style_changer~complete_stylex.

ENDCLASS.



CLASS ZCL_EXCEL_STYLE_CHANGER IMPLEMENTATION.


  METHOD clear_initial_colorxfields.

    IF is_color-rgb IS INITIAL.
      CLEAR cs_xcolor-rgb.
    ENDIF.
    IF is_color-indexed IS INITIAL.
      CLEAR cs_xcolor-indexed.
    ENDIF.
    IF is_color-theme IS INITIAL.
      CLEAR cs_xcolor-theme.
    ENDIF.
    IF is_color-tint IS INITIAL.
      CLEAR cs_xcolor-tint.
    ENDIF.

  ENDMETHOD.


  METHOD create.

    DATA: style TYPE REF TO zcl_excel_style_changer.

    CREATE OBJECT style.
    style->excel = excel.
    result = style.

  ENDMETHOD.


  METHOD move_supplied_borders.

    DATA: cs_borderx TYPE zexcel_s_cstylex_border.

    IF iv_border_supplied = abap_true.  " only act if parameter was supplied
      IF iv_xborder_supplied = abap_true. "
        cs_borderx = is_xborder.             " use supplied x-parameter
      ELSE.
        CLEAR cs_borderx WITH 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        IF is_border-border_style IS  INITIAL.
          CLEAR cs_borderx-border_style.
        ENDIF.
        clear_initial_colorxfields(
          EXPORTING
            is_color  = is_border-border_color
          CHANGING
            cs_xcolor = cs_borderx-border_color ).
      ENDIF.
      cs_complete_style_border = is_border.
      cs_complete_stylex_border = cs_borderx.
    ENDIF.

  ENDMETHOD.


  METHOD zif_excel_style_changer~apply.

    DATA: stylemapping TYPE zexcel_s_stylemapping,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          l_guid       TYPE zexcel_cell_style.

    lo_worksheet = excel->get_worksheet_by_name( ip_sheet_name = ip_worksheet->get_title( ) ).
    IF lo_worksheet <> ip_worksheet.
      zcx_excel=>raise_text( 'Worksheet doesn''t correspond to workbook of style changer'(001) ).
    ENDIF.

    TRY.
        ip_worksheet->get_cell( EXPORTING ip_column = ip_column
                                          ip_row    = ip_row
                                IMPORTING ep_guid   = l_guid ).
        stylemapping = excel->get_style_to_guid( l_guid ).
      CATCH zcx_excel.
* Error --> use submitted style
    ENDTRY.


    IF multiple_change_requested-complete = abap_true.
      stylemapping-complete_style = complete_style.
      stylemapping-complete_stylex = complete_stylex.
    ENDIF.

    IF multiple_change_requested-font = abap_true.
      stylemapping-complete_style-font = complete_style-font.
      stylemapping-complete_stylex-font = complete_stylex-font.
    ENDIF.

    IF multiple_change_requested-fill = abap_true.
      stylemapping-complete_style-fill = complete_style-fill.
      stylemapping-complete_stylex-fill = complete_stylex-fill.
    ENDIF.

    IF multiple_change_requested-borders-complete = abap_true.
      stylemapping-complete_style-borders = complete_style-borders.
      stylemapping-complete_stylex-borders = complete_stylex-borders.
    ENDIF.

    IF multiple_change_requested-borders-allborders = abap_true.
      stylemapping-complete_style-borders-allborders = complete_style-borders-allborders.
      stylemapping-complete_stylex-borders-allborders = complete_stylex-borders-allborders.
    ENDIF.

    IF multiple_change_requested-borders-diagonal = abap_true.
      stylemapping-complete_style-borders-diagonal = complete_style-borders-diagonal.
      stylemapping-complete_stylex-borders-diagonal = complete_stylex-borders-diagonal.
    ENDIF.

    IF multiple_change_requested-borders-down = abap_true.
      stylemapping-complete_style-borders-down = complete_style-borders-down.
      stylemapping-complete_stylex-borders-down = complete_stylex-borders-down.
    ENDIF.

    IF multiple_change_requested-borders-left = abap_true.
      stylemapping-complete_style-borders-left = complete_style-borders-left.
      stylemapping-complete_stylex-borders-left = complete_stylex-borders-left.
    ENDIF.

    IF multiple_change_requested-borders-right = abap_true.
      stylemapping-complete_style-borders-right = complete_style-borders-right.
      stylemapping-complete_stylex-borders-right = complete_stylex-borders-right.
    ENDIF.

    IF multiple_change_requested-borders-top = abap_true.
      stylemapping-complete_style-borders-top = complete_style-borders-top.
      stylemapping-complete_stylex-borders-top = complete_stylex-borders-top.
    ENDIF.

    IF multiple_change_requested-alignment = abap_true.
      stylemapping-complete_style-alignment = complete_style-alignment.
      stylemapping-complete_stylex-alignment = complete_stylex-alignment.
    ENDIF.

    IF multiple_change_requested-protection = abap_true.
      stylemapping-complete_style-protection = complete_style-protection.
      stylemapping-complete_stylex-protection = complete_stylex-protection.
    ENDIF.

    IF complete_stylex-number_format = abap_true.
      stylemapping-complete_style-number_format-format_code = complete_style-number_format-format_code.
      stylemapping-complete_stylex-number_format-format_code = abap_true.
    ENDIF.
    IF complete_stylex-font-bold = abap_true.
      stylemapping-complete_style-font-bold = complete_style-font-bold.
      stylemapping-complete_stylex-font-bold = complete_stylex-font-bold.
    ENDIF.
    IF complete_stylex-font-color = abap_true.
      stylemapping-complete_style-font-color = complete_style-font-color.
      stylemapping-complete_stylex-font-color = complete_stylex-font-color.
    ENDIF.
    IF complete_stylex-font-color-rgb = abap_true.
      stylemapping-complete_style-font-color-rgb = complete_style-font-color-rgb.
      stylemapping-complete_stylex-font-color-rgb = complete_stylex-font-color-rgb.
    ENDIF.
    IF complete_stylex-font-color-indexed = abap_true.
      stylemapping-complete_style-font-color-indexed = complete_style-font-color-indexed.
      stylemapping-complete_stylex-font-color-indexed = complete_stylex-font-color-indexed.
    ENDIF.
    IF complete_stylex-font-color-theme = abap_true.
      stylemapping-complete_style-font-color-theme = complete_style-font-color-theme.
      stylemapping-complete_stylex-font-color-theme = complete_stylex-font-color-theme.
    ENDIF.
    IF complete_stylex-font-color-tint = abap_true.
      stylemapping-complete_style-font-color-tint = complete_style-font-color-tint.
      stylemapping-complete_stylex-font-color-tint = complete_stylex-font-color-tint.
    ENDIF.
    IF complete_stylex-font-family = abap_true.
      stylemapping-complete_style-font-family = complete_style-font-family.
      stylemapping-complete_stylex-font-family = complete_stylex-font-family.
    ENDIF.
    IF complete_stylex-font-italic = abap_true.
      stylemapping-complete_style-font-italic = complete_style-font-italic.
      stylemapping-complete_stylex-font-italic = complete_stylex-font-italic.
    ENDIF.
    IF complete_stylex-font-name = abap_true.
      stylemapping-complete_style-font-name = complete_style-font-name.
      stylemapping-complete_stylex-font-name = complete_stylex-font-name.
    ENDIF.
    IF complete_stylex-font-scheme = abap_true.
      stylemapping-complete_style-font-scheme = complete_style-font-scheme.
      stylemapping-complete_stylex-font-scheme = complete_stylex-font-scheme.
    ENDIF.
    IF complete_stylex-font-size = abap_true.
      stylemapping-complete_style-font-size = complete_style-font-size.
      stylemapping-complete_stylex-font-size = complete_stylex-font-size.
    ENDIF.
    IF complete_stylex-font-strikethrough = abap_true.
      stylemapping-complete_style-font-strikethrough = complete_style-font-strikethrough.
      stylemapping-complete_stylex-font-strikethrough = complete_stylex-font-strikethrough.
    ENDIF.
    IF complete_stylex-font-underline = abap_true.
      stylemapping-complete_style-font-underline = complete_style-font-underline.
      stylemapping-complete_stylex-font-underline = complete_stylex-font-underline.
    ENDIF.
    IF complete_stylex-font-underline_mode = abap_true.
      stylemapping-complete_style-font-underline_mode = complete_style-font-underline_mode.
      stylemapping-complete_stylex-font-underline_mode = complete_stylex-font-underline_mode.
    ENDIF.

    IF complete_stylex-fill-filltype = abap_true.
      stylemapping-complete_style-fill-filltype = complete_style-fill-filltype.
      stylemapping-complete_stylex-fill-filltype = complete_stylex-fill-filltype.
    ENDIF.
    IF complete_stylex-fill-rotation = abap_true.
      stylemapping-complete_style-fill-rotation = complete_style-fill-rotation.
      stylemapping-complete_stylex-fill-rotation = complete_stylex-fill-rotation.
    ENDIF.
    IF complete_stylex-fill-fgcolor = abap_true.
      stylemapping-complete_style-fill-fgcolor = complete_style-fill-fgcolor.
      stylemapping-complete_stylex-fill-fgcolor = complete_stylex-fill-fgcolor.
    ENDIF.
    IF complete_stylex-fill-fgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-fgcolor-rgb = complete_style-fill-fgcolor-rgb.
      stylemapping-complete_stylex-fill-fgcolor-rgb = complete_stylex-fill-fgcolor-rgb.
    ENDIF.
    IF complete_stylex-fill-fgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-fgcolor-indexed = complete_style-fill-fgcolor-indexed.
      stylemapping-complete_stylex-fill-fgcolor-indexed = complete_stylex-fill-fgcolor-indexed.
    ENDIF.
    IF complete_stylex-fill-fgcolor-theme = abap_true.
      stylemapping-complete_style-fill-fgcolor-theme = complete_style-fill-fgcolor-theme.
      stylemapping-complete_stylex-fill-fgcolor-theme = complete_stylex-fill-fgcolor-theme.
    ENDIF.
    IF complete_stylex-fill-fgcolor-tint = abap_true.
      stylemapping-complete_style-fill-fgcolor-tint = complete_style-fill-fgcolor-tint.
      stylemapping-complete_stylex-fill-fgcolor-tint = complete_stylex-fill-fgcolor-tint.
    ENDIF.

    IF complete_stylex-fill-bgcolor = abap_true.
      stylemapping-complete_style-fill-bgcolor = complete_style-fill-bgcolor.
      stylemapping-complete_stylex-fill-bgcolor = complete_stylex-fill-bgcolor.
    ENDIF.
    IF complete_stylex-fill-bgcolor-rgb = abap_true.
      stylemapping-complete_style-fill-bgcolor-rgb = complete_style-fill-bgcolor-rgb.
      stylemapping-complete_stylex-fill-bgcolor-rgb = complete_stylex-fill-bgcolor-rgb.
    ENDIF.
    IF complete_stylex-fill-bgcolor-indexed = abap_true.
      stylemapping-complete_style-fill-bgcolor-indexed = complete_style-fill-bgcolor-indexed.
      stylemapping-complete_stylex-fill-bgcolor-indexed = complete_stylex-fill-bgcolor-indexed.
    ENDIF.
    IF complete_stylex-fill-bgcolor-theme = abap_true.
      stylemapping-complete_style-fill-bgcolor-theme = complete_style-fill-bgcolor-theme.
      stylemapping-complete_stylex-fill-bgcolor-theme = complete_stylex-fill-bgcolor-theme.
    ENDIF.
    IF complete_stylex-fill-bgcolor-tint = abap_true.
      stylemapping-complete_style-fill-bgcolor-tint = complete_style-fill-bgcolor-tint.
      stylemapping-complete_stylex-fill-bgcolor-tint = complete_stylex-fill-bgcolor-tint.
    ENDIF.

    IF complete_stylex-fill-gradtype-type = abap_true.
      stylemapping-complete_style-fill-gradtype-type = complete_style-fill-gradtype-type.
      stylemapping-complete_stylex-fill-gradtype-type = complete_stylex-fill-gradtype-type.
    ENDIF.
    IF complete_stylex-fill-gradtype-degree = abap_true.
      stylemapping-complete_style-fill-gradtype-degree = complete_style-fill-gradtype-degree.
      stylemapping-complete_stylex-fill-gradtype-degree = complete_stylex-fill-gradtype-degree.
    ENDIF.
    IF complete_stylex-fill-gradtype-bottom = abap_true.
      stylemapping-complete_style-fill-gradtype-bottom = complete_style-fill-gradtype-bottom.
      stylemapping-complete_stylex-fill-gradtype-bottom = complete_stylex-fill-gradtype-bottom.
    ENDIF.
    IF complete_stylex-fill-gradtype-left = abap_true.
      stylemapping-complete_style-fill-gradtype-left = complete_style-fill-gradtype-left.
      stylemapping-complete_stylex-fill-gradtype-left = complete_stylex-fill-gradtype-left.
    ENDIF.
    IF complete_stylex-fill-gradtype-top = abap_true.
      stylemapping-complete_style-fill-gradtype-top = complete_style-fill-gradtype-top.
      stylemapping-complete_stylex-fill-gradtype-top = complete_stylex-fill-gradtype-top.
    ENDIF.
    IF complete_stylex-fill-gradtype-right = abap_true.
      stylemapping-complete_style-fill-gradtype-right = complete_style-fill-gradtype-right.
      stylemapping-complete_stylex-fill-gradtype-right = complete_stylex-fill-gradtype-right.
    ENDIF.
    IF complete_stylex-fill-gradtype-position1 = abap_true.
      stylemapping-complete_style-fill-gradtype-position1 = complete_style-fill-gradtype-position1.
      stylemapping-complete_stylex-fill-gradtype-position1 = complete_stylex-fill-gradtype-position1.
    ENDIF.
    IF complete_stylex-fill-gradtype-position2 = abap_true.
      stylemapping-complete_style-fill-gradtype-position2 = complete_style-fill-gradtype-position2.
      stylemapping-complete_stylex-fill-gradtype-position2 = complete_stylex-fill-gradtype-position2.
    ENDIF.
    IF complete_stylex-fill-gradtype-position3 = abap_true.
      stylemapping-complete_style-fill-gradtype-position3 = complete_style-fill-gradtype-position3.
      stylemapping-complete_stylex-fill-gradtype-position3 = complete_stylex-fill-gradtype-position3.
    ENDIF.



    IF complete_stylex-borders-diagonal_mode = abap_true.
      stylemapping-complete_style-borders-diagonal_mode = complete_style-borders-diagonal_mode.
      stylemapping-complete_stylex-borders-diagonal_mode = complete_stylex-borders-diagonal_mode.
    ENDIF.
    IF complete_stylex-alignment-horizontal = abap_true.
      stylemapping-complete_style-alignment-horizontal = complete_style-alignment-horizontal.
      stylemapping-complete_stylex-alignment-horizontal = complete_stylex-alignment-horizontal.
    ENDIF.
    IF complete_stylex-alignment-vertical = abap_true.
      stylemapping-complete_style-alignment-vertical = complete_style-alignment-vertical.
      stylemapping-complete_stylex-alignment-vertical = complete_stylex-alignment-vertical.
    ENDIF.
    IF complete_stylex-alignment-textrotation = abap_true.
      stylemapping-complete_style-alignment-textrotation = complete_style-alignment-textrotation.
      stylemapping-complete_stylex-alignment-textrotation = complete_stylex-alignment-textrotation.
    ENDIF.
    IF complete_stylex-alignment-wraptext = abap_true.
      stylemapping-complete_style-alignment-wraptext = complete_style-alignment-wraptext.
      stylemapping-complete_stylex-alignment-wraptext = complete_stylex-alignment-wraptext.
    ENDIF.
    IF complete_stylex-alignment-shrinktofit = abap_true.
      stylemapping-complete_style-alignment-shrinktofit = complete_style-alignment-shrinktofit.
      stylemapping-complete_stylex-alignment-shrinktofit = complete_stylex-alignment-shrinktofit.
    ENDIF.
    IF complete_stylex-alignment-indent = abap_true.
      stylemapping-complete_style-alignment-indent = complete_style-alignment-indent.
      stylemapping-complete_stylex-alignment-indent = complete_stylex-alignment-indent.
    ENDIF.
    IF complete_stylex-protection-hidden = abap_true.
      stylemapping-complete_style-protection-hidden = complete_style-protection-hidden.
      stylemapping-complete_stylex-protection-hidden = complete_stylex-protection-hidden.
    ENDIF.
    IF complete_stylex-protection-locked = abap_true.
      stylemapping-complete_style-protection-locked = complete_style-protection-locked.
      stylemapping-complete_stylex-protection-locked = complete_stylex-protection-locked.
    ENDIF.

    IF complete_stylex-borders-allborders-border_style = abap_true.
      stylemapping-complete_style-borders-allborders-border_style = complete_style-borders-allborders-border_style.
      stylemapping-complete_stylex-borders-allborders-border_style = complete_stylex-borders-allborders-border_style.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-rgb = complete_style-borders-allborders-border_color-rgb.
      stylemapping-complete_stylex-borders-allborders-border_color-rgb = complete_stylex-borders-allborders-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-indexed = complete_style-borders-allborders-border_color-indexed.
      stylemapping-complete_stylex-borders-allborders-border_color-indexed = complete_stylex-borders-allborders-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-theme = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-theme = complete_style-borders-allborders-border_color-theme.
      stylemapping-complete_stylex-borders-allborders-border_color-theme = complete_stylex-borders-allborders-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-allborders-border_color-tint = abap_true.
      stylemapping-complete_style-borders-allborders-border_color-tint = complete_style-borders-allborders-border_color-tint.
      stylemapping-complete_stylex-borders-allborders-border_color-tint = complete_stylex-borders-allborders-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-diagonal-border_style = abap_true.
      stylemapping-complete_style-borders-diagonal-border_style = complete_style-borders-diagonal-border_style.
      stylemapping-complete_stylex-borders-diagonal-border_style = complete_stylex-borders-diagonal-border_style.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-rgb = complete_style-borders-diagonal-border_color-rgb.
      stylemapping-complete_stylex-borders-diagonal-border_color-rgb = complete_stylex-borders-diagonal-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-indexed = complete_style-borders-diagonal-border_color-indexed.
      stylemapping-complete_stylex-borders-diagonal-border_color-indexed = complete_stylex-borders-diagonal-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-theme = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-theme = complete_style-borders-diagonal-border_color-theme.
      stylemapping-complete_stylex-borders-diagonal-border_color-theme = complete_stylex-borders-diagonal-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-diagonal-border_color-tint = abap_true.
      stylemapping-complete_style-borders-diagonal-border_color-tint = complete_style-borders-diagonal-border_color-tint.
      stylemapping-complete_stylex-borders-diagonal-border_color-tint = complete_stylex-borders-diagonal-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-down-border_style = abap_true.
      stylemapping-complete_style-borders-down-border_style = complete_style-borders-down-border_style.
      stylemapping-complete_stylex-borders-down-border_style = complete_stylex-borders-down-border_style.
    ENDIF.
    IF complete_stylex-borders-down-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-down-border_color-rgb = complete_style-borders-down-border_color-rgb.
      stylemapping-complete_stylex-borders-down-border_color-rgb = complete_stylex-borders-down-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-down-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-down-border_color-indexed = complete_style-borders-down-border_color-indexed.
      stylemapping-complete_stylex-borders-down-border_color-indexed = complete_stylex-borders-down-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-down-border_color-theme = abap_true.
      stylemapping-complete_style-borders-down-border_color-theme = complete_style-borders-down-border_color-theme.
      stylemapping-complete_stylex-borders-down-border_color-theme = complete_stylex-borders-down-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-down-border_color-tint = abap_true.
      stylemapping-complete_style-borders-down-border_color-tint = complete_style-borders-down-border_color-tint.
      stylemapping-complete_stylex-borders-down-border_color-tint = complete_stylex-borders-down-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-left-border_style = abap_true.
      stylemapping-complete_style-borders-left-border_style = complete_style-borders-left-border_style.
      stylemapping-complete_stylex-borders-left-border_style = complete_stylex-borders-left-border_style.
    ENDIF.
    IF complete_stylex-borders-left-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-left-border_color-rgb = complete_style-borders-left-border_color-rgb.
      stylemapping-complete_stylex-borders-left-border_color-rgb = complete_stylex-borders-left-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-left-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-left-border_color-indexed = complete_style-borders-left-border_color-indexed.
      stylemapping-complete_stylex-borders-left-border_color-indexed = complete_stylex-borders-left-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-left-border_color-theme = abap_true.
      stylemapping-complete_style-borders-left-border_color-theme = complete_style-borders-left-border_color-theme.
      stylemapping-complete_stylex-borders-left-border_color-theme = complete_stylex-borders-left-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-left-border_color-tint = abap_true.
      stylemapping-complete_style-borders-left-border_color-tint = complete_style-borders-left-border_color-tint.
      stylemapping-complete_stylex-borders-left-border_color-tint = complete_stylex-borders-left-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-right-border_style = abap_true.
      stylemapping-complete_style-borders-right-border_style = complete_style-borders-right-border_style.
      stylemapping-complete_stylex-borders-right-border_style = complete_stylex-borders-right-border_style.
    ENDIF.
    IF complete_stylex-borders-right-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-right-border_color-rgb = complete_style-borders-right-border_color-rgb.
      stylemapping-complete_stylex-borders-right-border_color-rgb = complete_stylex-borders-right-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-right-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-right-border_color-indexed = complete_style-borders-right-border_color-indexed.
      stylemapping-complete_stylex-borders-right-border_color-indexed = complete_stylex-borders-right-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-right-border_color-theme = abap_true.
      stylemapping-complete_style-borders-right-border_color-theme = complete_style-borders-right-border_color-theme.
      stylemapping-complete_stylex-borders-right-border_color-theme = complete_stylex-borders-right-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-right-border_color-tint = abap_true.
      stylemapping-complete_style-borders-right-border_color-tint = complete_style-borders-right-border_color-tint.
      stylemapping-complete_stylex-borders-right-border_color-tint = complete_stylex-borders-right-border_color-tint.
    ENDIF.

    IF complete_stylex-borders-top-border_style = abap_true.
      stylemapping-complete_style-borders-top-border_style = complete_style-borders-top-border_style.
      stylemapping-complete_stylex-borders-top-border_style = complete_stylex-borders-top-border_style.
    ENDIF.
    IF complete_stylex-borders-top-border_color-rgb = abap_true.
      stylemapping-complete_style-borders-top-border_color-rgb = complete_style-borders-top-border_color-rgb.
      stylemapping-complete_stylex-borders-top-border_color-rgb = complete_stylex-borders-top-border_color-rgb.
    ENDIF.
    IF complete_stylex-borders-top-border_color-indexed = abap_true.
      stylemapping-complete_style-borders-top-border_color-indexed = complete_style-borders-top-border_color-indexed.
      stylemapping-complete_stylex-borders-top-border_color-indexed = complete_stylex-borders-top-border_color-indexed.
    ENDIF.
    IF complete_stylex-borders-top-border_color-theme = abap_true.
      stylemapping-complete_style-borders-top-border_color-theme = complete_style-borders-top-border_color-theme.
      stylemapping-complete_stylex-borders-top-border_color-theme = complete_stylex-borders-top-border_color-theme.
    ENDIF.
    IF complete_stylex-borders-top-border_color-tint = abap_true.
      stylemapping-complete_style-borders-top-border_color-tint = complete_style-borders-top-border_color-tint.
      stylemapping-complete_stylex-borders-top-border_color-tint = complete_stylex-borders-top-border_color-tint.
    ENDIF.


* Now we have a completly filled styles.
* This can be used to get the guid
* Return guid if requested.  Might be used if copy&paste of styles is requested
    ep_guid = me->excel->get_static_cellstyle_guid( ip_cstyle_complete  = stylemapping-complete_style
                                                   ip_cstylex_complete = stylemapping-complete_stylex  ).
    lo_worksheet->set_cell_style( ip_column = ip_column
                                  ip_row    = ip_row
                                  ip_style  = ep_guid ).

  ENDMETHOD.


  METHOD zif_excel_style_changer~get_guid.

    result = excel->get_static_cellstyle_guid( ip_cstyle_complete  = complete_style
                                               ip_cstylex_complete = complete_stylex  ).

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_horizontal.

    complete_style-alignment-horizontal = value.
    complete_stylex-alignment-horizontal = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_indent.

    complete_style-alignment-indent = value.
    complete_stylex-alignment-indent = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_shrinktofit.

    complete_style-alignment-shrinktofit = value.
    complete_stylex-alignment-shrinktofit = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_textrotation.

    complete_style-alignment-textrotation = value.
    complete_stylex-alignment-textrotation = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_vertical.

    complete_style-alignment-vertical = value.
    complete_stylex-alignment-vertical = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_alignment_wraptext.

    complete_style-alignment-wraptext = value.
    complete_stylex-alignment-wraptext = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allborders_color.

    complete_style-borders-allborders-border_color = value.
    complete_stylex-borders-allborders-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allborders_style.

    complete_style-borders-allborders-border_style = value.
    complete_stylex-borders-allborders-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allbo_color_indexe.

    complete_style-borders-allborders-border_color-indexed = value.
    complete_stylex-borders-allborders-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allbo_color_rgb.

    complete_style-borders-allborders-border_color-rgb = value.
    complete_stylex-borders-allborders-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allbo_color_theme.

    complete_style-borders-allborders-border_color-theme = value.
    complete_stylex-borders-allborders-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_allbo_color_tint.

    complete_style-borders-allborders-border_color-tint = value.
    complete_stylex-borders-allborders-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_color.

    complete_style-borders-diagonal-border_color = value.
    complete_stylex-borders-diagonal-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_color_ind.

    complete_style-borders-diagonal-border_color-indexed = value.
    complete_stylex-borders-diagonal-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_color_rgb.

    complete_style-borders-diagonal-border_color-rgb = value.
    complete_stylex-borders-diagonal-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_color_the.

    complete_style-borders-diagonal-border_color-theme = value.
    complete_stylex-borders-diagonal-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_color_tin.

    complete_style-borders-diagonal-border_color-tint = value.
    complete_stylex-borders-diagonal-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_mode.

    complete_style-borders-diagonal_mode = value.
    complete_stylex-borders-diagonal_mode = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_diagonal_style.

    complete_style-borders-diagonal-border_style = value.
    complete_stylex-borders-diagonal-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_color.

    complete_style-borders-down-border_color = value.
    complete_stylex-borders-down-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_color_indexed.

    complete_style-borders-down-border_color-indexed = value.
    complete_stylex-borders-down-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_color_rgb.

    complete_style-borders-down-border_color-rgb = value.
    complete_stylex-borders-down-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_color_theme.

    complete_style-borders-down-border_color-theme = value.
    complete_stylex-borders-down-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_color_tint.

    complete_style-borders-down-border_color-tint = value.
    complete_stylex-borders-down-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_down_style.

    complete_style-borders-down-border_style = value.
    complete_stylex-borders-down-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_color.

    complete_style-borders-left-border_color = value.
    complete_stylex-borders-left-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_color_indexed.

    complete_style-borders-left-border_color-indexed = value.
    complete_stylex-borders-left-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_color_rgb.

    complete_style-borders-left-border_color-rgb = value.
    complete_stylex-borders-left-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_color_theme.

    complete_style-borders-left-border_color-theme = value.
    complete_stylex-borders-left-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_color_tint.

    complete_style-borders-left-border_color-tint = value.
    complete_stylex-borders-left-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_left_style.

    complete_style-borders-left-border_style = value.
    complete_stylex-borders-left-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_color.

    complete_style-borders-right-border_color = value.
    complete_stylex-borders-right-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_color_indexe.

    complete_style-borders-right-border_color-indexed = value.
    complete_stylex-borders-right-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_color_rgb.

    complete_style-borders-right-border_color-rgb = value.
    complete_stylex-borders-right-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_color_theme.

    complete_style-borders-right-border_color-theme = value.
    complete_stylex-borders-right-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_color_tint.

    complete_style-borders-right-border_color-tint = value.
    complete_stylex-borders-right-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_right_style.

    complete_style-borders-right-border_style = value.
    complete_stylex-borders-right-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_color.

    complete_style-borders-top-border_color = value.
    complete_stylex-borders-top-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_color_indexed.

    complete_style-borders-top-border_color-indexed = value.
    complete_stylex-borders-top-border_color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_color_rgb.

    complete_style-borders-top-border_color-rgb = value.
    complete_stylex-borders-top-border_color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_color_theme.

    complete_style-borders-top-border_color-theme = value.
    complete_stylex-borders-top-border_color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_color_tint.

    complete_style-borders-top-border_color-tint = value.
    complete_stylex-borders-top-border_color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_borders_top_style.

    complete_style-borders-top-border_style = value.
    complete_stylex-borders-top-border_style = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete.

    complete_style = ip_complete.
    complete_stylex = ip_xcomplete.
    multiple_change_requested-complete = abap_true.
    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_alignment.

    DATA: alignmentx LIKE ip_xalignment.

    IF ip_xalignment IS SUPPLIED.
      alignmentx = ip_xalignment.
    ELSE.
      CLEAR alignmentx WITH 'X'.
      IF ip_alignment-horizontal IS INITIAL.
        CLEAR alignmentx-horizontal.
      ENDIF.
      IF ip_alignment-vertical IS INITIAL.
        CLEAR alignmentx-vertical.
      ENDIF.
    ENDIF.

    complete_style-alignment  = ip_alignment .
    complete_stylex-alignment = alignmentx   .
    multiple_change_requested-alignment = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders.

    DATA: bordersx LIKE ip_xborders.
    IF ip_xborders IS SUPPLIED.
      bordersx = ip_xborders.
    ELSE.
      CLEAR bordersx WITH 'X'.
      IF ip_borders-allborders-border_style IS INITIAL.
        CLEAR bordersx-allborders-border_style.
      ENDIF.
      IF ip_borders-diagonal-border_style IS INITIAL.
        CLEAR bordersx-diagonal-border_style.
      ENDIF.
      IF ip_borders-down-border_style IS INITIAL.
        CLEAR bordersx-down-border_style.
      ENDIF.
      IF ip_borders-left-border_style IS INITIAL.
        CLEAR bordersx-left-border_style.
      ENDIF.
      IF ip_borders-right-border_style IS INITIAL.
        CLEAR bordersx-right-border_style.
      ENDIF.
      IF ip_borders-top-border_style IS INITIAL.
        CLEAR bordersx-top-border_style.
      ENDIF.

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-allborders-border_color
        CHANGING
          cs_xcolor = bordersx-allborders-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-diagonal-border_color
        CHANGING
          cs_xcolor = bordersx-diagonal-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-down-border_color
        CHANGING
          cs_xcolor = bordersx-down-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-left-border_color
        CHANGING
          cs_xcolor = bordersx-left-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-right-border_color
        CHANGING
          cs_xcolor = bordersx-right-border_color ).

      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_borders-top-border_color
        CHANGING
          cs_xcolor = bordersx-top-border_color ).

    ENDIF.

    complete_style-borders = ip_borders.
    complete_stylex-borders = bordersx.
    multiple_change_requested-borders-complete = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_all.

    lv_xborder_supplied = boolc( ip_xborders_allborders IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_allborders
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_allborders
      CHANGING
        cs_complete_style_border  = complete_style-borders-allborders
        cs_complete_stylex_border = complete_stylex-borders-allborders ).
    multiple_change_requested-borders-allborders = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_diagonal.

    lv_xborder_supplied = boolc( ip_xborders_diagonal IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_diagonal
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_diagonal
      CHANGING
        cs_complete_style_border  = complete_style-borders-diagonal
        cs_complete_stylex_border = complete_stylex-borders-diagonal ).
    multiple_change_requested-borders-diagonal = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_down.

    lv_xborder_supplied = boolc( ip_xborders_down IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_down
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_down
      CHANGING
        cs_complete_style_border  = complete_style-borders-down
        cs_complete_stylex_border = complete_stylex-borders-down ).
    multiple_change_requested-borders-down = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_left.

    lv_xborder_supplied = boolc( ip_xborders_left IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_left
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_left
      CHANGING
        cs_complete_style_border  = complete_style-borders-left
        cs_complete_stylex_border = complete_stylex-borders-left ).
    multiple_change_requested-borders-left = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_right.

    lv_xborder_supplied = boolc( ip_xborders_right IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_right
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_right
      CHANGING
        cs_complete_style_border  = complete_style-borders-right
        cs_complete_stylex_border = complete_stylex-borders-right ).
    multiple_change_requested-borders-right = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_borders_top.

    lv_xborder_supplied = boolc( ip_xborders_top IS SUPPLIED ).
    move_supplied_borders(
      EXPORTING
        iv_border_supplied        = lv_border_supplied
        is_border                 = ip_borders_top
        iv_xborder_supplied       = lv_xborder_supplied
        is_xborder                = ip_xborders_top
      CHANGING
        cs_complete_style_border  = complete_style-borders-top
        cs_complete_stylex_border = complete_stylex-borders-top ).
    multiple_change_requested-borders-top = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_fill.

    DATA: fillx LIKE ip_xfill.
    IF ip_xfill IS SUPPLIED.
      fillx = ip_xfill.
    ELSE.
      CLEAR fillx WITH 'X'.
      IF ip_fill-filltype IS INITIAL.
        CLEAR fillx-filltype.
      ENDIF.
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_fill-fgcolor
        CHANGING
          cs_xcolor = fillx-fgcolor ).
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_fill-bgcolor
        CHANGING
          cs_xcolor = fillx-bgcolor ).

    ENDIF.

    complete_style-fill = ip_fill.
    complete_stylex-fill = fillx.
    multiple_change_requested-fill = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_font.

    DATA: fontx TYPE zexcel_s_cstylex_font.

    IF ip_xfont IS SUPPLIED.
      fontx = ip_xfont.
    ELSE.
* Only supplied values should be used - exception: Flags bold and italic strikethrough underline
      fontx-bold = 'X'.
      fontx-italic = 'X'.
      fontx-strikethrough = 'X'.
      fontx-underline_mode = 'X'.
      CLEAR fontx-color WITH 'X'.
      clear_initial_colorxfields(
        EXPORTING
          is_color  = ip_font-color
        CHANGING
          cs_xcolor = fontx-color ).
      IF ip_font-family IS NOT INITIAL.
        fontx-family = 'X'.
      ENDIF.
      IF ip_font-name IS NOT INITIAL.
        fontx-name = 'X'.
      ENDIF.
      IF ip_font-scheme IS NOT INITIAL.
        fontx-scheme = 'X'.
      ENDIF.
      IF ip_font-size IS NOT INITIAL.
        fontx-size = 'X'.
      ENDIF.
      IF ip_font-underline_mode IS NOT INITIAL.
        fontx-underline_mode = 'X'.
      ENDIF.
    ENDIF.

    complete_style-font = ip_font.
    complete_stylex-font = fontx.
    multiple_change_requested-font = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_complete_protection.

    MOVE-CORRESPONDING ip_protection  TO complete_style-protection.
    IF ip_xprotection IS SUPPLIED.
      MOVE-CORRESPONDING ip_xprotection TO complete_stylex-protection.
    ELSE.
      IF ip_protection-hidden IS NOT INITIAL.
        complete_stylex-protection-hidden = 'X'.
      ENDIF.
      IF ip_protection-locked IS NOT INITIAL.
        complete_stylex-protection-locked = 'X'.
      ENDIF.
    ENDIF.
    multiple_change_requested-protection = abap_true.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_bgcolor.

    complete_style-fill-bgcolor = value.
    complete_stylex-fill-bgcolor-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_bgcolor_indexed.

    complete_style-fill-bgcolor-indexed = value.
    complete_stylex-fill-bgcolor-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_bgcolor_rgb.

    complete_style-fill-bgcolor-rgb = value.
    complete_stylex-fill-bgcolor-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_bgcolor_theme.

    complete_style-fill-bgcolor-theme = value.
    complete_stylex-fill-bgcolor-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_bgcolor_tint.

    complete_style-fill-bgcolor-tint = value.
    complete_stylex-fill-bgcolor-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_fgcolor.

    complete_style-fill-fgcolor = value.
    complete_stylex-fill-fgcolor-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_fgcolor_indexed.

    complete_style-fill-fgcolor-indexed = value.
    complete_stylex-fill-fgcolor-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_fgcolor_rgb.

    complete_style-fill-fgcolor-rgb = value.
    complete_stylex-fill-fgcolor-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_fgcolor_theme.

    complete_style-fill-fgcolor-theme = value.
    complete_stylex-fill-fgcolor-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_fgcolor_tint.

    complete_style-fill-fgcolor-tint = value.
    complete_stylex-fill-fgcolor-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_filltype.

    complete_style-fill-filltype = value.
    complete_stylex-fill-filltype = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_bottom.

    complete_style-fill-gradtype-bottom = value.
    complete_stylex-fill-gradtype-bottom = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_degree.

    complete_style-fill-gradtype-degree = value.
    complete_stylex-fill-gradtype-degree = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_left.

    complete_style-fill-gradtype-left = value.
    complete_stylex-fill-gradtype-left = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_position1.

    complete_style-fill-gradtype-position1 = value.
    complete_stylex-fill-gradtype-position1 = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_position2.

    complete_style-fill-gradtype-position2 = value.
    complete_stylex-fill-gradtype-position2 = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_position3.

    complete_style-fill-gradtype-position3 = value.
    complete_stylex-fill-gradtype-position3 = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_right.

    complete_style-fill-gradtype-right = value.
    complete_stylex-fill-gradtype-right = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_top.

    complete_style-fill-gradtype-top = value.
    complete_stylex-fill-gradtype-top = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_gradtype_type.

    complete_style-fill-gradtype-type = value.
    complete_stylex-fill-gradtype-type = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_fill_rotation.

    complete_style-fill-rotation = value.
    complete_stylex-fill-rotation = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_bold.

    complete_style-font-bold = value.
    complete_stylex-font-bold = 'X'.
    single_change_requested-font-bold = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_color.

    complete_style-font-color = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_color_indexed.

    complete_style-font-color-indexed = value.
    complete_stylex-font-color-indexed = 'X'.
    single_change_requested-font-color-indexed = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_color_rgb.

    complete_style-font-color-rgb = value.
    complete_stylex-font-color-rgb = 'X'.
    single_change_requested-font-color-rgb = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_color_theme.

    complete_style-font-color-theme = value.
    complete_stylex-font-color-theme = 'X'.
    single_change_requested-font-color-theme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_color_tint.

    complete_style-font-color-tint = value.
    complete_stylex-font-color-tint = 'X'.
    single_change_requested-font-color-tint = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_family.

    complete_style-font-family = value.
    complete_stylex-font-family = 'X'.
    single_change_requested-font-family = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_italic.

    complete_style-font-italic = value.
    complete_stylex-font-italic = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_name.

    complete_style-font-name = value.
    complete_stylex-font-name = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_scheme.

    complete_style-font-scheme = value.
    complete_stylex-font-scheme = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_size.

    complete_style-font-size = value.
    complete_stylex-font-size = abap_true.
    single_change_requested-font-size = abap_true.
    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_strikethrough.

    complete_style-font-strikethrough = value.
    complete_stylex-font-strikethrough = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_underline.

    complete_style-font-underline = value.
    complete_stylex-font-underline = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_font_underline_mode.

    complete_style-font-underline_mode = value.
    complete_stylex-font-underline_mode = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_number_format.

    complete_style-number_format-format_code = value.
    complete_stylex-number_format-format_code = abap_true.
    single_change_requested-number_format-format_code = abap_true.
    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_protection_hidden.

    complete_style-protection-hidden = value.
    complete_stylex-protection-hidden = 'X'.

    result = me.

  ENDMETHOD.


  METHOD zif_excel_style_changer~set_protection_locked.

    complete_style-protection-locked = value.
    complete_stylex-protection-locked = 'X'.

    result = me.

  ENDMETHOD.
ENDCLASS.
