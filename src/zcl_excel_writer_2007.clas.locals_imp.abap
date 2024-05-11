*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations
CLASS lcl_create_xl_sheet DEFINITION DEFERRED.
CLASS zcl_excel_writer_2007 DEFINITION LOCAL FRIENDS lcl_create_xl_sheet.
CLASS lcl_create_xl_sheet DEFINITION CREATE PUBLIC .

  PUBLIC SECTION.
    METHODS create IMPORTING io_worksheet         TYPE REF TO zcl_excel_worksheet
                             iv_active            TYPE flag DEFAULT ''
                             io_document          TYPE REF TO if_ixml_document
                             io_excel_writer_2007 TYPE REF TO zcl_excel_writer_2007
                   RAISING   zcx_excel.
  PROTECTED SECTION.
  PRIVATE SECTION.
    TYPES: BEGIN OF colors,
             colorrgb TYPE zexcel_color,
           END OF colors,
           BEGIN OF cfvo,
             value TYPE zexcel_conditional_value,
             type  TYPE zexcel_conditional_type,
           END OF cfvo,
           BEGIN OF ty_condformating_range,
             dimension_range     TYPE string,
             condformatting_node TYPE REF TO if_ixml_element,
           END OF ty_condformating_range,
           ty_condformating_ranges TYPE STANDARD TABLE OF ty_condformating_range.
    CONSTANTS:
      lc_xml_node_sheetpr           TYPE string VALUE 'sheetPr',
      lc_xml_node_tabcolor          TYPE string VALUE 'tabColor',
      lc_xml_node_outlinepr         TYPE string VALUE 'outlinePr',
      lc_xml_node_pagesetuppr       TYPE string VALUE 'pageSetUpPr',
      lc_xml_node_dimension         TYPE string VALUE 'dimension',
      lc_xml_node_sheetviews        TYPE string VALUE 'sheetViews',
      lc_xml_node_sheetview         TYPE string VALUE 'sheetView',
      lc_xml_node_selection         TYPE string VALUE 'selection',
      lc_xml_node_pane              TYPE string VALUE 'pane',
      lc_xml_node_sheetformatpr     TYPE string VALUE 'sheetFormatPr',
      lc_xml_node_cols              TYPE string VALUE 'cols',
      lc_xml_node_col               TYPE string VALUE 'col',
      lc_xml_node_sheetprotection   TYPE string VALUE 'sheetProtection',
      lc_xml_node_autofilter        TYPE string VALUE 'autoFilter',
      lc_xml_node_filtercolumn      TYPE string VALUE 'filterColumn',
      lc_xml_node_filters           TYPE string VALUE 'filters',
      lc_xml_node_filter            TYPE string VALUE 'filter',
      lc_xml_node_mergecell         TYPE string VALUE 'mergeCell',
      lc_xml_node_mergecells        TYPE string VALUE 'mergeCells',
      lc_xml_node_condformatting    TYPE string VALUE 'conditionalFormatting',
      lc_xml_node_cfrule            TYPE string VALUE 'cfRule',
      lc_xml_node_color             TYPE string VALUE 'color',      " Databar by Albert Lladanosa
      lc_xml_node_databar           TYPE string VALUE 'dataBar',    " Databar by Albert Lladanosa
      lc_xml_node_colorscale        TYPE string VALUE 'colorScale',
      lc_xml_node_iconset           TYPE string VALUE 'iconSet',
      lc_xml_node_cfvo              TYPE string VALUE 'cfvo',
      lc_xml_node_formula           TYPE string VALUE 'formula',
      lc_xml_node_datavalidations   TYPE string VALUE 'dataValidations',
      lc_xml_node_datavalidation    TYPE string VALUE 'dataValidation',
      lc_xml_node_formula1          TYPE string VALUE 'formula1',
      lc_xml_node_formula2          TYPE string VALUE 'formula2',
      lc_xml_node_pagemargins       TYPE string VALUE 'pageMargins',
      lc_xml_node_pagesetup         TYPE string VALUE 'pageSetup',
      lc_xml_node_headerfooter      TYPE string VALUE 'headerFooter',
      lc_xml_node_oddheader         TYPE string VALUE 'oddHeader',
      lc_xml_node_oddfooter         TYPE string VALUE 'oddFooter',
      lc_xml_node_evenheader        TYPE string VALUE 'evenHeader',
      lc_xml_node_evenfooter        TYPE string VALUE 'evenFooter',
      lc_xml_node_drawing           TYPE string VALUE 'drawing',
      lc_xml_node_drawing_for_cmt   TYPE string VALUE 'legacyDrawing',
      lc_xml_node_drawing_for_hd_ft TYPE string VALUE 'legacyDrawingHF'.
    CONSTANTS:
      lc_xml_attr_summarybelow       TYPE string VALUE 'summaryBelow',
      lc_xml_attr_summaryright       TYPE string VALUE 'summaryRight',
      lc_xml_attr_fittopage          TYPE string VALUE 'fitToPage',
      lc_xml_attr_tabcolor_rgb       TYPE string VALUE 'rgb',
      lc_xml_attr_ref                TYPE string VALUE 'ref',
      lc_xml_attr_tabselected        TYPE string VALUE 'tabSelected',
      lc_xml_attr_showzeros          TYPE string VALUE 'showZeros',
      lc_xml_attr_zoomscale          TYPE string VALUE 'zoomScale',
      lc_xml_attr_zoomscalenormal    TYPE string VALUE 'zoomScaleNormal',
      lc_xml_attr_zoomscalepageview  TYPE string VALUE 'zoomScalePageLayoutView',
      lc_xml_attr_zoomscalesheetview TYPE string VALUE 'zoomScaleSheetLayoutView',
      lc_xml_attr_workbookviewid     TYPE string VALUE 'workbookViewId',
      lc_xml_attr_showgridlines      TYPE string VALUE 'showGridLines',
      lc_xml_attr_showrowcolheaders  TYPE string VALUE 'showRowColHeaders',
      lc_xml_attr_activecell         TYPE string VALUE 'activeCell',
      lc_xml_attr_sqref              TYPE string VALUE 'sqref',
      lc_xml_attr_true               TYPE string VALUE 'true',
      lc_xml_attr_customheight       TYPE string VALUE 'customHeight',
      lc_xml_attr_defaultrowheight   TYPE string VALUE 'defaultRowHeight',
      lc_xml_attr_defaultcolwidth    TYPE string VALUE 'defaultColWidth',
      lc_xml_attr_outlinelevelcol    TYPE string VALUE 'x14ac:outlineLevelCol',
      lc_xml_attr_min                TYPE string VALUE 'min',
      lc_xml_attr_max                TYPE string VALUE 'max',
      lc_xml_attr_hidden             TYPE string VALUE 'hidden',
      lc_xml_attr_width              TYPE string VALUE 'width',
      lc_xml_attr_defaultwidth       TYPE string VALUE '9.10',
      lc_xml_attr_style              TYPE string VALUE 'style',
      lc_xml_attr_bestfit            TYPE string VALUE 'bestFit',
      lc_xml_attr_customwidth        TYPE string VALUE 'customWidth',
      lc_xml_attr_collapsed          TYPE string VALUE 'collapsed',
      lc_xml_attr_outlinelevel       TYPE string VALUE 'outlineLevel',
      lc_xml_attr_password           TYPE string VALUE 'password',
      lc_xml_attr_sheet              TYPE string VALUE 'sheet',
      lc_xml_attr_objects            TYPE string VALUE 'objects',
      lc_xml_attr_scenarios          TYPE string VALUE 'scenarios',
      lc_xml_attr_autofilter         TYPE string VALUE 'autoFilter',
      lc_xml_attr_deletecolumns      TYPE string VALUE 'deleteColumns',
      lc_xml_attr_deleterows         TYPE string VALUE 'deleteRows',
      lc_xml_attr_formatcells        TYPE string VALUE 'formatCells',
      lc_xml_attr_formatcolumns      TYPE string VALUE 'formatColumns',
      lc_xml_attr_formatrows         TYPE string VALUE 'formatRows',
      lc_xml_attr_insertcolumns      TYPE string VALUE 'insertColumns',
      lc_xml_attr_inserthyperlinks   TYPE string VALUE 'insertHyperlinks',
      lc_xml_attr_insertrows         TYPE string VALUE 'insertRows',
      lc_xml_attr_pivottables        TYPE string VALUE 'pivotTables',
      lc_xml_attr_selectlockedcells  TYPE string VALUE 'selectLockedCells',
      lc_xml_attr_selectunlockedcell TYPE string VALUE 'selectUnlockedCells',
      lc_xml_attr_sort               TYPE string VALUE 'sort',
      lc_xml_attr_val                TYPE string VALUE 'val',
      lc_xml_attr_colid              TYPE string VALUE 'colId',
      lc_xml_attr_filtermode         TYPE string VALUE 'filterMode',
      lc_xml_attr_count              TYPE string VALUE 'count',
      lc_xml_attr_type               TYPE string VALUE 'type',
      lc_xml_attr_operator           TYPE string VALUE 'operator',
      lc_xml_attr_allowblank         TYPE string VALUE 'allowBlank',
      lc_xml_attr_showinputmessage   TYPE string VALUE 'showInputMessage',
      lc_xml_attr_showerrormessage   TYPE string VALUE 'showErrorMessage',
      lc_xml_attr_showdropdown       TYPE string VALUE 'ShowDropDown', " 'showDropDown' does not work
      lc_xml_attr_errortitle         TYPE string VALUE 'errorTitle',
      lc_xml_attr_error              TYPE string VALUE 'error',
      lc_xml_attr_errorstyle         TYPE string VALUE 'errorStyle',
      lc_xml_attr_prompttitle        TYPE string VALUE 'promptTitle',
      lc_xml_attr_prompt             TYPE string VALUE 'prompt',
      lc_xml_attr_gridlines          TYPE string VALUE 'gridLines',
      lc_xml_attr_left               TYPE string VALUE 'left',
      lc_xml_attr_right              TYPE string VALUE 'right',
      lc_xml_attr_top                TYPE string VALUE 'top',
      lc_xml_attr_bottom             TYPE string VALUE 'bottom',
      lc_xml_attr_header             TYPE string VALUE 'header',
      lc_xml_attr_footer             TYPE string VALUE 'footer',
      lc_xml_attr_blackandwhite      TYPE string VALUE 'blackAndWhite',
      lc_xml_attr_cellcomments       TYPE string VALUE 'cellComments',
      lc_xml_attr_copies             TYPE string VALUE 'copies',
      lc_xml_attr_draft              TYPE string VALUE 'draft',
      lc_xml_attr_errors             TYPE string VALUE 'errors',
      lc_xml_attr_firstpagenumber    TYPE string VALUE 'firstPageNumber',
      lc_xml_attr_fittoheight        TYPE string VALUE 'fitToHeight',
      lc_xml_attr_fittowidth         TYPE string VALUE 'fitToWidth',
      lc_xml_attr_horizontaldpi      TYPE string VALUE 'horizontalDpi',
      lc_xml_attr_orientation        TYPE string VALUE 'orientation',
      lc_xml_attr_pageorder          TYPE string VALUE 'pageOrder',
      lc_xml_attr_paperheight        TYPE string VALUE 'paperHeight',
      lc_xml_attr_papersize          TYPE string VALUE 'paperSize',
      lc_xml_attr_paperwidth         TYPE string VALUE 'paperWidth',
      lc_xml_attr_scale              TYPE string VALUE 'scale',
      lc_xml_attr_usefirstpagenumber TYPE string VALUE 'useFirstPageNumber',
      lc_xml_attr_useprinterdefaults TYPE string VALUE 'usePrinterDefaults',
      lc_xml_attr_verticaldpi        TYPE string VALUE 'verticalDpi',
      lc_xml_attr_differentoddeven   TYPE string VALUE 'differentOddEven',
      lc_xml_attr_iconset            TYPE string VALUE 'iconSet',
      lc_xml_attr_showvalue          TYPE string VALUE 'showValue',
      lc_xml_attr_dxfid              TYPE string VALUE 'dxfId',
      lc_xml_attr_priority           TYPE string VALUE 'priority',
      lc_xml_attr_text               TYPE string VALUE 'text'.
    DATA:
      o_excel_ref    TYPE REF TO zcl_excel_writer_2007,
      o_worksheet    TYPE REF TO zcl_excel_worksheet,
      o_document     TYPE REF TO if_ixml_document,
      o_element_root TYPE REF TO if_ixml_element,
      v_active       TYPE flag,
      v_relation_id  TYPE i VALUE 0.
    METHODS:
      add_sheetpr,
      add_dimension                   RAISING zcx_excel,
      add_sheet_views                 RAISING zcx_excel,
      add_sheetformatpr               RAISING zcx_excel,
      add_cols                        RAISING zcx_excel,
      add_sheet_protection,
      add_autofilter                  RAISING zcx_excel,
      add_merge_cells                 RAISING zcx_excel,
      add_conditional_formatting      RAISING zcx_excel,
      add_data_validations,
      add_hyperlinks,
      add_print_options,
      add_page_margins,
      add_page_setup,
      add_header_footer,
      add_drawing,
      add_drawing_for_comments,
      add_drawing_for_header_footer,
      add_table_parts,
      add_sheet_data                  RAISING zcx_excel,
      add_page_breaks,
      add_ignored_errors.
ENDCLASS.

CLASS lcl_create_xl_sheet IMPLEMENTATION.
  METHOD create.
    o_excel_ref    = io_excel_writer_2007.
    o_worksheet    = io_worksheet.
    o_document     = io_document.
    v_active       = iv_active.
    o_element_root = o_document->get_root_element( ).

    add_sheetpr( ).
    " dimension node
    add_dimension( ).
    " sheetViews node
    add_sheet_views( ).
    add_sheetformatpr( ).
    add_cols( ).
*--------------------------------------------------------------------*
* Sheet content - use own method to create this
*--------------------------------------------------------------------*
    add_sheet_data( ).
    add_sheet_protection( ).
    add_autofilter( ).
    " Merged cells
    add_merge_cells(  ).
    " Conditional formatting node
    add_conditional_formatting( ).
    add_data_validations( ).
    " Hyperlinks
    add_hyperlinks( ).
    " PrintOptions
    add_print_options( ).
    " pageMargins node
    add_page_margins( ).
* pageSetup node
    add_page_setup( ).
* { headerFooter necessary?   >
    add_header_footer( ).
    add_page_breaks( ).
* drawing
    add_drawing( ).
    add_drawing_for_comments( ).
    add_drawing_for_header_footer( ).
* ignoredErrors
    add_ignored_errors( ).
    add_table_parts( ).
  ENDMETHOD.
  METHOD add_sheetpr.

    DATA:
      lo_element   TYPE REF TO if_ixml_element,
      lo_element_2 TYPE REF TO if_ixml_element,
      lv_value     TYPE string.

    " sheetPr
    lo_element = o_document->create_simple_element( name   = lc_xml_node_sheetpr
                                                    parent = o_document ).
    " TODO tabColor
    IF o_worksheet->tabcolor IS NOT INITIAL.
      lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_tabcolor
                                                        parent = lo_element ).
* Theme not supported yet - start with RGB
      lv_value = o_worksheet->tabcolor-rgb.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabcolor_rgb
                                      value = lv_value ).
    ENDIF.

    " outlinePr
    lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_outlinepr
                                                      parent = o_document ).

    lv_value = o_worksheet->zif_excel_sheet_properties~summarybelow.
    CONDENSE lv_value.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_summarybelow
                                    value = lv_value ).

    lv_value = o_worksheet->zif_excel_sheet_properties~summaryright.
    CONDENSE lv_value.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_summaryright
                                    value = lv_value ).

    lo_element->append_child( new_child = lo_element_2 ).

    IF o_worksheet->sheet_setup->fit_to_page IS NOT INITIAL.
      lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_pagesetuppr
                                                        parent = o_document ).
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_fittopage
                                      value = `1` ).
      lo_element->append_child( new_child = lo_element_2 ). " pageSetupPr node
    ENDIF.

    o_element_root->append_child( new_child = lo_element ).
  ENDMETHOD.
  METHOD add_dimension.
    DATA:
      lo_element TYPE REF TO if_ixml_element,
      lv_value   TYPE string.

    lo_element = o_document->create_simple_element( name   = lc_xml_node_dimension
                                                    parent = o_document ).
    lv_value = o_worksheet->get_dimension_range( ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_ref
                                  value = lv_value ).
    o_element_root->append_child( new_child = lo_element ).
  ENDMETHOD.
  METHOD add_sheet_views.
    DATA:
      lo_element   TYPE REF TO if_ixml_element,
      lo_element_2 TYPE REF TO if_ixml_element,
      lo_element_3 TYPE REF TO if_ixml_element.

    DATA: lv_value                    TYPE string,
          lv_freeze_cell_row          TYPE zexcel_cell_row,
          lv_freeze_cell_column       TYPE zexcel_cell_column,
          lv_freeze_cell_column_alpha TYPE zexcel_cell_column_alpha.


    lo_element = o_document->create_simple_element( name   = lc_xml_node_sheetviews
                                                    parent = o_document ).
    " sheetView node
    lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_sheetview
                                                      parent = o_document ).
    IF o_worksheet->zif_excel_sheet_properties~show_zeros EQ abap_false.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showzeros
                                      value = '0' ).
    ENDIF.
    IF   v_active = abap_true
      OR o_worksheet->zif_excel_sheet_properties~selected EQ abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabselected
                                      value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabselected
                                      value = '0' ).
    ENDIF.
    " Zoom scale
    IF o_worksheet->zif_excel_sheet_properties~zoomscale NE 0.
      IF o_worksheet->zif_excel_sheet_properties~zoomscale GT 400.
        o_worksheet->zif_excel_sheet_properties~zoomscale = 400.
      ELSEIF o_worksheet->zif_excel_sheet_properties~zoomscale LT 10.
        o_worksheet->zif_excel_sheet_properties~zoomscale = 10.
      ENDIF.
      lv_value = o_worksheet->zif_excel_sheet_properties~zoomscale.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscale
                                      value = lv_value ).
    ENDIF.
    IF o_worksheet->zif_excel_sheet_properties~zoomscale_normal NE 0.
      IF o_worksheet->zif_excel_sheet_properties~zoomscale_normal GT 400.
        o_worksheet->zif_excel_sheet_properties~zoomscale_normal = 400.
      ELSEIF o_worksheet->zif_excel_sheet_properties~zoomscale_normal LT 10.
        o_worksheet->zif_excel_sheet_properties~zoomscale_normal = 10.
      ENDIF.
      lv_value = o_worksheet->zif_excel_sheet_properties~zoomscale_normal.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalenormal
                                      value = lv_value ).
    ENDIF.
    IF o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview NE 0.
      IF o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview GT 400.
        o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 400.
      ELSEIF o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview LT 10.
        o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 10.
      ENDIF.
      lv_value = o_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalepageview
                                      value = lv_value ).
    ENDIF.
    IF o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview NE 0.
      IF o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview GT 400.
        o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 400.
      ELSEIF o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview LT 10.
        o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 10.
      ENDIF.
      lv_value = o_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalesheetview
                                      value = lv_value ).
    ENDIF.
    IF o_worksheet->zif_excel_sheet_properties~get_right_to_left( ) EQ abap_true.
      lo_element_2->set_attribute_ns( name  = 'rightToLeft'
                                      value = '1' ).
    ENDIF.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_workbookviewid
                                    value = '0' ).
    " showGridLines attribute
    IF o_worksheet->show_gridlines = abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showgridlines
                                      value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showgridlines
                                      value = '0' ).
    ENDIF.

    " showRowColHeaders attribute
    IF o_worksheet->show_rowcolheaders = abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showrowcolheaders
                                      value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showrowcolheaders
                                      value = '0' ).
    ENDIF.

    IF o_worksheet->sheetview_top_left_cell IS NOT INITIAL.
      lo_element_2->set_attribute_ns( name  = 'topLeftCell'
                                      value = o_worksheet->sheetview_top_left_cell ).
    ENDIF.

    " freeze panes
    o_worksheet->get_freeze_cell( IMPORTING ep_row    = lv_freeze_cell_row
                                            ep_column = lv_freeze_cell_column ).

    IF lv_freeze_cell_row IS NOT INITIAL AND lv_freeze_cell_column IS NOT INITIAL.
      lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_pane
                                                        parent = lo_element_2 ).

      IF lv_freeze_cell_row > 1.
        lv_value = lv_freeze_cell_row - 1.
        CONDENSE lv_value.
        lo_element_3->set_attribute_ns( name  = 'ySplit'
                                        value = lv_value ).
      ENDIF.

      IF lv_freeze_cell_column > 1.
        lv_value = lv_freeze_cell_column - 1.
        CONDENSE lv_value.
        lo_element_3->set_attribute_ns( name  = 'xSplit'
                                        value = lv_value ).
      ENDIF.

      IF o_worksheet->pane_top_left_cell IS NOT INITIAL.
        lo_element_3->set_attribute_ns( name  = 'topLeftCell'
                                        value = o_worksheet->pane_top_left_cell ).
      ELSE.
        lv_value = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_freeze_cell_column
                                                                     i_row    = lv_freeze_cell_row ).
        lo_element_3->set_attribute_ns( name  = 'topLeftCell'
                                        value = lv_value ).
      ENDIF.

      lo_element_3->set_attribute_ns( name  = 'activePane'
                                      value = 'bottomRight' ).

      lo_element_3->set_attribute_ns( name  = 'state'
                                      value = 'frozen' ).

      lo_element_2->append_child( new_child = lo_element_3 ).
    ENDIF.
    " selection node
    lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_selection
                                                      parent = o_document ).
    lv_value = o_worksheet->get_active_cell( ).
    lo_element_3->set_attribute_ns( name  = lc_xml_attr_activecell
                                    value = lv_value ).

    lo_element_3->set_attribute_ns( name  = lc_xml_attr_sqref
                                    value = lv_value ).

    lo_element_2->append_child( new_child = lo_element_3 ). " sheetView node

    lo_element->append_child( new_child = lo_element_2 ). " sheetView node

    o_element_root->append_child( new_child = lo_element ). " sheetViews node
  ENDMETHOD.

  METHOD add_sheetformatpr.
    DATA: lo_element         TYPE REF TO if_ixml_element,
          lo_column_default  TYPE REF TO zcl_excel_column,
          lo_row_default     TYPE REF TO zcl_excel_row,lv_value           TYPE string,
          lo_column_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_column          TYPE REF TO zcl_excel_column,
          lo_row_iterator    TYPE REF TO zcl_excel_collection_iterator,
          outline_level_col  TYPE i VALUE 0.

    lo_column_iterator = o_worksheet->get_columns_iterator( ).
    lo_row_iterator = o_worksheet->get_rows_iterator( ).
    " Calculate col
    IF NOT lo_column_iterator IS BOUND.
      o_worksheet->calculate_column_widths( ).
      lo_column_iterator = o_worksheet->get_columns_iterator( ).
    ENDIF.

    " sheetFormatPr node
    lo_element = o_document->create_simple_element( name   = lc_xml_node_sheetformatpr
                                                    parent = o_document ).
    " defaultRowHeight
    lo_row_default = o_worksheet->get_default_row( ).
    IF lo_row_default IS BOUND.
      IF lo_row_default->get_row_height( ) >= 0.
        lo_element->set_attribute_ns( name  = lc_xml_attr_customheight
                                      value = lc_xml_attr_true ).
        lv_value = lo_row_default->get_row_height( ).
      ELSE.
        lv_value = '12.75'.
      ENDIF.
    ELSE.
      lv_value = '12.75'.
    ENDIF.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaultrowheight
                                  value = lv_value ).
    " defaultColWidth
    lo_column_default = o_worksheet->get_default_column( ).
    IF lo_column_default IS BOUND AND lo_column_default->get_width( ) >= 0.
      lv_value = lo_column_default->get_width( ).
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_defaultcolwidth
                                    value = lv_value ).
    ENDIF.

    " outlineLevelCol
    WHILE lo_column_iterator->has_next( ) = abap_true.
      lo_column ?= lo_column_iterator->get_next( ).
      IF lo_column->get_outline_level( ) > outline_level_col.
        outline_level_col = lo_column->get_outline_level( ).
      ENDIF.
    ENDWHILE.

    lv_value = outline_level_col.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element->set_attribute_ns( name  = lc_xml_attr_outlinelevelcol
                                  value = lv_value ).

    o_element_root->append_child( new_child = lo_element ). " sheetFormatPr node
  ENDMETHOD.

  METHOD add_cols.
    DATA: lo_element         TYPE REF TO if_ixml_element,
          lo_element_2       TYPE REF TO if_ixml_element,
          lo_column_default  TYPE REF TO zcl_excel_column,
          lv_value           TYPE string,
          lv_column          TYPE zexcel_cell_column,
          lv_style_guid      TYPE zexcel_cell_style,
          ls_style_mapping   TYPE zexcel_s_styles_mapping,
          lo_column_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_column          TYPE REF TO zcl_excel_column.

* Reset column iterator
    lo_column_iterator = o_worksheet->get_columns_iterator( ).
    lo_column_default = o_worksheet->get_default_column( ).
    IF o_worksheet->zif_excel_sheet_properties~get_style( ) IS NOT INITIAL OR lo_column_iterator->has_next( ) = abap_true.
      " cols node
      lo_element = o_document->create_simple_element( name   = lc_xml_node_cols
                                                      parent = o_document ).
      " This code have to be enhanced in order to manage also column style properties
      " Now it is an out/out
      IF lo_column_iterator->has_next( ) = abap_true.
        WHILE lo_column_iterator->has_next( ) = abap_true.
          lo_column ?= lo_column_iterator->get_next( ).
          " col node
          lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_col
                                                            parent = o_document ).
          lv_value = lo_column->get_column_index( ).
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_min
                                          value = lv_value ).
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_max
                                          value = lv_value ).
          " Width
          IF lo_column->get_width( ) < 0.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_width
                                            value = lc_xml_attr_defaultwidth ).
          ELSE.
            lv_value = lo_column->get_width( ).
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_width
                                            value = lv_value ).
          ENDIF.
          " Column visibility
          IF lo_column->get_visible( ) = abap_false.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                            value = lc_xml_attr_true ).
          ENDIF.
          "  Auto size?
          IF lo_column->get_auto_size( ) = abap_true.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_bestfit
                                            value = lc_xml_attr_true ).
          ENDIF.
          " Custom width?
          IF lo_column_default IS BOUND.
            IF lo_column->get_width( ) <> lo_column_default->get_width( ).
              lo_element_2->set_attribute_ns( name  = lc_xml_attr_customwidth
                                              value = lc_xml_attr_true ).

            ENDIF.
          ELSE.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_customwidth
                                            value = lc_xml_attr_true ).
          ENDIF.
          " Collapsed
          IF lo_column->get_collapsed( ) = abap_true.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_collapsed
                                            value = lc_xml_attr_true ).
          ENDIF.
          " outlineLevel
          IF lo_column->get_outline_level( ) > 0.
            lv_value = lo_column->get_outline_level( ).

            SHIFT lv_value RIGHT DELETING TRAILING space.
            SHIFT lv_value LEFT DELETING LEADING space.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_outlinelevel
                                            value = lv_value ).
          ENDIF.
          " Style
          lv_style_guid = lo_column->get_column_style_guid( ).                      "ins issue #157 -  set column style
          CLEAR ls_style_mapping.
          READ TABLE o_excel_ref->styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
          IF sy-subrc = 0.                                                                                     "ins issue #295
            lv_value = ls_style_mapping-style.                                                                 "ins issue #295
            SHIFT lv_value RIGHT DELETING TRAILING space.
            SHIFT lv_value LEFT DELETING LEADING space.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_style
                                            value = lv_value ).
          ENDIF.                                                                                              "ins issue #237

          lo_element->append_child( new_child = lo_element_2 ). " col node
        ENDWHILE.

      ENDIF.
* Always pass through this coding
      IF o_worksheet->zif_excel_sheet_properties~get_style( ) IS NOT INITIAL.
        DATA: lts_sorted_columns TYPE SORTED TABLE OF zexcel_cell_column WITH UNIQUE KEY table_line.
        TYPES: BEGIN OF ty_missing_columns,
                 first_column TYPE zexcel_cell_column,
                 last_column  TYPE zexcel_cell_column,
               END OF ty_missing_columns.
        DATA: t_missing_columns TYPE STANDARD TABLE OF ty_missing_columns WITH NON-UNIQUE DEFAULT KEY,
              missing_column    LIKE LINE OF t_missing_columns.

* First collect columns that were already handled before.  The rest has to be inserted now
        lo_column_iterator = o_worksheet->get_columns_iterator( ).
        WHILE lo_column_iterator->has_next( ) = abap_true.
          lo_column ?= lo_column_iterator->get_next( ).
          lv_column = zcl_excel_common=>convert_column2int( lo_column->get_column_index( ) ).
          INSERT lv_column INTO TABLE lts_sorted_columns.
        ENDWHILE.

* Now find all columns that were missing so far
        missing_column-first_column = 1.
        LOOP AT lts_sorted_columns INTO lv_column.
          IF lv_column > missing_column-first_column.
            missing_column-last_column = lv_column - 1.
            APPEND missing_column TO t_missing_columns.
          ENDIF.
          missing_column-first_column = lv_column + 1.
        ENDLOOP.
        missing_column-last_column = zcl_excel_common=>c_excel_sheet_max_col.
        APPEND missing_column TO t_missing_columns.
* Now apply stylesetting ( and other defaults - I copy it from above.  Whoever programmed that seems to know what to do  :o)
        LOOP AT t_missing_columns INTO missing_column.
* End of insertion issue #157 -  set column style
          lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_col
                                                            parent = o_document ).
          lv_value = missing_column-first_column.             "ins issue #157  -  set sheet style ( add missing columns
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_min
                                          value = lv_value ).
*        lv_value = zcl_excel_common=>c_excel_sheet_max_col."del issue #157  -  set sheet style ( add missing columns
          lv_value = missing_column-last_column.              "ins issue #157  -  set sheet style ( add missing columns
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_max
                                          value = lv_value ).
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_width
                                          value = lc_xml_attr_defaultwidth ).
          lv_style_guid = o_worksheet->zif_excel_sheet_properties~get_style( ).
          READ TABLE o_excel_ref->styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
          lv_value = ls_style_mapping-style.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
          lo_element->append_child( new_child = lo_element_2 ). " col node
        ENDLOOP.      "ins issue #157  -  set sheet style ( add missing columns

      ENDIF.
*--------------------------------------------------------------------*
* issue #367 add feature hide columns from
*--------------------------------------------------------------------*
      IF o_worksheet->zif_excel_sheet_properties~hide_columns_from IS NOT INITIAL.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_col
                                                          parent = o_document ).
        lv_value = zcl_excel_common=>convert_column2int( o_worksheet->zif_excel_sheet_properties~hide_columns_from ).
        CONDENSE lv_value NO-GAPS.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_min
                                        value = lv_value ).
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_max
                                        value = '16384' ).
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                        value = '1' ).
        lo_element->append_child( new_child = lo_element_2 ). " col node
      ENDIF.

      o_element_root->append_child( new_child = lo_element ). " cols node
    ENDIF.
  ENDMETHOD.

  METHOD add_sheet_protection.
    DATA:
      lo_element TYPE REF TO if_ixml_element,
      lv_value   TYPE string.

*< Begin of insertion Issue #572 - Protect sheet with filter caused Excel error
* Autofilter must be set AFTER sheet protection in XML
    IF o_worksheet->zif_excel_sheet_protection~protected EQ abap_true.
      " sheetProtection node
      lo_element = o_document->create_simple_element( name   = lc_xml_node_sheetprotection
                                                      parent = o_document ).
      lv_value = o_worksheet->zif_excel_sheet_protection~password.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_password
                                      value = lv_value ).
      ENDIF.
      lv_value = o_worksheet->zif_excel_sheet_protection~auto_filter.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_autofilter
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~delete_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_deletecolumns
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~delete_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_deleterows
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~format_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatcells
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~format_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatcolumns
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~format_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatrows
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~insert_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_insertcolumns
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~insert_hyperlinks.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_inserthyperlinks
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~insert_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_insertrows
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~objects.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_objects
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~pivot_tables.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_pivottables
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~scenarios.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_scenarios
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~select_locked_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_selectlockedcells
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~select_unlocked_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_selectunlockedcell
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~sheet.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_sheet
                                    value = lv_value ).
      lv_value = o_worksheet->zif_excel_sheet_protection~sort.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_sort
                                    value = lv_value ).

      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
*> End of insertion Issue #572 - Protect sheet with filter caused Excel error
  ENDMETHOD.

  METHOD add_autofilter.
    DATA:
      lo_element     TYPE REF TO if_ixml_element,
      lo_element_2   TYPE REF TO if_ixml_element,
      lo_element_3   TYPE REF TO if_ixml_element,
      lo_element_4   TYPE REF TO if_ixml_element,
      lv_value       TYPE string,
      lv_column      TYPE zexcel_cell_column,
      lt_values      TYPE zexcel_t_autofilter_values,
      ls_values      TYPE zexcel_s_autofilter_values,
      lo_autofilters TYPE REF TO zcl_excel_autofilters,
      lo_autofilter  TYPE REF TO zcl_excel_autofilter,
      lv_ref         TYPE string.

    lo_autofilters = o_excel_ref->excel->get_autofilters_reference( ).
    lo_autofilter = lo_autofilters->get( io_worksheet = o_worksheet ) .

    IF lo_autofilter IS BOUND.
* Create node autofilter
      lo_element = o_document->create_simple_element( name   = lc_xml_node_autofilter
                                                      parent = o_document ).
      lv_ref = lo_autofilter->get_filter_range( ) .
      CONDENSE lv_ref NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_ref
                                    value = lv_ref ).
      lt_values = lo_autofilter->get_values( ) .
      IF lt_values IS NOT INITIAL.
* If we filter we need to set the filter mode to 1.
        lo_element_2 = o_document->find_from_name( name = lc_xml_node_sheetpr ).
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_filtermode
                                        value = '1' ).
* Create node filtercolumn
        CLEAR lv_column.
        LOOP AT lt_values INTO ls_values.
          IF ls_values-column <> lv_column.
            IF lv_column IS NOT INITIAL.
              lo_element_2->append_child( new_child = lo_element_3 ).
              lo_element->append_child( new_child = lo_element_2 ).
            ENDIF.
            lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_filtercolumn
                                                              parent = lo_element ).
            lv_column   = ls_values-column - lo_autofilter->filter_area-col_start.
            lv_value = lv_column.
            CONDENSE lv_value NO-GAPS.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_colid
                                            value = lv_value ).
            lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_filters
                                                              parent = lo_element_2 ).
            lv_column = ls_values-column.
          ENDIF.
          lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_filter
                                                            parent = lo_element_3 ).
          lo_element_4->set_attribute_ns( name  = lc_xml_attr_val
                                          value = ls_values-value ).
          lo_element_3->append_child( new_child = lo_element_4 ). " value node
        ENDLOOP.
        lo_element_2->append_child( new_child = lo_element_3 ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.
      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
  ENDMETHOD.

  METHOD add_merge_cells.
    DATA:
      lo_element     TYPE REF TO if_ixml_element,
      lo_element_2   TYPE REF TO if_ixml_element,
      lv_value       TYPE string,
      lt_range_merge TYPE string_table,
      merge_count    TYPE int4.

    FIELD-SYMBOLS: <fs_range_merge>         LIKE LINE OF lt_range_merge.

    lt_range_merge = o_worksheet->get_merge( ).
    IF lt_range_merge IS NOT INITIAL.
      lo_element = o_document->create_simple_element( name   = lc_xml_node_mergecells
                                                      parent = o_document ).
      DESCRIBE TABLE lt_range_merge LINES merge_count.
      lv_value = merge_count.
      CONDENSE lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                    value = lv_value ).
      LOOP AT lt_range_merge ASSIGNING <fs_range_merge>.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_mergecell
                                                          parent = o_document ).

        lo_element_2->set_attribute_ns( name  = lc_xml_attr_ref
                                        value = <fs_range_merge> ).
        lo_element->append_child( new_child = lo_element_2 ).
        o_element_root->append_child( new_child = lo_element ).
        o_worksheet->delete_merge( ).
      ENDLOOP.
    ENDIF.
  ENDMETHOD.

  METHOD add_conditional_formatting.
    DATA: lo_element               TYPE REF TO if_ixml_element,
          lo_element_2             TYPE REF TO if_ixml_element,
          lo_element_3             TYPE REF TO if_ixml_element,
          lo_element_4             TYPE REF TO if_ixml_element,
          lo_iterator              TYPE REF TO zcl_excel_collection_iterator,
          lo_style_cond            TYPE REF TO zcl_excel_style_cond,
          lv_value                 TYPE string,
          ls_databar               TYPE zexcel_conditional_databar,      " Databar by Albert Lladanosa
          ls_colorscale            TYPE zexcel_conditional_colorscale,
          ls_iconset               TYPE zexcel_conditional_iconset,
          ls_cellis                TYPE zexcel_conditional_cellis,
          ls_textfunction          TYPE zcl_excel_style_cond=>ts_conditional_textfunction,
          lv_column_start          TYPE zexcel_cell_column_alpha,
          lv_row_start             TYPE zexcel_cell_row,
          lv_cell_coords           TYPE zexcel_cell_coords,
          ls_expression            TYPE zexcel_conditional_expression,
          ls_conditional_top10     TYPE zexcel_conditional_top10,
          ls_conditional_above_avg TYPE zexcel_conditional_above_avg,
          lt_cfvo                  TYPE TABLE OF cfvo,
          ls_cfvo                  TYPE cfvo,
          lt_colors                TYPE TABLE OF colors,
          ls_colors                TYPE colors,
          ls_style_cond_mapping    TYPE zexcel_s_styles_cond_mapping,
          lt_condformating_ranges  TYPE ty_condformating_ranges,
          ls_condformating_range   TYPE ty_condformating_range.

    FIELD-SYMBOLS: <ls_condformating_range> TYPE ty_condformating_range.

    lo_iterator = o_worksheet->get_style_cond_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_style_cond ?= lo_iterator->get_next( ).
      IF lo_style_cond->rule IS INITIAL.
        CONTINUE.
      ENDIF.

      lv_value = lo_style_cond->get_dimension_range( ).

      READ TABLE lt_condformating_ranges WITH KEY dimension_range = lv_value ASSIGNING <ls_condformating_range>.
      IF sy-subrc = 0.
        lo_element = <ls_condformating_range>-condformatting_node.
      ELSE.
        lo_element = o_document->create_simple_element( name   = lc_xml_node_condformatting
                                                        parent = o_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_sqref
                                      value = lv_value ).

        ls_condformating_range-dimension_range = lv_value.
        ls_condformating_range-condformatting_node = lo_element.
        INSERT ls_condformating_range INTO TABLE lt_condformating_ranges.

      ENDIF.

      " cfRule node
      lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_cfrule
                                                        parent = o_document ).
      IF lo_style_cond->rule = zcl_excel_style_cond=>c_rule_textfunction.
        IF lo_style_cond->mode_textfunction-textfunction = zcl_excel_style_cond=>c_textfunction_notcontains.
          lv_value = `notContainsText`.
        ELSE.
          lv_value = lo_style_cond->mode_textfunction-textfunction.
        ENDIF.
      ELSE.
        lv_value = lo_style_cond->rule.
      ENDIF.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_type
                                      value = lv_value ).
      lv_value = lo_style_cond->priority.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_priority
                                      value = lv_value ).

      CASE lo_style_cond->rule.
          " Start >> Databar by Albert Lladanosa
        WHEN zcl_excel_style_cond=>c_rule_databar.

          ls_databar = lo_style_cond->mode_databar.

          CLEAR lt_cfvo.
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_databar
                                                            parent = o_document ).

          ls_cfvo-value = ls_databar-cfvo1_value.
          ls_cfvo-type = ls_databar-cfvo1_type.
          APPEND ls_cfvo TO lt_cfvo.

          ls_cfvo-value = ls_databar-cfvo2_value.
          ls_cfvo-type = ls_databar-cfvo2_type.
          APPEND ls_cfvo TO lt_cfvo.

          LOOP AT lt_cfvo INTO ls_cfvo.
            " cfvo node
            lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_cfvo
                                                              parent = o_document ).
            lv_value = ls_cfvo-type.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_type
                                            value = lv_value ).
            lv_value = ls_cfvo-value.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
            lo_element_3->append_child( new_child = lo_element_4 ). " cfvo node
          ENDLOOP.

          lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_color
                                                            parent = o_document ).
          lv_value = ls_databar-colorrgb.
          lo_element_4->set_attribute_ns( name  = lc_xml_attr_tabcolor_rgb
                                          value = lv_value ).

          lo_element_3->append_child( new_child = lo_element_4 ). " color node

          lo_element_2->append_child( new_child = lo_element_3 ). " databar node
          " End << Databar by Albert Lladanosa

        WHEN zcl_excel_style_cond=>c_rule_colorscale.

          ls_colorscale = lo_style_cond->mode_colorscale.

          CLEAR: lt_cfvo, lt_colors.
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_colorscale
                                                            parent = o_document ).

          ls_cfvo-value = ls_colorscale-cfvo1_value.
          ls_cfvo-type = ls_colorscale-cfvo1_type.
          APPEND ls_cfvo TO lt_cfvo.

          ls_cfvo-value = ls_colorscale-cfvo2_value.
          ls_cfvo-type = ls_colorscale-cfvo2_type.
          APPEND ls_cfvo TO lt_cfvo.

          ls_cfvo-value = ls_colorscale-cfvo3_value.
          ls_cfvo-type = ls_colorscale-cfvo3_type.
          APPEND ls_cfvo TO lt_cfvo.

          APPEND ls_colorscale-colorrgb1 TO lt_colors.
          APPEND ls_colorscale-colorrgb2 TO lt_colors.
          APPEND ls_colorscale-colorrgb3 TO lt_colors.

          LOOP AT lt_cfvo INTO ls_cfvo.

            IF ls_cfvo IS INITIAL.
              CONTINUE.
            ENDIF.

            " cfvo node
            lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_cfvo
                                                              parent = o_document ).
            lv_value = ls_cfvo-type.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_type
                                            value = lv_value ).
            lv_value = ls_cfvo-value.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
            lo_element_3->append_child( new_child = lo_element_4 ). " cfvo node
          ENDLOOP.
          LOOP AT lt_colors INTO ls_colors.

            IF ls_colors IS INITIAL.
              CONTINUE.
            ENDIF.

            lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_color
                                                              parent = o_document ).
            lv_value = ls_colors-colorrgb.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_tabcolor_rgb
                                            value = lv_value ).

            lo_element_3->append_child( new_child = lo_element_4 ). " color node
          ENDLOOP.

          lo_element_2->append_child( new_child = lo_element_3 ). " databar node

        WHEN zcl_excel_style_cond=>c_rule_iconset.

          ls_iconset = lo_style_cond->mode_iconset.

          CLEAR lt_cfvo.
          " iconset node
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_iconset
                                                            parent = o_document ).
          IF ls_iconset-iconset NE zcl_excel_style_cond=>c_iconset_3trafficlights.
            lv_value = ls_iconset-iconset.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_iconset
                                            value = lv_value ).
          ENDIF.

          " Set the showValue attribute
          lv_value = ls_iconset-showvalue.
          lo_element_3->set_attribute_ns( name  = lc_xml_attr_showvalue
                                          value = lv_value ).

          CASE ls_iconset-iconset.
            WHEN zcl_excel_style_cond=>c_iconset_3trafficlights2 OR
                 zcl_excel_style_cond=>c_iconset_3arrows OR
                 zcl_excel_style_cond=>c_iconset_3arrowsgray OR
                 zcl_excel_style_cond=>c_iconset_3flags OR
                 zcl_excel_style_cond=>c_iconset_3signs OR
                 zcl_excel_style_cond=>c_iconset_3symbols OR
                 zcl_excel_style_cond=>c_iconset_3symbols2 OR
                 zcl_excel_style_cond=>c_iconset_3trafficlights OR
                 zcl_excel_style_cond=>c_iconset_3trafficlights2.
              ls_cfvo-value = ls_iconset-cfvo1_value.
              ls_cfvo-type = ls_iconset-cfvo1_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo2_value.
              ls_cfvo-type = ls_iconset-cfvo2_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo3_value.
              ls_cfvo-type = ls_iconset-cfvo3_type.
              APPEND ls_cfvo TO lt_cfvo.
            WHEN zcl_excel_style_cond=>c_iconset_4arrows OR
                 zcl_excel_style_cond=>c_iconset_4arrowsgray OR
                 zcl_excel_style_cond=>c_iconset_4rating OR
                 zcl_excel_style_cond=>c_iconset_4redtoblack OR
                 zcl_excel_style_cond=>c_iconset_4trafficlights.
              ls_cfvo-value = ls_iconset-cfvo1_value.
              ls_cfvo-type = ls_iconset-cfvo1_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo2_value.
              ls_cfvo-type = ls_iconset-cfvo2_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo3_value.
              ls_cfvo-type = ls_iconset-cfvo3_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo4_value.
              ls_cfvo-type = ls_iconset-cfvo4_type.
              APPEND ls_cfvo TO lt_cfvo.
            WHEN zcl_excel_style_cond=>c_iconset_5arrows OR
                 zcl_excel_style_cond=>c_iconset_5arrowsgray OR
                 zcl_excel_style_cond=>c_iconset_5quarters OR
                 zcl_excel_style_cond=>c_iconset_5rating.
              ls_cfvo-value = ls_iconset-cfvo1_value.
              ls_cfvo-type = ls_iconset-cfvo1_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo2_value.
              ls_cfvo-type = ls_iconset-cfvo2_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo3_value.
              ls_cfvo-type = ls_iconset-cfvo3_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo4_value.
              ls_cfvo-type = ls_iconset-cfvo4_type.
              APPEND ls_cfvo TO lt_cfvo.
              ls_cfvo-value = ls_iconset-cfvo5_value.
              ls_cfvo-type = ls_iconset-cfvo5_type.
              APPEND ls_cfvo TO lt_cfvo.
            WHEN OTHERS.
              CLEAR lt_cfvo.
          ENDCASE.

          LOOP AT lt_cfvo INTO ls_cfvo.
            " cfvo node
            lo_element_4 = o_document->create_simple_element( name   = lc_xml_node_cfvo
                                                              parent = o_document ).
            lv_value = ls_cfvo-type.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_type
                                            value = lv_value ).
            lv_value = ls_cfvo-value.
            lo_element_4->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
            lo_element_3->append_child( new_child = lo_element_4 ). " cfvo node
          ENDLOOP.


          lo_element_2->append_child( new_child = lo_element_3 ). " iconset node

        WHEN zcl_excel_style_cond=>c_rule_cellis.
          ls_cellis = lo_style_cond->mode_cellis.
          READ TABLE o_excel_ref->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_cellis-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          lv_value = ls_cellis-operator.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_operator
                                          value = lv_value ).
          " formula node
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula
                                                            parent = o_document ).
          lv_value = ls_cellis-formula.
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node
          IF ls_cellis-formula2 IS NOT INITIAL.
            lv_value = ls_cellis-formula2.
            lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula
                                                              parent = o_document ).
            lo_element_3->set_value( value = lv_value ).
            lo_element_2->append_child( new_child = lo_element_3 ). " 2nd formula node
          ENDIF.
*--------------------------------------------------------------------------------------*
* The below code creates an EXM structure in the following format:
* -<conditionalFormatting sqref="G6:G12">-
*   <cfRule operator="beginsWith" priority="4" dxfId="4" type="beginsWith" text="1">
*     <formula>LEFT(G6,LEN("1"))="1"</formula>
*   </cfRule>
* </conditionalFormatting>
*--------------------------------------------------------------------------------------*
        WHEN zcl_excel_style_cond=>c_rule_textfunction.
          ls_textfunction = lo_style_cond->mode_textfunction.
          READ TABLE o_excel_ref->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_cellis-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          lv_value = ls_textfunction-textfunction.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_operator
                                          value = lv_value ).

          " text
          lv_value = ls_textfunction-text.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_text
                                          value = lv_value ).

          " formula node
          zcl_excel_common=>convert_range2column_a_row(
            EXPORTING
              i_range        = lo_style_cond->get_dimension_range( )
            IMPORTING
              e_column_start = lv_column_start
              e_row_start    = lv_row_start ).
          lv_cell_coords = |{ lv_column_start }{ lv_row_start }|.
          CASE ls_textfunction-textfunction.
            WHEN zcl_excel_style_cond=>c_textfunction_beginswith.
              lv_value = |LEFT({ lv_cell_coords },LEN("{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }"))=|
                      && |"{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }"|.
            WHEN zcl_excel_style_cond=>c_textfunction_containstext.
              lv_value = |NOT(ISERROR(SEARCH("{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }",{ lv_cell_coords })))|.
            WHEN zcl_excel_style_cond=>c_textfunction_endswith.
              lv_value = |RIGHT({ lv_cell_coords },LEN("{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }"))=|
                      && |"{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }"|.
            WHEN zcl_excel_style_cond=>c_textfunction_notcontains.
              lv_value = |ISERROR(SEARCH("{ escape( val = ls_textfunction-text format = cl_abap_format=>e_html_text ) }",{ lv_cell_coords }))|.
            WHEN OTHERS.
          ENDCASE.
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula
                                                            parent = o_document ).
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node

        WHEN zcl_excel_style_cond=>c_rule_expression.
          ls_expression = lo_style_cond->mode_expression.
          READ TABLE o_excel_ref->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_expression-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          " formula node
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula
                                                            parent = o_document ).
          lv_value = ls_expression-formula.
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node

* begin of ins issue #366 - missing conditional rules: top10
        WHEN zcl_excel_style_cond=>c_rule_top10.
          ls_conditional_top10 = lo_style_cond->mode_top10.
          READ TABLE o_excel_ref->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_conditional_top10-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          lv_value = ls_conditional_top10-topxx_count.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = 'rank'
                                          value = lv_value ).
          IF ls_conditional_top10-bottom = 'X'.
            lo_element_2->set_attribute_ns( name  = 'bottom'
                                            value = '1' ).
          ENDIF.
          IF ls_conditional_top10-percent = 'X'.
            lo_element_2->set_attribute_ns( name  = 'percent'
                                            value = '1' ).
          ENDIF.

        WHEN zcl_excel_style_cond=>c_rule_above_average.
          ls_conditional_above_avg = lo_style_cond->mode_above_average.
          READ TABLE o_excel_ref->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_conditional_above_avg-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).

          IF ls_conditional_above_avg-above_average IS INITIAL. " = below average
            lo_element_2->set_attribute_ns( name  = 'aboveAverage'
                                            value = '0' ).
          ENDIF.
          IF ls_conditional_above_avg-equal_average = 'X'. " = equal average also
            lo_element_2->set_attribute_ns( name  = 'equalAverage'
                                            value = '1' ).
          ENDIF.
          IF ls_conditional_above_avg-standard_deviation <> 0. " standard deviation instead of value
            lv_value = ls_conditional_above_avg-standard_deviation.
            lo_element_2->set_attribute_ns( name  = 'stdDev'
                                            value = lv_value ).
          ENDIF.

* end of ins issue #366 - missing conditional rules: top10

      ENDCASE.

      lo_element->append_child( new_child = lo_element_2 ). " cfRule node

      o_element_root->append_child( new_child = lo_element ). " Conditional formatting node
    ENDWHILE.
  ENDMETHOD.

  METHOD add_data_validations.
    DATA: lo_element         TYPE REF TO if_ixml_element,
          lo_element_2       TYPE REF TO if_ixml_element,
          lo_element_3       TYPE REF TO if_ixml_element,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_data_validation TYPE REF TO zcl_excel_data_validation,
          lv_value           TYPE string,
          lv_cell_row_s      TYPE string.

    IF o_worksheet->get_data_validations_size( ) GT 0.
      " dataValidations node
      lo_element = o_document->create_simple_element( name   = lc_xml_node_datavalidations
                                                      parent = o_document ).
      " Conditional formatting node
      lo_iterator = o_worksheet->get_data_validations_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_data_validation ?= lo_iterator->get_next( ).
        " dataValidation node
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_datavalidation
                                                          parent = o_document ).
        lv_value = lo_data_validation->type.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lv_value ).
        IF NOT lo_data_validation->operator IS INITIAL.
          lv_value = lo_data_validation->operator.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_operator
                                          value = lv_value ).
        ENDIF.
        IF lo_data_validation->allowblank EQ abap_true.
          lv_value = '1'.
        ELSE.
          lv_value = '0'.
        ENDIF.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_allowblank
                                        value = lv_value ).
        IF lo_data_validation->showinputmessage EQ abap_true.
          lv_value = '1'.
        ELSE.
          lv_value = '0'.
        ENDIF.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_showinputmessage
                                        value = lv_value ).
        IF lo_data_validation->showerrormessage EQ abap_true.
          lv_value = '1'.
        ELSE.
          lv_value = '0'.
        ENDIF.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_showerrormessage
                                        value = lv_value ).
        IF lo_data_validation->showdropdown EQ abap_true.
          lv_value = '1'.
        ELSE.
          lv_value = '0'.
        ENDIF.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_showdropdown
                                        value = lv_value ).
        IF NOT lo_data_validation->errortitle IS INITIAL.
          lv_value = lo_data_validation->errortitle.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_errortitle
                                          value = lv_value ).
        ENDIF.
        IF NOT lo_data_validation->error IS INITIAL.
          lv_value = lo_data_validation->error.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_error
                                          value = lv_value ).
        ENDIF.
        IF NOT lo_data_validation->errorstyle IS INITIAL.
          lv_value = lo_data_validation->errorstyle.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_errorstyle
                                          value = lv_value ).
        ENDIF.
        IF NOT lo_data_validation->prompttitle IS INITIAL.
          lv_value = lo_data_validation->prompttitle.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_prompttitle
                                          value = lv_value ).
        ENDIF.
        IF NOT lo_data_validation->prompt IS INITIAL.
          lv_value = lo_data_validation->prompt.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_prompt
                                          value = lv_value ).
        ENDIF.
        lv_cell_row_s = lo_data_validation->cell_row.
        CONDENSE lv_cell_row_s.
        CONCATENATE lo_data_validation->cell_column lv_cell_row_s INTO lv_value.
        IF lo_data_validation->cell_row_to IS NOT INITIAL.
          lv_cell_row_s = lo_data_validation->cell_row_to.
          CONDENSE lv_cell_row_s.
          CONCATENATE lv_value ':' lo_data_validation->cell_column_to lv_cell_row_s INTO lv_value.
        ENDIF.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_sqref
                                        value = lv_value ).
        " formula1 node
        lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula1
                                                          parent = o_document ).
        lv_value = lo_data_validation->formula1.
        lo_element_3->set_value( value = lv_value ).

        lo_element_2->append_child( new_child = lo_element_3 ). " formula1 node
        " formula2 node
        IF NOT lo_data_validation->formula2 IS INITIAL.
          lo_element_3 = o_document->create_simple_element( name   = lc_xml_node_formula2
                                                            parent = o_document ).
          lv_value = lo_data_validation->formula2.
          lo_element_3->set_value( value = lv_value ).

          lo_element_2->append_child( new_child = lo_element_3 ). " formula2 node
        ENDIF.

        lo_element->append_child( new_child = lo_element_2 ). " dataValidation node
      ENDWHILE.
      o_element_root->append_child( new_child = lo_element ). " dataValidations node
    ENDIF.
  ENDMETHOD.

  METHOD add_hyperlinks.
    DATA: lo_element          TYPE REF TO if_ixml_element,
          lo_element_2        TYPE REF TO if_ixml_element,
          lo_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lv_value            TYPE string,
          lv_hyperlinks_count TYPE i,
          lo_link             TYPE REF TO zcl_excel_hyperlink.

    lv_hyperlinks_count = o_worksheet->get_hyperlinks_size( ).
    IF lv_hyperlinks_count > 0.
      lo_element = o_document->create_simple_element( name   = 'hyperlinks'
                                                      parent = o_document ).

      lo_iterator = o_worksheet->get_hyperlinks_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_link ?= lo_iterator->get_next( ).

        lo_element_2 = o_document->create_simple_element( name   = 'hyperlink'
                                                          parent = lo_element ).

        lv_value = lo_link->get_ref( ).
        lo_element_2->set_attribute_ns( name  = 'ref'
                                        value = lv_value ).

        IF lo_link->is_internal( ) = abap_true.
          lv_value = lo_link->get_url( ).
          lo_element_2->set_attribute_ns( name  = 'location'
                                          value = lv_value ).
        ELSE.
          ADD 1 TO v_relation_id.

          lv_value = v_relation_id.
          CONDENSE lv_value.
          CONCATENATE 'rId' lv_value INTO lv_value.

          lo_element_2->set_attribute_ns( name  = 'r:id'
                                          value = lv_value ).

        ENDIF.

        lo_element->append_child( new_child = lo_element_2 ).
      ENDWHILE.

      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
  ENDMETHOD.

  METHOD add_print_options.
    DATA:
          lo_element      TYPE REF TO if_ixml_element.

    IF o_worksheet->print_gridlines                  = abap_true
      OR o_worksheet->sheet_setup->vertical_centered   = abap_true
      OR o_worksheet->sheet_setup->horizontal_centered = abap_true.
      lo_element = o_document->create_simple_element( name   = 'printOptions'
                                                      parent = o_document ).

      IF o_worksheet->print_gridlines = abap_true.
        lo_element->set_attribute_ns( name  = lc_xml_attr_gridlines
                                      value = 'true' ).
      ENDIF.

      IF o_worksheet->sheet_setup->horizontal_centered = abap_true.
        lo_element->set_attribute_ns( name  = 'horizontalCentered'
                                      value = 'true' ).
      ENDIF.

      IF o_worksheet->sheet_setup->vertical_centered = abap_true.
        lo_element->set_attribute_ns( name  = 'verticalCentered'
                                      value = 'true' ).
      ENDIF.

      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
  ENDMETHOD.

  METHOD add_page_margins.
    DATA:
      lo_element TYPE REF TO if_ixml_element,
      lv_value   TYPE string.

    lo_element = o_document->create_simple_element( name   = lc_xml_node_pagemargins
                                                    parent = o_document ).

    lv_value = o_worksheet->sheet_setup->margin_left.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_left
                                  value = lv_value ).
    lv_value = o_worksheet->sheet_setup->margin_right.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_right
                                  value = lv_value ).
    lv_value = o_worksheet->sheet_setup->margin_top.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_top
                                  value = lv_value ).
    lv_value = o_worksheet->sheet_setup->margin_bottom.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_bottom
                                  value = lv_value ).
    lv_value = o_worksheet->sheet_setup->margin_header.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_header
                                  value = lv_value ).
    lv_value = o_worksheet->sheet_setup->margin_footer.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_footer
                                  value = lv_value ).
    o_element_root->append_child( new_child = lo_element ). " pageMargins node
  ENDMETHOD.

  METHOD add_page_setup.
    DATA:
      lo_element TYPE REF TO if_ixml_element,
      lv_value   TYPE string.

    lo_element = o_document->create_simple_element( name   = lc_xml_node_pagesetup
                                                    parent = o_document ).

    IF o_worksheet->sheet_setup->black_and_white IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_blackandwhite
                                    value = `1` ).
    ENDIF.

    IF o_worksheet->sheet_setup->cell_comments IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_cellcomments
                                    value = o_worksheet->sheet_setup->cell_comments ).
    ENDIF.

    IF o_worksheet->sheet_setup->copies IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->copies.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_copies
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->draft IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_draft
                                    value = `1` ).
    ENDIF.

    IF o_worksheet->sheet_setup->errors IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_errors
                                    value = o_worksheet->sheet_setup->errors ).
    ENDIF.

    IF o_worksheet->sheet_setup->first_page_number IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->first_page_number.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_firstpagenumber
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->fit_to_page IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->fit_to_height.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fittoheight
                                    value = lv_value ).
      lv_value = o_worksheet->sheet_setup->fit_to_width.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fittowidth
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->horizontal_dpi IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->horizontal_dpi.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_horizontaldpi
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->orientation IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->orientation.
      lo_element->set_attribute_ns( name  = lc_xml_attr_orientation
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->page_order IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_pageorder
                                    value = o_worksheet->sheet_setup->page_order ).
    ENDIF.

    IF o_worksheet->sheet_setup->paper_height IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->paper_height.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_paperheight
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->paper_size IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->paper_size.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_papersize
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->paper_width IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->paper_width.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_paperwidth
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->scale IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->scale.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_scale
                                    value = lv_value ).
    ENDIF.

    IF o_worksheet->sheet_setup->use_first_page_num IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_usefirstpagenumber
                                    value = `1` ).
    ENDIF.

    IF o_worksheet->sheet_setup->use_printer_defaults IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_useprinterdefaults
                                    value = `1` ).
    ENDIF.

    IF o_worksheet->sheet_setup->vertical_dpi IS NOT INITIAL.
      lv_value = o_worksheet->sheet_setup->vertical_dpi.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_verticaldpi
                                    value = lv_value ).
    ENDIF.

    o_element_root->append_child( new_child = lo_element ). " pageSetup node
  ENDMETHOD.

  METHOD add_header_footer.
    DATA:
      lo_element   TYPE REF TO if_ixml_element,
      lo_element_2 TYPE REF TO if_ixml_element,
      lv_value     TYPE string.

    IF    o_worksheet->sheet_setup->odd_header IS NOT INITIAL
       OR o_worksheet->sheet_setup->odd_footer IS NOT INITIAL
       OR o_worksheet->sheet_setup->diff_oddeven_headerfooter = abap_true.

      lo_element = o_document->create_simple_element( name   = lc_xml_node_headerfooter
                                                      parent = o_document ).

      " Different header/footer for odd/even pages?
      IF o_worksheet->sheet_setup->diff_oddeven_headerfooter = abap_true.
        lo_element->set_attribute_ns( name  = lc_xml_attr_differentoddeven
                                      value = '1' ).
      ENDIF.

      " OddHeader
      CLEAR: lv_value.
      o_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_header = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_oddheader
                                                          parent = o_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " OddFooter
      CLEAR: lv_value.
      o_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_footer = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_oddfooter
                                                          parent = o_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " evenHeader
      CLEAR: lv_value.
      o_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_even_header = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_evenheader
                                                          parent = o_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " evenFooter
      CLEAR: lv_value.
      o_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_even_footer = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = o_document->create_simple_element( name   = lc_xml_node_evenfooter
                                                          parent = o_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      o_element_root->append_child( new_child = lo_element ). " headerFooter

    ENDIF.
  ENDMETHOD.

  METHOD add_drawing.
    DATA:
      lo_element  TYPE REF TO if_ixml_element,
      lv_value    TYPE string,
      lo_drawings TYPE REF TO zcl_excel_drawings.

    lo_drawings = o_worksheet->get_drawings( ).
    IF lo_drawings->is_empty( ) = abap_false.
      lo_element = o_document->create_simple_element( name   = lc_xml_node_drawing
                                                      parent = o_document ).
      ADD 1 TO v_relation_id.

      lv_value = v_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name  = 'r:id'
                                 value = lv_value ).
      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
  ENDMETHOD.

  METHOD add_drawing_for_comments.
    DATA: lo_element              TYPE REF TO if_ixml_element,
          lv_value                TYPE string,
          lo_drawing_for_comments TYPE REF TO zcl_excel_comments.
    " (Legacy) drawings for comments

    lo_drawing_for_comments = o_worksheet->get_comments( ).
    IF lo_drawing_for_comments->is_empty( ) = abap_false.
      lo_element = o_document->create_simple_element( name   = lc_xml_node_drawing_for_cmt
                                                      parent = o_document ).
      ADD 1 TO v_relation_id.  " +1 for legacyDrawings

      lv_value = v_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name  = 'r:id'
                                 value = lv_value ).
      o_element_root->append_child( new_child = lo_element ).

      ADD 1 TO v_relation_id.  "for xl/comments.xml (not referenced in XL sheet but counted in sheet rels)
    ENDIF.
* End   - Add - Issue #180
  ENDMETHOD.

  METHOD add_drawing_for_header_footer.
    DATA:
      lo_element  TYPE REF TO if_ixml_element,
      lv_value    TYPE string,
      lt_drawings TYPE zexcel_t_drawings.
* Header/Footer Image

    lt_drawings = o_worksheet->get_header_footer_drawings( ).
    IF lines( lt_drawings ) > 0. "Header or footer image exist
      lo_element = o_document->create_simple_element( name   = lc_xml_node_drawing_for_hd_ft
                                                      parent = o_document ).
      ADD 1 TO v_relation_id.  " +1 for legacyDrawings
      lv_value = v_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name  = 'r:id'
                                 value = lv_value ).
      o_element_root->append_child( new_child = lo_element ).
    ENDIF.
*
  ENDMETHOD.

  METHOD add_table_parts.
    DATA:
      lo_element     TYPE REF TO if_ixml_element,
      lo_element_2   TYPE REF TO if_ixml_element,
      lo_iterator    TYPE REF TO zcl_excel_collection_iterator,
      lo_table       TYPE REF TO zcl_excel_table,
      lv_table_count TYPE i,
      lv_value       TYPE string.

* tables

    lv_table_count = o_worksheet->get_tables_size( ).
    IF lv_table_count > 0.
      lo_element = o_document->create_simple_element( name   = 'tableParts'
                                                      parent = o_document ).
      lv_value = lv_table_count.
      CONDENSE lv_value.
      lo_element->set_attribute_ns( name  = 'count'
                                    value = lv_value ).

      lo_iterator = o_worksheet->get_tables_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_iterator->get_next( ).
        ADD 1 TO v_relation_id.

        lv_value = v_relation_id.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element_2 = o_document->create_simple_element( name   = 'tablePart'
                                                          parent = lo_element ).
        lo_element_2->set_attribute_ns( name  = 'r:id'
                                        value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).

      ENDWHILE.

      o_element_root->append_child( new_child = lo_element ).

    ENDIF.
  ENDMETHOD.

  METHOD add_sheet_data.
    DATA:
      lo_element     TYPE REF TO if_ixml_element.

    lo_element = o_excel_ref->create_xl_sheet_sheet_data( io_worksheet = o_worksheet
                                                          io_document  = o_document ).

    o_element_root->append_child( new_child = lo_element ). " sheetData node
  ENDMETHOD.

  METHOD add_page_breaks.

    TRY.
        o_excel_ref->create_xl_sheet_pagebreaks( io_document  = o_document
                                                 io_parent    = o_element_root
                                                 io_worksheet = o_worksheet ).
      CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    ENDTRY.

  ENDMETHOD.

  METHOD add_ignored_errors.

    o_excel_ref->create_xl_sheet_ignored_errors( io_worksheet    = o_worksheet
                                                 io_document     = o_document
                                                 io_element_root = o_element_root ).

  ENDMETHOD.

ENDCLASS.
