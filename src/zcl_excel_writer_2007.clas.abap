CLASS zcl_excel_writer_2007 DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    INTERFACES zif_excel_writer .
    METHODS constructor.

  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    TYPES: BEGIN OF mty_column_formula_used,
             id TYPE zexcel_s_cell_data-column_formula_id,
             si TYPE string,
             "! type: shared, etc.
             t  TYPE string,
           END OF mty_column_formula_used,
           mty_column_formulas_used TYPE HASHED TABLE OF mty_column_formula_used WITH UNIQUE KEY id.
    CONSTANTS c_content_types TYPE string VALUE '[Content_Types].xml'. "#EC NOTEXT
    CONSTANTS c_docprops_app TYPE string VALUE 'docProps/app.xml'. "#EC NOTEXT
    CONSTANTS c_docprops_core TYPE string VALUE 'docProps/core.xml'. "#EC NOTEXT
    CONSTANTS c_relationships TYPE string VALUE '_rels/.rels'. "#EC NOTEXT
    CONSTANTS c_xl_calcchain TYPE string VALUE 'xl/calcChain.xml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings TYPE string VALUE 'xl/drawings/drawing#.xml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings_rels TYPE string VALUE 'xl/drawings/_rels/drawing#.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_relationships TYPE string VALUE 'xl/_rels/workbook.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_sharedstrings TYPE string VALUE 'xl/sharedStrings.xml'. "#EC NOTEXT
    CONSTANTS c_xl_sheet TYPE string VALUE 'xl/worksheets/sheet#.xml'. "#EC NOTEXT
    CONSTANTS c_xl_sheet_rels TYPE string VALUE 'xl/worksheets/_rels/sheet#.xml.rels'. "#EC NOTEXT
    CONSTANTS c_xl_styles TYPE string VALUE 'xl/styles.xml'. "#EC NOTEXT
    CONSTANTS c_xl_theme TYPE string VALUE 'xl/theme/theme1.xml'. "#EC NOTEXT
    CONSTANTS c_xl_workbook TYPE string VALUE 'xl/workbook.xml'. "#EC NOTEXT
    DATA excel TYPE REF TO zcl_excel .
    DATA shared_strings TYPE zexcel_t_shared_string .
    DATA styles_cond_mapping TYPE zexcel_t_styles_cond_mapping .
    DATA styles_mapping TYPE zexcel_t_styles_mapping .
    CONSTANTS c_xl_comments TYPE string VALUE 'xl/comments#.xml'. "#EC NOTEXT
    CONSTANTS cl_xl_drawing_for_comments TYPE string VALUE 'xl/drawings/vmlDrawing#.vml'. "#EC NOTEXT
    CONSTANTS c_xl_drawings_vml_rels TYPE string VALUE 'xl/drawings/_rels/vmlDrawing#.vml.rels'. "#EC NOTEXT
    DATA ixml TYPE REF TO if_ixml.

    METHODS create_xl_sheet_sheet_data
      IMPORTING
        !io_document                   TYPE REF TO if_ixml_document
        !io_worksheet                  TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(rv_ixml_sheet_data_root) TYPE REF TO if_ixml_element
      RAISING
        zcx_excel .
    METHODS add_further_data_to_zip
      IMPORTING
        !io_zip TYPE REF TO cl_abap_zip .
    METHODS create
      RETURNING
        VALUE(ep_excel) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_content_types
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_docprops_app
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_docprops_core
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_dxf_style
      IMPORTING
        !iv_cell_style    TYPE zexcel_cell_style
        !io_dxf_element   TYPE REF TO if_ixml_element
        !io_ixml_document TYPE REF TO if_ixml_document
        !it_cellxfs       TYPE zexcel_t_cellxfs
        !it_fonts         TYPE zexcel_t_style_font
        !it_fills         TYPE zexcel_t_style_fill
      CHANGING
        !cv_dfx_count     TYPE i .
    METHODS create_relationships
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_charts
      IMPORTING
        !io_drawing       TYPE REF TO zcl_excel_drawing
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_comments
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawing_anchor
      IMPORTING
        !io_drawing      TYPE REF TO zcl_excel_drawing
        !io_document     TYPE REF TO if_ixml_document
        !ip_index        TYPE i
      RETURNING
        VALUE(ep_anchor) TYPE REF TO if_ixml_element .
    METHODS create_xl_drawing_for_comments
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_relationships
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_sharedstrings
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_sheet
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
        !iv_active        TYPE flag DEFAULT ''
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_sheet_ignored_errors
      IMPORTING
        io_worksheet    TYPE REF TO zcl_excel_worksheet
        io_document     TYPE REF TO if_ixml_document
        io_element_root TYPE REF TO if_ixml_element.
    METHODS create_xl_sheet_pagebreaks
      IMPORTING
        !io_document  TYPE REF TO if_ixml_document
        !io_parent    TYPE REF TO if_ixml_element
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS create_xl_sheet_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
        !iv_drawing_index TYPE i
        !iv_comment_index TYPE i
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_styles
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_styles_color_node
      IMPORTING
        !io_document        TYPE REF TO if_ixml_document
        !io_parent          TYPE REF TO if_ixml_element
        !iv_color_elem_name TYPE string DEFAULT 'color'
        !is_color           TYPE zexcel_s_style_color .
    METHODS create_xl_styles_font_node
      IMPORTING
        !io_document TYPE REF TO if_ixml_document
        !io_parent   TYPE REF TO if_ixml_element
        !is_font     TYPE zexcel_s_style_font
        !iv_use_rtf  TYPE abap_bool DEFAULT abap_false .
    METHODS create_xl_table
      IMPORTING
        !io_table         TYPE REF TO zcl_excel_table
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS create_xl_theme
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_workbook
      RETURNING
        VALUE(ep_content) TYPE xstring
      RAISING
        zcx_excel .
    METHODS get_shared_string_index
      IMPORTING
        !ip_cell_value  TYPE zexcel_cell_value
        !it_rtf         TYPE zexcel_t_rtf OPTIONAL
      RETURNING
        VALUE(ep_index) TYPE int4 .
    METHODS create_xl_drawings_vml
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS set_vml_string
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS create_xl_drawings_vml_rels
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS set_vml_shape_footer
      IMPORTING
        !is_footer        TYPE zexcel_s_worksheet_head_foot
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS set_vml_shape_header
      IMPORTING
        !is_header        TYPE zexcel_s_worksheet_head_foot
      RETURNING
        VALUE(ep_content) TYPE string .
    METHODS create_xl_drawing_for_hdft_im
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xl_drawings_hdft_rels
      IMPORTING
        !io_worksheet     TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ep_content) TYPE xstring .
    METHODS create_xml_document
      RETURNING
        VALUE(ro_document) TYPE REF TO if_ixml_document.
    METHODS render_xml_document
      IMPORTING
        io_document       TYPE REF TO if_ixml_document
      RETURNING
        VALUE(ep_content) TYPE xstring.
    METHODS create_xl_sheet_column_formula
      IMPORTING
        io_document             TYPE REF TO if_ixml_document
        it_column_formulas      TYPE zcl_excel_worksheet=>mty_th_column_formula
        is_sheet_content        TYPE zexcel_s_cell_data
      EXPORTING
        eo_element              TYPE REF TO if_ixml_element
      CHANGING
        ct_column_formulas_used TYPE mty_column_formulas_used
        cv_si                   TYPE i
      RAISING
        zcx_excel.
    METHODS is_formula_shareable
      IMPORTING
        ip_formula          TYPE string
      RETURNING
        VALUE(ep_shareable) TYPE abap_bool
      RAISING
        zcx_excel.
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_WRITER_2007
*"* do not include other source files here!!!
    CONSTANTS c_off TYPE string VALUE '0'.                  "#EC NOTEXT
    CONSTANTS c_on TYPE string VALUE '1'.                   "#EC NOTEXT
    CONSTANTS c_xl_printersettings TYPE string VALUE 'xl/printerSettings/printerSettings#.bin'. "#EC NOTEXT
    TYPES: tv_charbool TYPE c LENGTH 5.

    METHODS add_1_val_child_node
      IMPORTING
        io_document   TYPE REF TO if_ixml_document
        io_parent     TYPE REF TO if_ixml_element
        iv_elem_name  TYPE string
        iv_attr_name  TYPE string
        iv_attr_value TYPE string.
    METHODS flag2bool
      IMPORTING
        !ip_flag          TYPE flag
      RETURNING
        VALUE(ep_boolean) TYPE tv_charbool  .
ENDCLASS.



CLASS zcl_excel_writer_2007 IMPLEMENTATION.


  METHOD add_1_val_child_node.

    DATA: lo_child TYPE REF TO if_ixml_element.

    lo_child = io_document->create_simple_element( name   = iv_elem_name
                                                   parent = io_document ).
    IF iv_attr_name IS NOT INITIAL.
      lo_child->set_attribute_ns( name  = iv_attr_name
                                  value = iv_attr_value ).
    ENDIF.
    io_parent->append_child( new_child = lo_child ).

  ENDMETHOD.


  METHOD add_further_data_to_zip.
* Can be used by child classes like xlsm-writer to write additional data to zip archive
  ENDMETHOD.


  METHOD constructor.
    me->ixml = cl_ixml=>create( ).
  ENDMETHOD.


  METHOD create.

* Office 2007 file format is a cab of several xml files with extension .xlsx

    DATA: lo_zip              TYPE REF TO cl_abap_zip,
          lo_worksheet        TYPE REF TO zcl_excel_worksheet,
          lo_active_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lo_nested_iterator  TYPE REF TO zcl_excel_collection_iterator,
          lo_table            TYPE REF TO zcl_excel_table,
          lo_drawing          TYPE REF TO zcl_excel_drawing,
          lo_drawings         TYPE REF TO zcl_excel_drawings,
          lo_comment          TYPE REF TO zcl_excel_comment,   " (+) Issue #180
          lo_comments         TYPE REF TO zcl_excel_comments.  " (+) Issue #180

    DATA: lv_content                TYPE xstring,
          lv_active                 TYPE flag,
          lv_xl_sheet               TYPE string,
          lv_xl_sheet_rels          TYPE string,
          lv_xl_drawing_for_comment TYPE string,   " (+) Issue #180
          lv_xl_comment             TYPE string,   " (+) Issue #180
          lv_xl_drawing             TYPE string,
          lv_xl_drawing_rels        TYPE string,
          lv_index_str              TYPE string,
          lv_value                  TYPE string,
          lv_sheet_index            TYPE i,
          lv_drawing_index          TYPE i,
          lv_comment_index          TYPE i.        " (+) Issue #180

**********************************************************************

**********************************************************************
* Start of insertion # issue 139 - Dateretention of cellstyles
    me->excel->add_static_styles( ).
* End of insertion # issue 139 - Dateretention of cellstyles

**********************************************************************
* STEP 1: Create archive object file (ZIP)
    CREATE OBJECT lo_zip.

**********************************************************************
* STEP 2: Add [Content_Types].xml to zip
    lv_content = me->create_content_types( ).
    lo_zip->add( name    = me->c_content_types
                 content = lv_content ).

**********************************************************************
* STEP 3: Add _rels/.rels to zip
    lv_content = me->create_relationships( ).
    lo_zip->add( name    = me->c_relationships
                 content = lv_content ).

**********************************************************************
* STEP 4: Add docProps/app.xml to zip
    lv_content = me->create_docprops_app( ).
    lo_zip->add( name    = me->c_docprops_app
                 content = lv_content ).

**********************************************************************
* STEP 5: Add docProps/core.xml to zip
    lv_content = me->create_docprops_core( ).
    lo_zip->add( name    = me->c_docprops_core
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_relationships( ).
    lo_zip->add( name    = me->c_xl_relationships
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_theme( ).
    lo_zip->add( name    = me->c_xl_theme
                 content = lv_content ).

**********************************************************************
* STEP 7: Add xl/workbook.xml to zip
    lv_content = me->create_xl_workbook( ).
    lo_zip->add( name    = me->c_xl_workbook
                 content = lv_content ).

**********************************************************************
* STEP 8: Add xl/workbook.xml to zip
    lv_content = me->create_xl_styles( ).
    lo_zip->add( name    = me->c_xl_styles
                 content = lv_content ).

**********************************************************************
* STEP 9: Add sharedStrings.xml to zip
    lv_content = me->create_xl_sharedstrings( ).
    lo_zip->add( name    = me->c_xl_sharedstrings
                 content = lv_content ).

**********************************************************************
* STEP 10: Add sheet#.xml and drawing#.xml to zip
    lo_iterator = me->excel->get_worksheets_iterator( ).
    lo_active_worksheet = me->excel->get_active_worksheet( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lv_sheet_index = sy-index.

      lo_worksheet ?= lo_iterator->get_next( ).
      IF lo_active_worksheet->get_guid( ) EQ lo_worksheet->get_guid( ).
        lv_active = abap_true.
      ELSE.
        lv_active = abap_false.
      ENDIF.
      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.

      lv_index_str = lv_sheet_index.
      CONDENSE lv_index_str NO-GAPS.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet WITH lv_index_str.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

* Begin - Add - Issue #180
* Add comments **********************************
      lo_comments = lo_worksheet->get_comments( ).
      IF lo_comments->is_empty( ) = abap_false.
        lv_comment_index = lv_comment_index + 1.

        " Create comment itself
        lv_content = me->create_xl_comments( lo_worksheet ).
        lv_xl_comment = me->c_xl_comments.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_comment
                     content = lv_content ).

        " Create vmlDrawing that will host the comment
        lv_content = me->create_xl_drawing_for_comments( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_for_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).
      ENDIF.
* End   - Add - Issue #180

* Add drawings **********************************
      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_drawing_index = lv_drawing_index + 1.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        lv_index_str = lv_drawing_index.
        CONDENSE lv_index_str NO-GAPS.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      ENDIF.

* Add Header/Footer image
      DATA: lt_drawings TYPE zexcel_t_drawings.
      lt_drawings = lo_worksheet->get_header_footer_drawings( ).
      IF lines( lt_drawings ) > 0. "Header or footer image exist

        lv_comment_index = lv_comment_index + 1.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.

        " Create vmlDrawing that will host the image
        lv_content = me->create_xl_drawing_for_hdft_im( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_for_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).

        " Create vmlDrawing REL that will host the image
        lv_content = me->create_xl_drawings_hdft_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_vml_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      ENDIF.


      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index ).      " (+) Issue #180

      lv_index_str = lv_sheet_index.
      CONDENSE lv_index_str NO-GAPS.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet_rels WITH lv_index_str.
      lo_zip->add( name    = lv_xl_sheet_rels
                   content = lv_content ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      WHILE lo_nested_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).
        lv_content = me->create_xl_table( lo_table ).

        lv_value = lo_table->get_name( ).
        CONCATENATE 'xl/tables/' lv_value '.xml' INTO lv_value.
        lo_zip->add( name = lv_value
                      content = lv_content ).
      ENDWHILE.



    ENDWHILE.

**********************************************************************
* STEP 11: Add media
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).
      lv_value = lo_drawing->get_media_name( ).
      CONCATENATE 'xl/media/' lv_value INTO lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    ENDWHILE.

**********************************************************************
* STEP 12: Add charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).

      "-------------Added by Alessandro Iannacci - Only if template exist
      IF lv_content IS NOT INITIAL AND me->excel->use_template EQ abap_true.
        lv_value = lo_drawing->get_media_name( ).
        CONCATENATE 'xl/charts/' lv_value INTO lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      ELSE. "ADD CUSTOM CHART!!!!
        lv_content = me->create_xl_charts( lo_drawing ).
        lv_value = lo_drawing->get_media_name( ).
        CONCATENATE 'xl/charts/' lv_value INTO lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      ENDIF.
      "-------------------------------------------------
    ENDWHILE.

* Second to last step: Allow further information put into the zip archive by child classes
    me->add_further_data_to_zip( lo_zip ).

**********************************************************************
* Last step: Create the final zip
    ep_excel = lo_zip->save( ).

  ENDMETHOD.


  METHOD create_content_types.


** Constant node name
    DATA: lc_xml_node_types        TYPE string VALUE 'Types',
          lc_xml_node_override     TYPE string VALUE 'Override',
          lc_xml_node_default      TYPE string VALUE 'Default',
          " Node attributes
          lc_xml_attr_partname     TYPE string VALUE 'PartName',
          lc_xml_attr_extension    TYPE string VALUE 'Extension',
          lc_xml_attr_contenttype  TYPE string VALUE 'ContentType',
          " Node namespace
          lc_xml_node_types_ns     TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/content-types',
          " Node extension
          lc_xml_node_rels_ext     TYPE string VALUE 'rels',
          lc_xml_node_xml_ext      TYPE string VALUE 'xml',
          lc_xml_node_xml_vml      TYPE string VALUE 'vml',   " (+) GGAR
          " Node partnumber
          lc_xml_node_theme_pn     TYPE string VALUE '/xl/theme/theme1.xml',
          lc_xml_node_styles_pn    TYPE string VALUE '/xl/styles.xml',
          lc_xml_node_workb_pn     TYPE string VALUE '/xl/workbook.xml',
          lc_xml_node_props_pn     TYPE string VALUE '/docProps/app.xml',
          lc_xml_node_worksheet_pn TYPE string VALUE '/xl/worksheets/sheet#.xml',
          lc_xml_node_strings_pn   TYPE string VALUE '/xl/sharedStrings.xml',
          lc_xml_node_core_pn      TYPE string VALUE '/docProps/core.xml',
          lc_xml_node_chart_pn     TYPE string VALUE '/xl/charts/chart#.xml',
          " Node contentType
          lc_xml_node_theme_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.theme+xml',
          lc_xml_node_styles_ct    TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          lc_xml_node_workb_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          lc_xml_node_rels_ct      TYPE string VALUE 'application/vnd.openxmlformats-package.relationships+xml',
          lc_xml_node_vml_ct       TYPE string VALUE 'application/vnd.openxmlformats-officedocument.vmlDrawing',
          lc_xml_node_xml_ct       TYPE string VALUE 'application/xml',
          lc_xml_node_props_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
          lc_xml_node_worksheet_ct TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
          lc_xml_node_strings_ct   TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          lc_xml_node_core_ct      TYPE string VALUE 'application/vnd.openxmlformats-package.core-properties+xml',
          lc_xml_node_table_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
          lc_xml_node_comments_ct  TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',   " (+) GGAR
          lc_xml_node_drawings_ct  TYPE string VALUE 'application/vnd.openxmlformats-officedocument.drawing+xml',
          lc_xml_node_chart_ct     TYPE string VALUE 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_element         TYPE REF TO if_ixml_element,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_nested_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_table           TYPE REF TO zcl_excel_table.

    DATA: lv_worksheets_num        TYPE i,
          lv_worksheets_numc       TYPE n LENGTH 3,
          lv_xml_node_worksheet_pn TYPE string,
          lv_value                 TYPE string,
          lv_comment_index         TYPE i VALUE 1,  " (+) GGAR
          lv_drawing_index         TYPE i VALUE 1,
          lv_index_str             TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node types
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_types
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
    value = lc_xml_node_types_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_rels_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_rels_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_xml_ct ).
    lo_element_root->append_child( new_child = lo_element ).

* Begin - Add - GGAR
    " VML node (for comments)
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_xml_vml ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_vml_ct ).
    lo_element_root->append_child( new_child = lo_element ).
* End   - Add - GGAR

    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_theme_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_theme_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_styles_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_styles_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Workbook node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_workb_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_workb_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Properties node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_props_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_props_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Worksheet node
    lv_worksheets_num = excel->get_worksheets_size( ).
    DO lv_worksheets_num TIMES.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).

      lv_worksheets_numc = sy-index.
      SHIFT lv_worksheets_numc LEFT DELETING LEADING '0'.
      lv_xml_node_worksheet_pn = lc_xml_node_worksheet_pn.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_worksheet_pn WITH lv_worksheets_numc.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_xml_node_worksheet_pn ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_worksheet_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDDO.

    lo_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      WHILE lo_nested_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).

        lv_value = lo_table->get_name( ).
        CONCATENATE '/xl/tables/' lv_value '.xml' INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_table_ct ).
        lo_element_root->append_child( new_child = lo_element ).
      ENDWHILE.

* Begin - Add - GGAR
      " Comments
      DATA: lo_comments TYPE REF TO zcl_excel_comments.

      lo_comments = lo_worksheet->get_comments( ).
      IF lo_comments->is_empty( ) = abap_false.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.
        CONCATENATE '/' me->c_xl_comments INTO lv_value.
        REPLACE '#' WITH lv_index_str INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                         parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_comments_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        ADD 1 TO lv_comment_index.
      ENDIF.
* End   - Add - GGAR

      " Drawings
      DATA: lo_drawings TYPE REF TO zcl_excel_drawings.

      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_index_str = lv_drawing_index.
        CONDENSE lv_index_str NO-GAPS.
        CONCATENATE '/' me->c_xl_drawings INTO lv_value.
        REPLACE '#' WITH lv_index_str INTO lv_value.

        lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_drawings_ct ).
        lo_element_root->append_child( new_child = lo_element ).

        ADD 1 TO lv_drawing_index.
      ENDIF.
    ENDWHILE.

    " media mimes
    DATA: lo_drawing    TYPE REF TO zcl_excel_drawing,
          lt_media_type TYPE TABLE OF mimetypes-extension,
          lv_media_type TYPE mimetypes-extension,
          lv_mime_type  TYPE mimetypes-type.

    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_media_type = lo_drawing->get_media_type( ).
      COLLECT lv_media_type INTO lt_media_type.
    ENDWHILE.

    LOOP AT lt_media_type INTO lv_media_type.
      CALL FUNCTION 'SDOK_MIMETYPE_GET'
        EXPORTING
          extension = lv_media_type
        IMPORTING
          mimetype  = lv_mime_type.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                       parent = lo_document ).
      lv_value = lv_media_type.
      lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                    value = lv_value ).
      lv_value = lv_mime_type.
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDLOOP.

    " Charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                       parent = lo_document ).
      lv_index_str = lo_drawing->get_index( ).
      CONDENSE lv_index_str.
      lv_value = lc_xml_node_chart_pn.
      REPLACE ALL OCCURRENCES OF '#' IN lv_value WITH lv_index_str.
      lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                    value = lc_xml_node_chart_ct ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_strings_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_strings_ct ).
    lo_element_root->append_child( new_child = lo_element ).

    " Strings node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_override
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_partname
                                  value = lc_xml_node_core_pn ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_core_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.


  METHOD create_docprops_app.


** Constant node name
    DATA: lc_xml_node_properties        TYPE string VALUE 'Properties',
          lc_xml_node_application       TYPE string VALUE 'Application',
          lc_xml_node_docsecurity       TYPE string VALUE 'DocSecurity',
          lc_xml_node_scalecrop         TYPE string VALUE 'ScaleCrop',
          lc_xml_node_headingpairs      TYPE string VALUE 'HeadingPairs',
          lc_xml_node_vector            TYPE string VALUE 'vector',
          lc_xml_node_variant           TYPE string VALUE 'variant',
          lc_xml_node_lpstr             TYPE string VALUE 'lpstr',
          lc_xml_node_i4                TYPE string VALUE 'i4',
          lc_xml_node_titlesofparts     TYPE string VALUE 'TitlesOfParts',
          lc_xml_node_company           TYPE string VALUE 'Company',
          lc_xml_node_linksuptodate     TYPE string VALUE 'LinksUpToDate',
          lc_xml_node_shareddoc         TYPE string VALUE 'SharedDoc',
          lc_xml_node_hyperlinkschanged TYPE string VALUE 'HyperlinksChanged',
          lc_xml_node_appversion        TYPE string VALUE 'AppVersion',
          " Namespace prefix
          lc_vt_ns                      TYPE string VALUE 'vt',
          lc_xml_node_props_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
          lc_xml_node_props_vt_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
          " Node attributes
          lc_xml_attr_size              TYPE string VALUE 'size',
          lc_xml_attr_basetype          TYPE string VALUE 'baseType'.

    DATA: lo_document            TYPE REF TO if_ixml_document,
          lo_element_root        TYPE REF TO if_ixml_element,
          lo_element             TYPE REF TO if_ixml_element,
          lo_sub_element_vector  TYPE REF TO if_ixml_element,
          lo_sub_element_variant TYPE REF TO if_ixml_element,
          lo_sub_element_lpstr   TYPE REF TO if_ixml_element,
          lo_sub_element_i4      TYPE REF TO if_ixml_element,
          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet.

    DATA: lv_value                TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node properties
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_properties
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_props_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:vt'
                                       value = lc_xml_node_props_vt_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Application
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_application
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~application.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " DocSecurity
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_docsecurity
                                                              parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~docsecurity.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " ScaleCrop
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_scalecrop
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~scalecrop ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HeadingPairs
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_headingpairs
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = '2' ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_variant ).

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** lpstr node
    lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                  prefix = lc_vt_ns
                                                                  parent = lo_document ).
    lv_value = excel->get_worksheets_name( ).
    lo_sub_element_lpstr->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_lpstr ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    " ** variant node
    lo_sub_element_variant = lo_document->create_simple_element_ns( name   = lc_xml_node_variant
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).

    " *** i4 node
    lo_sub_element_i4 = lo_document->create_simple_element_ns( name   = lc_xml_node_i4
                                                               prefix = lc_vt_ns
                                                               parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element_i4->set_value( value = lv_value ).
    lo_sub_element_variant->append_child( new_child = lo_sub_element_i4 ). " lpstr node

    lo_sub_element_vector->append_child( new_child = lo_sub_element_variant ). " variant node

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " HeadingPairs


    " TitlesOfParts
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_titlesofparts
                                                     parent = lo_document ).


    " * vector node
    lo_sub_element_vector = lo_document->create_simple_element_ns( name   = lc_xml_node_vector
                                                                   prefix = lc_vt_ns
                                                                   parent = lo_document ).
    lv_value = excel->get_worksheets_size( ).
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_size
                                             value   = lv_value ).
    lo_sub_element_vector->set_attribute_ns( name    = lc_xml_attr_basetype
                                             value   = lc_xml_node_lpstr ).

    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " ** lpstr node
      lo_sub_element_lpstr = lo_document->create_simple_element_ns( name   = lc_xml_node_lpstr
                                                                    prefix = lc_vt_ns
                                                                    parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_value = lo_worksheet->get_title( ).
      lo_sub_element_lpstr->set_value( value = lv_value ).
      lo_sub_element_vector->append_child( new_child = lo_sub_element_lpstr ). " lpstr node
    ENDWHILE.

    lo_element->append_child( new_child = lo_sub_element_vector ). " vector node

    lo_element_root->append_child( new_child = lo_element ). " TitlesOfParts



    " Company
    IF excel->zif_excel_book_properties~company IS NOT INITIAL.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_company
                                                       parent = lo_document ).
      lv_value = excel->zif_excel_book_properties~company.
      lo_element->set_value( value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    " LinksUpToDate
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_linksuptodate
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~linksuptodate ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " SharedDoc
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_shareddoc
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~shareddoc ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " HyperlinksChanged
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_hyperlinkschanged
                                                     parent = lo_document ).
    lv_value = me->flag2bool( excel->zif_excel_book_properties~hyperlinkschanged ).
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " AppVersion
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_appversion
                                                     parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~appversion.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.


  METHOD create_docprops_core.


** Constant node name
    DATA: lc_xml_node_coreproperties TYPE string VALUE 'coreProperties',
          lc_xml_node_creator        TYPE string VALUE 'creator',
          lc_xml_node_description    TYPE string VALUE 'description',
          lc_xml_node_lastmodifiedby TYPE string VALUE 'lastModifiedBy',
          lc_xml_node_created        TYPE string VALUE 'created',
          lc_xml_node_modified       TYPE string VALUE 'modified',
          " Node attributes
          lc_xml_attr_type           TYPE string VALUE 'type',
          lc_xml_attr_target         TYPE string VALUE 'dcterms:W3CDTF',
          " Node namespace
          lc_cp_ns                   TYPE string VALUE 'cp',
          lc_dc_ns                   TYPE string VALUE 'dc',
          lc_dcterms_ns              TYPE string VALUE 'dcterms',
*        lc_dcmitype_ns              TYPE string VALUE 'dcmitype',
          lc_xsi_ns                  TYPE string VALUE 'xsi',
          lc_xml_node_cp_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
          lc_xml_node_dc_ns          TYPE string VALUE 'http://purl.org/dc/elements/1.1/',
          lc_xml_node_dcterms_ns     TYPE string VALUE 'http://purl.org/dc/terms/',
          lc_xml_node_dcmitype_ns    TYPE string VALUE 'http://purl.org/dc/dcmitype/',
          lc_xml_node_xsi_ns         TYPE string VALUE 'http://www.w3.org/2001/XMLSchema-instance'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

    DATA: lv_value TYPE string,
          lv_date  TYPE d,
          lv_time  TYPE t.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node coreProperties
    lo_element_root  = lo_document->create_simple_element_ns( name   = lc_xml_node_coreproperties
                                                              prefix = lc_cp_ns
                                                              parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:cp'
                                       value = lc_xml_node_cp_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dc'
                                       value = lc_xml_node_dc_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcterms'
                                       value = lc_xml_node_dcterms_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:dcmitype'
                                       value = lc_xml_node_dcmitype_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xsi'
                                       value = lc_xml_node_xsi_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Creator node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_creator
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~creator.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Description node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_description
                                                        prefix = lc_dc_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~description.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " lastModifiedBy node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_lastmodifiedby
                                                        prefix = lc_cp_ns
                                                        parent = lo_document ).
    lv_value = excel->zif_excel_book_properties~lastmodifiedby.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Created node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_created
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).

    CONVERT TIME STAMP excel->zif_excel_book_properties~created TIME ZONE sy-zonlo INTO DATE lv_date TIME lv_time.
    CONCATENATE lv_date lv_time INTO lv_value RESPECTING BLANKS.
    REPLACE ALL OCCURRENCES OF REGEX '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' IN lv_value WITH '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " Modified node
    lo_element = lo_document->create_simple_element_ns( name   = lc_xml_node_modified
                                                        prefix = lc_dcterms_ns
                                                        parent = lo_document ).
    lo_element->set_attribute_ns( name    = lc_xml_attr_type
                                  prefix  = lc_xsi_ns
                                  value   = lc_xml_attr_target ).
    CONVERT TIME STAMP excel->zif_excel_book_properties~modified TIME ZONE sy-zonlo INTO DATE lv_date TIME lv_time.
    CONCATENATE lv_date lv_time INTO lv_value RESPECTING BLANKS.
    REPLACE ALL OCCURRENCES OF REGEX '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' IN lv_value WITH '$1-$2-$3T$4:$5:$6Z'.
    lo_element->set_value( value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.


  METHOD create_dxf_style.

    CONSTANTS: lc_xml_node_dxf         TYPE string VALUE 'dxf',
               lc_xml_node_font        TYPE string VALUE 'font',
               lc_xml_node_b           TYPE string VALUE 'b',            "bold
               lc_xml_node_i           TYPE string VALUE 'i',            "italic
               lc_xml_node_u           TYPE string VALUE 'u',            "underline
               lc_xml_node_strike      TYPE string VALUE 'strike',       "strikethrough
               lc_xml_attr_val         TYPE string VALUE 'val',
               lc_xml_node_fill        TYPE string VALUE 'fill',
               lc_xml_node_patternfill TYPE string VALUE 'patternFill',
               lc_xml_attr_patterntype TYPE string VALUE 'patternType',
               lc_xml_node_fgcolor     TYPE string VALUE 'fgColor',
               lc_xml_node_bgcolor     TYPE string VALUE 'bgColor'.

    DATA: ls_styles_mapping     TYPE zexcel_s_styles_mapping,
          ls_cellxfs            TYPE zexcel_s_cellxfs,
          ls_style_cond_mapping TYPE zexcel_s_styles_cond_mapping,
          lo_sub_element        TYPE REF TO if_ixml_element,
          lo_sub_element_2      TYPE REF TO if_ixml_element,
          lv_index              TYPE i,
          ls_font               TYPE zexcel_s_style_font,
          lo_element_font       TYPE REF TO if_ixml_element,
          lv_value              TYPE string,
          ls_fill               TYPE zexcel_s_style_fill,
          lo_element_fill       TYPE REF TO if_ixml_element.

    CHECK iv_cell_style IS NOT INITIAL.

    READ TABLE me->styles_mapping INTO ls_styles_mapping WITH KEY guid = iv_cell_style.
    ADD 1 TO ls_styles_mapping-style. " the numbering starts from 0
    READ TABLE it_cellxfs INTO ls_cellxfs INDEX ls_styles_mapping-style.
    ADD 1 TO ls_cellxfs-fillid.       " the numbering starts from 0

    READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY style = ls_styles_mapping-style.
    IF sy-subrc EQ 0.
      ls_style_cond_mapping-guid  = iv_cell_style.
      APPEND ls_style_cond_mapping TO me->styles_cond_mapping.
    ELSE.
      ls_style_cond_mapping-guid  = iv_cell_style.
      ls_style_cond_mapping-style = ls_styles_mapping-style.
      ls_style_cond_mapping-dxf   = cv_dfx_count.
      APPEND ls_style_cond_mapping TO me->styles_cond_mapping.
      ADD 1 TO cv_dfx_count.

      " dxf node
      lo_sub_element = io_ixml_document->create_simple_element( name   = lc_xml_node_dxf
                                                                parent = io_ixml_document ).

      "Conditional formatting font style correction by Alessandro Iannacci START
      lv_index = ls_cellxfs-fontid + 1.
      READ TABLE it_fonts INTO ls_font INDEX lv_index.
      IF ls_font IS NOT INITIAL.
        lo_element_font = io_ixml_document->create_simple_element( name   = lc_xml_node_font
                                                              parent = io_ixml_document ).
        IF ls_font-bold EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_b
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-italic EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_i
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-underline EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_u
                                                               parent = io_ixml_document ).
          lv_value = ls_font-underline_mode.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_val
                                            value = lv_value ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        IF ls_font-strikethrough EQ abap_true.
          lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_strike
                                                               parent = io_ixml_document ).
          lo_element_font->append_child( new_child = lo_sub_element_2 ).
        ENDIF.
        "color
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_element_font
            is_color           = ls_font-color ).
        lo_sub_element->append_child( new_child = lo_element_font ).
      ENDIF.
      "---Conditional formatting font style correction by Alessandro Iannacci END


      READ TABLE it_fills INTO ls_fill INDEX ls_cellxfs-fillid.
      IF ls_fill IS NOT INITIAL.
        " fill properties
        lo_element_fill = io_ixml_document->create_simple_element( name   = lc_xml_node_fill
                                                                 parent = io_ixml_document ).
        "pattern
        lo_sub_element_2 = io_ixml_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = io_ixml_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_patterntype
                                            value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = io_ixml_document
            io_parent          = lo_sub_element_2
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        IF  ls_fill-fgcolor-rgb IS INITIAL AND
          ls_fill-fgcolor-indexed EQ zcl_excel_style_color=>c_indexed_not_set AND
          ls_fill-fgcolor-theme EQ zcl_excel_style_color=>c_theme_not_set AND
          ls_fill-fgcolor-tint IS INITIAL AND ls_fill-bgcolor-indexed EQ zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = io_ixml_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        ENDIF.

        lo_element_fill->append_child( new_child = lo_sub_element_2 ). "pattern

        lo_sub_element->append_child( new_child = lo_element_fill ).
      ENDIF.
    ENDIF.

    io_dxf_element->append_child( new_child = lo_sub_element ).
  ENDMETHOD.


  METHOD create_relationships.


** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_rid1_id       TYPE string VALUE 'rId1',
          lc_xml_node_rid2_id       TYPE string VALUE 'rId2',
          lc_xml_node_rid3_id       TYPE string VALUE 'rId3',
          " Node type
          lc_xml_node_rid1_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
          lc_xml_node_rid2_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
          lc_xml_node_rid3_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
          " Node target
          lc_xml_node_rid1_tg       TYPE string VALUE 'xl/workbook.xml',
          lc_xml_node_rid2_tg       TYPE string VALUE 'docProps/core.xml',
          lc_xml_node_rid3_tg       TYPE string VALUE 'docProps/app.xml'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes
    " Theme node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid3_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid3_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid3_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " Styles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid2_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid2_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid2_tg ).
    lo_element_root->append_child( new_child = lo_element ).

    " rels node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lc_xml_node_rid1_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid1_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid1_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.


  METHOD create_xl_charts.


** Constant node name
    CONSTANTS: lc_xml_node_chartspace         TYPE string VALUE 'c:chartSpace',
               lc_xml_node_ns_c               TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_ns_a               TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main',
               lc_xml_node_ns_r               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
               lc_xml_node_date1904           TYPE string VALUE 'c:date1904',
               lc_xml_node_lang               TYPE string VALUE 'c:lang',
               lc_xml_node_roundedcorners     TYPE string VALUE 'c:roundedCorners',
               lc_xml_node_altcont            TYPE string VALUE 'mc:AlternateContent',
               lc_xml_node_altcont_ns_mc      TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
               lc_xml_node_choice             TYPE string VALUE 'mc:Choice',
               lc_xml_node_choice_ns_requires TYPE string VALUE 'c14',
               lc_xml_node_choice_ns_c14      TYPE string VALUE 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
               lc_xml_node_style              TYPE string VALUE 'c14:style',
               lc_xml_node_fallback           TYPE string VALUE 'mc:Fallback',
               lc_xml_node_style2             TYPE string VALUE 'c:style',

               "---------------------------CHART
               lc_xml_node_chart              TYPE string VALUE 'c:chart',
               lc_xml_node_autotitledeleted   TYPE string VALUE 'c:autoTitleDeleted',
               "plotArea
               lc_xml_node_plotarea           TYPE string VALUE 'c:plotArea',
               lc_xml_node_layout             TYPE string VALUE 'c:layout',
               lc_xml_node_varycolors         TYPE string VALUE 'c:varyColors',
               lc_xml_node_ser                TYPE string VALUE 'c:ser',
               lc_xml_node_idx                TYPE string VALUE 'c:idx',
               lc_xml_node_order              TYPE string VALUE 'c:order',
               lc_xml_node_tx                 TYPE string VALUE 'c:tx',
               lc_xml_node_v                  TYPE string VALUE 'c:v',
               lc_xml_node_val                TYPE string VALUE 'c:val',
               lc_xml_node_cat                TYPE string VALUE 'c:cat',
               lc_xml_node_numref             TYPE string VALUE 'c:numRef',
               lc_xml_node_strref             TYPE string VALUE 'c:strRef',
               lc_xml_node_f                  TYPE string VALUE 'c:f', "this is the range
               lc_xml_node_overlap            TYPE string VALUE 'c:overlap',
               "note: numcache avoided
               lc_xml_node_dlbls              TYPE string VALUE 'c:dLbls',
               lc_xml_node_showlegendkey      TYPE string VALUE 'c:showLegendKey',
               lc_xml_node_showval            TYPE string VALUE 'c:showVal',
               lc_xml_node_showcatname        TYPE string VALUE 'c:showCatName',
               lc_xml_node_showsername        TYPE string VALUE 'c:showSerName',
               lc_xml_node_showpercent        TYPE string VALUE 'c:showPercent',
               lc_xml_node_showbubblesize     TYPE string VALUE 'c:showBubbleSize',
               "plotArea->pie
               lc_xml_node_piechart           TYPE string VALUE 'c:pieChart',
               lc_xml_node_showleaderlines    TYPE string VALUE 'c:showLeaderLines',
               lc_xml_node_firstsliceang      TYPE string VALUE 'c:firstSliceAng',
               "plotArea->line
               lc_xml_node_linechart          TYPE string VALUE 'c:lineChart',
               lc_xml_node_symbol             TYPE string VALUE 'c:symbol',
               lc_xml_node_marker             TYPE string VALUE 'c:marker',
               lc_xml_node_smooth             TYPE string VALUE 'c:smooth',
               "plotArea->bar
               lc_xml_node_invertifnegative   TYPE string VALUE 'c:invertIfNegative',
               lc_xml_node_barchart           TYPE string VALUE 'c:barChart',
               lc_xml_node_bardir             TYPE string VALUE 'c:barDir',
               lc_xml_node_gapwidth           TYPE string VALUE 'c:gapWidth',
               "plotArea->line + plotArea->bar
               lc_xml_node_grouping           TYPE string VALUE 'c:grouping',
               lc_xml_node_axid               TYPE string VALUE 'c:axId',
               lc_xml_node_catax              TYPE string VALUE 'c:catAx',
               lc_xml_node_valax              TYPE string VALUE 'c:valAx',
               lc_xml_node_scaling            TYPE string VALUE 'c:scaling',
               lc_xml_node_orientation        TYPE string VALUE 'c:orientation',
               lc_xml_node_delete             TYPE string VALUE 'c:delete',
               lc_xml_node_axpos              TYPE string VALUE 'c:axPos',
               lc_xml_node_numfmt             TYPE string VALUE 'c:numFmt',
               lc_xml_node_majorgridlines     TYPE string VALUE 'c:majorGridlines',
               lc_xml_node_majortickmark      TYPE string VALUE 'c:majorTickMark',
               lc_xml_node_minortickmark      TYPE string VALUE 'c:minorTickMark',
               lc_xml_node_ticklblpos         TYPE string VALUE 'c:tickLblPos',
               lc_xml_node_crossax            TYPE string VALUE 'c:crossAx',
               lc_xml_node_crosses            TYPE string VALUE 'c:crosses',
               lc_xml_node_auto               TYPE string VALUE 'c:auto',
               lc_xml_node_lblalgn            TYPE string VALUE 'c:lblAlgn',
               lc_xml_node_lbloffset          TYPE string VALUE 'c:lblOffset',
               lc_xml_node_nomultilvllbl      TYPE string VALUE 'c:noMultiLvlLbl',
               lc_xml_node_crossbetween       TYPE string VALUE 'c:crossBetween',
               "legend
               lc_xml_node_legend             TYPE string VALUE 'c:legend',
               "legend->pie
               lc_xml_node_legendpos          TYPE string VALUE 'c:legendPos',
*                  lc_xml_node_layout            TYPE string VALUE 'c:layout', "already exist
               lc_xml_node_overlay            TYPE string VALUE 'c:overlay',
               lc_xml_node_txpr               TYPE string VALUE 'c:txPr',
               lc_xml_node_bodypr             TYPE string VALUE 'a:bodyPr',
               lc_xml_node_lststyle           TYPE string VALUE 'a:lstStyle',
               lc_xml_node_p                  TYPE string VALUE 'a:p',
               lc_xml_node_ppr                TYPE string VALUE 'a:pPr',
               lc_xml_node_defrpr             TYPE string VALUE 'a:defRPr',
               lc_xml_node_endpararpr         TYPE string VALUE 'a:endParaRPr',
               "legend->bar + legend->line
               lc_xml_node_plotvisonly        TYPE string VALUE 'c:plotVisOnly',
               lc_xml_node_dispblanksas       TYPE string VALUE 'c:dispBlanksAs',
               lc_xml_node_showdlblsovermax   TYPE string VALUE 'c:showDLblsOverMax',
               "---------------------------END OF CHART

               lc_xml_node_printsettings      TYPE string VALUE 'c:printSettings',
               lc_xml_node_headerfooter       TYPE string VALUE 'c:headerFooter',
               lc_xml_node_pagemargins        TYPE string VALUE 'c:pageMargins',
               lc_xml_node_pagesetup          TYPE string VALUE 'c:pageSetup'.


    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element.


    DATA lo_element                               TYPE REF TO if_ixml_element.
    DATA lo_element2                              TYPE REF TO if_ixml_element.
    DATA lo_element3                              TYPE REF TO if_ixml_element.
    DATA lo_el_rootchart                           TYPE REF TO if_ixml_element.
    DATA lo_element4                              TYPE REF TO if_ixml_element.
    DATA lo_element5                              TYPE REF TO if_ixml_element.
    DATA lo_element6                              TYPE REF TO if_ixml_element.
    DATA lo_element7                              TYPE REF TO if_ixml_element.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_chartspace
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).

**********************************************************************
* STEP 4: Create chart

    DATA lo_chartb TYPE REF TO zcl_excel_graph_bars.
    DATA lo_chartp TYPE REF TO zcl_excel_graph_pie.
    DATA lo_chartl TYPE REF TO zcl_excel_graph_line.
    DATA lo_chart TYPE REF TO zcl_excel_graph.

    DATA ls_serie TYPE zcl_excel_graph=>s_series.
    DATA ls_ax TYPE zcl_excel_graph_bars=>s_ax.
    DATA lv_str TYPE string.

    "Identify chart type
    CASE io_drawing->graph_type.
      WHEN zcl_excel_drawing=>c_graph_bars.
        lo_chartb ?= io_drawing->graph.
      WHEN zcl_excel_drawing=>c_graph_pie.
        lo_chartp ?= io_drawing->graph.
      WHEN zcl_excel_drawing=>c_graph_line.
        lo_chartl ?= io_drawing->graph.
      WHEN OTHERS.
    ENDCASE.


    lo_chart = io_drawing->graph.

    lo_element = lo_document->create_simple_element( name = lc_xml_node_date1904
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_1904val ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_lang
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_langval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_roundedcorners
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_roundedcornersval ).

    lo_element = lo_document->create_simple_element( name = lc_xml_node_altcont
                                                         parent = lo_element_root ).
    lo_element->set_attribute_ns( name  = 'xmlns:mc'
                                      value = lc_xml_node_altcont_ns_mc ).

    "Choice
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_choice
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'Requires'
                                      value = lc_xml_node_choice_ns_requires ).
    lo_element2->set_attribute_ns( name  = 'xmlns:c14'
                                      value = lc_xml_node_choice_ns_c14 ).

    "C14:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_c14styleval ).

    "Fallback
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_fallback
                                                         parent = lo_element ).

    "C:style
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_style2
                                                         parent = lo_element2 ).
    lo_element3->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_styleval ).

    "---------------------------CHART
    lo_element = lo_document->create_simple_element( name = lc_xml_node_chart
                                                         parent = lo_element_root ).
    "Added
    IF lo_chart->title IS NOT INITIAL.
      lo_element2 = lo_document->create_simple_element( name = 'c:title'
                                                           parent = lo_element ).
      lo_element3 = lo_document->create_simple_element( name = 'c:tx'
                                                           parent = lo_element2 ).
      lo_element4 = lo_document->create_simple_element( name = 'c:rich'
                                                           parent = lo_element3 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:bodyPr'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:lstStyle'
                                                           parent = lo_element4 ).
      lo_element5 = lo_document->create_simple_element( name = 'a:p'
                                                           parent = lo_element4 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:pPr'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:defRPr'
                                                           parent = lo_element6 ).
      lo_element6 = lo_document->create_simple_element( name = 'a:r'
                                                           parent = lo_element5 ).
      lo_element7 = lo_document->create_simple_element( name = 'a:rPr'
                                                           parent = lo_element6 ).
      lo_element7->set_attribute_ns( name  = 'lang'
                                        value = 'en-US' ).
      lo_element7 = lo_document->create_simple_element( name = 'a:t'
                                                           parent = lo_element6 ).
      lo_element7->set_value( value = lo_chart->title ).
    ENDIF.
    "End
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_autotitledeleted
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_autotitledeletedval ).

    "plotArea
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotarea
                                                       parent = lo_element ).
    lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
    CASE io_drawing->graph_type.
      WHEN zcl_excel_drawing=>c_graph_bars.
        "----bar
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_barchart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bardir
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_bardirval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_varycolorsval ).

        "series
        LOOP AT lo_chartb->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_invertifnegative
                                                     parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-invertifnegative ).
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
        ENDLOOP.
        "endseries
        IF lo_chartb->ns_groupingval = zcl_excel_graph_bars=>c_groupingval_stacked.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_overlap
                                                            parent = lo_element3 ).
          lo_element4->set_attribute_ns( name  = 'val'
                                         value = '100' ).
        ENDIF.

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_gapwidth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartb->ns_gapwidthval ).

        "axes
        lo_el_rootchart = lo_element3.
        LOOP AT lo_chartb->axes INTO ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          CASE ls_ax-type.
            WHEN zcl_excel_graph_bars=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            WHEN zcl_excel_graph_bars=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            WHEN OTHERS.
          ENDCASE.
        ENDLOOP.
        "endaxes

      WHEN zcl_excel_drawing=>c_graph_pie.
        "----pie
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_piechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_varycolorsval ).

        "series
        LOOP AT lo_chartp->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
        ENDLOOP.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showbubblesizeval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showleaderlines
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_showleaderlinesval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_firstsliceang
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartp->ns_firstsliceangval ).
      WHEN zcl_excel_drawing=>c_graph_line.
        "----line
        lo_element3 = lo_document->create_simple_element( name = lc_xml_node_linechart
                                                     parent = lo_element2 ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_grouping
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_groupingval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_varycolors
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_varycolorsval ).

        "series
        LOOP AT lo_chartl->series INTO ls_serie.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ser
                                                     parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_idx
                                                     parent = lo_element4 ).
          IF ls_serie-idx IS NOT INITIAL.
            lv_str = ls_serie-idx.
          ELSE.
            lv_str = sy-tabix - 1.
          ENDIF.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_order
                                                     parent = lo_element4 ).
          lv_str = ls_serie-order.
          CONDENSE lv_str.
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = lv_str ).
          IF ls_serie-sername IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_tx
                                                      parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_v
                                                      parent = lo_element5 ).
            lo_element6->set_value( value = ls_serie-sername ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element4 ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_symbol
                                                     parent = lo_element5 ).
          lo_element6->set_attribute_ns( name  = 'val'
                                  value = ls_serie-symbol ).
          IF ls_serie-lbl IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_cat
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_strref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-lbl ).
          ENDIF.
          IF ls_serie-ref IS NOT INITIAL.
            lo_element5 = lo_document->create_simple_element( name = lc_xml_node_val
                                                       parent = lo_element4 ).
            lo_element6 = lo_document->create_simple_element( name = lc_xml_node_numref
                                                       parent = lo_element5 ).
            lo_element7 = lo_document->create_simple_element( name = lc_xml_node_f
                                                       parent = lo_element6 ).
            lo_element7->set_value( value = ls_serie-ref ).
          ENDIF.
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'val'
                                  value = ls_serie-smooth ).
        ENDLOOP.
        "endseries

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_dlbls
                                                     parent = lo_element3 ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showlegendkey
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showlegendkeyval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showval
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showvalval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showcatname
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showcatnameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showsername
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showsernameval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showpercent
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showpercentval ).
        lo_element5 = lo_document->create_simple_element( name = lc_xml_node_showbubblesize
                                                     parent = lo_element4 ).
        lo_element5->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_showbubblesizeval ).

        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_marker
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_markerval ).
        lo_element4 = lo_document->create_simple_element( name = lc_xml_node_smooth
                                                     parent = lo_element3 ).
        lo_element4->set_attribute_ns( name  = 'val'
                                  value = lo_chartl->ns_smoothval ).

        "axes
        lo_el_rootchart = lo_element3.
        LOOP AT lo_chartl->axes INTO ls_ax.
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_el_rootchart ).
          lo_element4->set_attribute_ns( name  = 'val'
                                  value = ls_ax-axid ).
          CASE ls_ax-type.
            WHEN zcl_excel_graph_line=>c_catax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_catax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_auto
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-auto ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lblalgn
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lblalgn ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lbloffset
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-lbloffset ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_nomultilvllbl
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-nomultilvllbl ).
            WHEN zcl_excel_graph_line=>c_valax.
              lo_element3 = lo_document->create_simple_element( name = lc_xml_node_valax
                                                     parent = lo_element2 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axid
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axid ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_scaling
                                                     parent = lo_element3 ).
              lo_element5 = lo_document->create_simple_element( name = lc_xml_node_orientation
                                                     parent = lo_element4 ).
              lo_element5->set_attribute_ns( name  = 'val'
                                             value = ls_ax-orientation ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_delete
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-delete ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_axpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-axpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majorgridlines
                                                     parent = lo_element3 ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_numfmt
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'formatCode'
                                             value = ls_ax-formatcode ).
              lo_element4->set_attribute_ns( name  = 'sourceLinked'
                                             value = ls_ax-sourcelinked ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_majortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-majortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_minortickmark
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-minortickmark ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_ticklblpos
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-ticklblpos ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossax
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossax ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crosses
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crosses ).
              lo_element4 = lo_document->create_simple_element( name = lc_xml_node_crossbetween
                                                     parent = lo_element3 ).
              lo_element4->set_attribute_ns( name  = 'val'
                                             value = ls_ax-crossbetween ).
            WHEN OTHERS.
          ENDCASE.
        ENDLOOP.
        "endaxes

      WHEN OTHERS.
    ENDCASE.

    "legend
    IF lo_chart->print_label EQ abap_true.
      lo_element2 = lo_document->create_simple_element( name = lc_xml_node_legend
                                                         parent = lo_element ).
      CASE io_drawing->graph_type.
        WHEN zcl_excel_drawing=>c_graph_bars.
          "----bar
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartb->ns_overlayval ).
        WHEN zcl_excel_drawing=>c_graph_line.
          "----line
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartl->ns_overlayval ).
        WHEN zcl_excel_drawing=>c_graph_pie.
          "----pie
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_legendpos
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_legendposval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_layout
                                                       parent = lo_element2 ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_overlay
                                                       parent = lo_element2 ).
          lo_element3->set_attribute_ns( name  = 'val'
                                    value = lo_chartp->ns_overlayval ).
          lo_element3 = lo_document->create_simple_element( name = lc_xml_node_txpr
                                                       parent = lo_element2 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_bodypr
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_lststyle
                                                       parent = lo_element3 ).
          lo_element4 = lo_document->create_simple_element( name = lc_xml_node_p
                                                       parent = lo_element3 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_ppr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'rtl'
                                    value = lo_chartp->ns_pprrtl ).
          lo_element6 = lo_document->create_simple_element( name = lc_xml_node_defrpr
                                                       parent = lo_element5 ).
          lo_element5 = lo_document->create_simple_element( name = lc_xml_node_endpararpr
                                                       parent = lo_element4 ).
          lo_element5->set_attribute_ns( name  = 'lang'
                                    value = lo_chartp->ns_endpararprlang ).
        WHEN OTHERS.
      ENDCASE.
    ENDIF.

    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_plotvisonly
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_plotvisonlyval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_dispblanksas
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_dispblanksasval ).
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_showdlblsovermax
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'val'
                                      value = lo_chart->ns_showdlblsovermaxval ).
    "---------------------------END OF CHART

    "printSettings
    lo_element = lo_document->create_simple_element( name = lc_xml_node_printsettings
                                                         parent = lo_element_root ).
    "headerFooter
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_headerfooter
                                                         parent = lo_element ).
    "pageMargins
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagemargins
                                                         parent = lo_element ).
    lo_element2->set_attribute_ns( name  = 'b'
                                      value = lo_chart->pagemargins-b ).
    lo_element2->set_attribute_ns( name  = 'l'
                                      value = lo_chart->pagemargins-l ).
    lo_element2->set_attribute_ns( name  = 'r'
                                      value = lo_chart->pagemargins-r ).
    lo_element2->set_attribute_ns( name  = 't'
                                      value = lo_chart->pagemargins-t ).
    lo_element2->set_attribute_ns( name  = 'header'
                                      value = lo_chart->pagemargins-header ).
    lo_element2->set_attribute_ns( name  = 'footer'
                                      value = lo_chart->pagemargins-footer ).
    "pageSetup
    lo_element2 = lo_document->create_simple_element( name = lc_xml_node_pagesetup
                                                         parent = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).
  ENDMETHOD.


  METHOD create_xl_comments.
** Constant node name
    CONSTANTS: lc_xml_node_comments    TYPE string VALUE 'comments',
               lc_xml_node_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               " authors
               lc_xml_node_author      TYPE string VALUE 'author',
               lc_xml_node_authors     TYPE string VALUE 'authors',
               " comments
               lc_xml_node_commentlist TYPE string VALUE 'commentList',
               lc_xml_node_comment     TYPE string VALUE 'comment',
               lc_xml_node_text        TYPE string VALUE 'text',
               lc_xml_node_r           TYPE string VALUE 'r',
               lc_xml_node_rpr         TYPE string VALUE 'rPr',
               lc_xml_node_b           TYPE string VALUE 'b',
               lc_xml_node_sz          TYPE string VALUE 'sz',
               lc_xml_node_color       TYPE string VALUE 'color',
               lc_xml_node_rfont       TYPE string VALUE 'rFont',
*             lc_xml_node_charset     TYPE string VALUE 'charset',
               lc_xml_node_family      TYPE string VALUE 'family',
               lc_xml_node_t           TYPE string VALUE 't',
               " comments attributes
               lc_xml_attr_ref         TYPE string VALUE 'ref',
               lc_xml_attr_authorid    TYPE string VALUE 'authorId',
               lc_xml_attr_val         TYPE string VALUE 'val',
               lc_xml_attr_indexed     TYPE string VALUE 'indexed',
               lc_xml_attr_xmlspacing  TYPE string VALUE 'xml:space'.


    DATA: lo_document            TYPE REF TO if_ixml_document,
          lo_element_root        TYPE REF TO if_ixml_element,
          lo_element_authors     TYPE REF TO if_ixml_element,
          lo_element_author      TYPE REF TO if_ixml_element,
          lo_element_commentlist TYPE REF TO if_ixml_element,
          lo_element_comment     TYPE REF TO if_ixml_element,
          lo_element_text        TYPE REF TO if_ixml_element,
          lo_element_r           TYPE REF TO if_ixml_element,
          lo_element_rpr         TYPE REF TO if_ixml_element,
          lo_element_b           TYPE REF TO if_ixml_element,
          lo_element_sz          TYPE REF TO if_ixml_element,
          lo_element_color       TYPE REF TO if_ixml_element,
          lo_element_rfont       TYPE REF TO if_ixml_element,
*       lo_element_charset     TYPE REF TO if_ixml_element,
          lo_element_family      TYPE REF TO if_ixml_element,
          lo_element_t           TYPE REF TO if_ixml_element,
          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_comments            TYPE REF TO zcl_excel_comments,
          lo_comment             TYPE REF TO zcl_excel_comment.
    DATA: lv_rel_id TYPE i,
          lv_author TYPE string.


**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_comments
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).

**********************************************************************
* STEP 4: Create authors
* TO-DO: management of several authors
    lo_element_authors = lo_document->create_simple_element( name   = lc_xml_node_authors
                                                             parent = lo_document ).

    lo_element_author  = lo_document->create_simple_element( name   = lc_xml_node_author
                                                             parent = lo_document ).
    lv_author = sy-uname.
    lo_element_author->set_value( lv_author ).

    lo_element_authors->append_child( new_child = lo_element_author ).
    lo_element_root->append_child( new_child = lo_element_authors ).

**********************************************************************
* STEP 5: Create comments

    lo_element_commentlist = lo_document->create_simple_element( name   = lc_xml_node_commentlist
                                                                 parent = lo_document ).

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_comment ?= lo_iterator->get_next( ).

      lo_element_comment = lo_document->create_simple_element( name   = lc_xml_node_comment
                                                               parent = lo_document ).
      lo_element_comment->set_attribute_ns( name  = lc_xml_attr_ref
                                            value = lo_comment->get_ref( ) ).
      lo_element_comment->set_attribute_ns( name  = lc_xml_attr_authorid
                                            value = '0' ).  " TO-DO

      lo_element_text = lo_document->create_simple_element( name   = lc_xml_node_text
                                                            parent = lo_document ).
      lo_element_r    = lo_document->create_simple_element( name   = lc_xml_node_r
                                                            parent = lo_document ).
      lo_element_rpr  = lo_document->create_simple_element( name   = lc_xml_node_rpr
                                                            parent = lo_document ).

      lo_element_b    = lo_document->create_simple_element( name   = lc_xml_node_b
                                                            parent = lo_document ).
      lo_element_rpr->append_child( new_child = lo_element_b ).

      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_sz     iv_attr_name = lc_xml_attr_val     iv_attr_value = '9' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_color  iv_attr_name = lc_xml_attr_indexed iv_attr_value = '81' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_rfont  iv_attr_name = lc_xml_attr_val     iv_attr_value = 'Tahoma' ).
      add_1_val_child_node( io_document = lo_document io_parent = lo_element_rpr iv_elem_name = lc_xml_node_family iv_attr_name = lc_xml_attr_val     iv_attr_value = '2' ).

      lo_element_r->append_child( new_child = lo_element_rpr ).

      lo_element_t    = lo_document->create_simple_element( name   = lc_xml_node_t
                                                            parent = lo_document ).
      lo_element_t->set_attribute_ns( name  = lc_xml_attr_xmlspacing
                                      value = 'preserve' ).
      lo_element_t->set_value( lo_comment->get_text( ) ).
      lo_element_r->append_child( new_child = lo_element_t ).

      lo_element_text->append_child( new_child = lo_element_r ).
      lo_element_comment->append_child( new_child = lo_element_text ).
      lo_element_commentlist->append_child( new_child = lo_element_comment ).
    ENDWHILE.

    lo_element_root->append_child( new_child = lo_element_commentlist ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_drawings.


** Constant node name
    CONSTANTS: lc_xml_node_wsdr   TYPE string VALUE 'xdr:wsDr',
               lc_xml_node_ns_xdr TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
               lc_xml_node_ns_a   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main'.

    DATA: lo_document           TYPE REF TO if_ixml_document,
          lo_element_root       TYPE REF TO if_ixml_element,
          lo_element_cellanchor TYPE REF TO if_ixml_element,
          lo_iterator           TYPE REF TO zcl_excel_collection_iterator,
          lo_drawings           TYPE REF TO zcl_excel_drawings,
          lo_drawing            TYPE REF TO zcl_excel_drawing.
    DATA: lv_rel_id            TYPE i.



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_wsdr
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:xdr'
                                       value = lc_xml_node_ns_xdr ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:a'
                                       value = lc_xml_node_ns_a ).

**********************************************************************
* STEP 4: Create drawings

    CLEAR: lv_rel_id.

    lo_drawings = io_worksheet->get_drawings( ).

    lo_iterator = lo_drawings->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      ADD 1 TO lv_rel_id.
      lo_element_cellanchor = me->create_xl_drawing_anchor(
              io_drawing    = lo_drawing
              io_document   = lo_document
              ip_index      = lv_rel_id ).

      lo_element_root->append_child( new_child = lo_element_cellanchor ).

    ENDWHILE.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_drawings_hdft_rels.

** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    DATA: lo_drawing      TYPE REF TO zcl_excel_drawing,
          lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lv_value        TYPE string,
          lv_relation_id  TYPE i,
          lt_temp         TYPE strtable,
          lt_drawings     TYPE zexcel_t_drawings.

    FIELD-SYMBOLS: <fs_temp>     TYPE sstrtable,
                   <fs_drawings> TYPE zexcel_s_drawings.


* BODY
**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

**********************************************************************


    lt_drawings = io_worksheet->get_header_footer_drawings( ).
    LOOP AT lt_drawings ASSIGNING <fs_drawings>. "Header or footer image exist
      ADD 1 TO lv_relation_id.
      lv_value = <fs_drawings>-drawing->get_index( ).
      READ TABLE lt_temp WITH KEY str = lv_value TRANSPORTING NO FIELDS.
      IF sy-subrc NE 0.
        APPEND INITIAL LINE TO lt_temp ASSIGNING <fs_temp>.
        <fs_temp>-row_index = sy-tabix.
        <fs_temp>-str = lv_value.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                           parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                      value = lc_xml_node_rid_image_tp ).

        lv_value = '../media/#'.
        REPLACE '#' IN lv_value WITH <fs_drawings>-drawing->get_media_name( ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                      value = lv_value ).
        lo_element_root->append_child( new_child = lo_element ).
      ENDIF.
    ENDLOOP.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.                    "create_xl_drawings_hdft_rels


  METHOD create_xl_drawings_rels.

** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_drawings     TYPE REF TO zcl_excel_drawings,
          lo_drawing      TYPE REF TO zcl_excel_drawing.

    DATA: lv_value   TYPE string,
          lv_counter TYPE i.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " Add sheet Relationship nodes here
    lv_counter = 0.
    lo_drawings = io_worksheet->get_drawings( ).
    lo_iterator = lo_drawings->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).
      ADD 1 TO lv_counter.

      lv_value = lv_counter.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                   parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).

      lv_value = lo_drawing->get_media_name( ).
      CASE lo_drawing->get_type( ).
        WHEN zcl_excel_drawing=>type_image.
          CONCATENATE '../media/' lv_value INTO lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_image_tp ).

        WHEN zcl_excel_drawing=>type_chart.
          CONCATENATE '../charts/' lv_value INTO lv_value.
          lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                        value = lc_xml_node_rid_chart_tp ).

      ENDCASE.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.


**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_drawings_vml.

    DATA:
      lo_xml_document TYPE REF TO cl_xml_document,
      ld_stream       TYPE string.


* INIT_RESULT
    CLEAR ep_content.


* BODY
    ld_stream = set_vml_string( ).

    CREATE OBJECT lo_xml_document.
    CALL METHOD lo_xml_document->parse_string
      EXPORTING
        stream = ld_stream.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = ld_stream
      IMPORTING
        buffer = ep_content
      EXCEPTIONS
        failed = 1
        OTHERS = 2.
    IF sy-subrc <> 0.
      CLEAR ep_content.
    ENDIF.


  ENDMETHOD.


  METHOD create_xl_drawings_vml_rels.

** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_image_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          lc_xml_node_rid_chart_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'.

    DATA: lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_drawing      TYPE REF TO zcl_excel_drawing,
          lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lv_value        TYPE string,
          lv_relation_id  TYPE i.


* BODY
**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes
    lv_relation_id = 0.
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).
      IF lo_drawing->get_type( ) = zcl_excel_drawing=>type_image_header_footer.
        ADD 1 TO lv_relation_id.
        lv_value = lv_relation_id.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                           parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_id
*                                    value = 'LOGO' ).
                                      value = lv_value ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                      value = lc_xml_node_rid_image_tp ).

        lv_value = '../media/#'.
        REPLACE '#' IN lv_value WITH lo_drawing->get_media_name( ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_target
*                                    value = '../media/LOGO.png' ).
                                      value = lv_value ).
        lo_element_root->append_child( new_child = lo_element ).
      ENDIF.

    ENDWHILE.



**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_drawing_anchor.

** Constant node name
    CONSTANTS: lc_xml_node_onecellanchor     TYPE string VALUE 'xdr:oneCellAnchor',
               lc_xml_node_twocellanchor     TYPE string VALUE 'xdr:twoCellAnchor',
               lc_xml_node_from              TYPE string VALUE 'xdr:from',
               lc_xml_node_to                TYPE string VALUE 'xdr:to',
               lc_xml_node_pic               TYPE string VALUE 'xdr:pic',
               lc_xml_node_ext               TYPE string VALUE 'xdr:ext',
               lc_xml_node_clientdata        TYPE string VALUE 'xdr:clientData',

               lc_xml_node_col               TYPE string VALUE 'xdr:col',
               lc_xml_node_coloff            TYPE string VALUE 'xdr:colOff',
               lc_xml_node_row               TYPE string VALUE 'xdr:row',
               lc_xml_node_rowoff            TYPE string VALUE 'xdr:rowOff',

               lc_xml_node_nvpicpr           TYPE string VALUE 'xdr:nvPicPr',
               lc_xml_node_cnvpr             TYPE string VALUE 'xdr:cNvPr',
               lc_xml_node_cnvpicpr          TYPE string VALUE 'xdr:cNvPicPr',
               lc_xml_node_piclocks          TYPE string VALUE 'a:picLocks',

               lc_xml_node_sppr              TYPE string VALUE 'xdr:spPr',
               lc_xml_node_apgeom            TYPE string VALUE 'a:prstGeom',
               lc_xml_node_aavlst            TYPE string VALUE 'a:avLst',

               lc_xml_node_graphicframe      TYPE string VALUE 'xdr:graphicFrame',
               lc_xml_node_nvgraphicframepr  TYPE string VALUE 'xdr:nvGraphicFramePr',
               lc_xml_node_cnvgraphicframepr TYPE string VALUE 'xdr:cNvGraphicFramePr',
               lc_xml_node_graphicframelocks TYPE string VALUE 'a:graphicFrameLocks',
               lc_xml_node_xfrm              TYPE string VALUE 'xdr:xfrm',
               lc_xml_node_aoff              TYPE string VALUE 'a:off',
               lc_xml_node_aext              TYPE string VALUE 'a:ext',
               lc_xml_node_agraphic          TYPE string VALUE 'a:graphic',
               lc_xml_node_agraphicdata      TYPE string VALUE 'a:graphicData',

               lc_xml_node_ns_c              TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
               lc_xml_node_cchart            TYPE string VALUE 'c:chart',

               lc_xml_node_blipfill          TYPE string VALUE 'xdr:blipFill',
               lc_xml_node_ablip             TYPE string VALUE 'a:blip',
               lc_xml_node_astretch          TYPE string VALUE 'a:stretch',
               lc_xml_node_ns_r              TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'.

    DATA: lo_element_graphicframe TYPE REF TO if_ixml_element,
          lo_element              TYPE REF TO if_ixml_element,
          lo_element2             TYPE REF TO if_ixml_element,
          lo_element3             TYPE REF TO if_ixml_element,
          lo_element_from         TYPE REF TO if_ixml_element,
          lo_element_to           TYPE REF TO if_ixml_element,
          lo_element_ext          TYPE REF TO if_ixml_element,
          lo_element_pic          TYPE REF TO if_ixml_element,
          lo_element_clientdata   TYPE REF TO if_ixml_element,
          ls_position             TYPE zexcel_drawing_position,
          lv_col                  TYPE string, " zexcel_cell_column,
          lv_row                  TYPE string, " zexcel_cell_row.
          lv_col_offset           TYPE string,
          lv_row_offset           TYPE string,
          lv_value                TYPE string.

    ls_position = io_drawing->get_position( ).

    IF ls_position-anchor = 'ONE'.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_onecellanchor
                                                                  parent = io_document ).
    ELSE.
      ep_anchor = io_document->create_simple_element( name   = lc_xml_node_twocellanchor
                                                                  parent = io_document ).
    ENDIF.

*   from cell ******************************
    lo_element_from = io_document->create_simple_element( name   = lc_xml_node_from
                                                          parent = io_document ).

    lv_col = ls_position-from-col.
    lv_row = ls_position-from-row.
    lv_col_offset = ls_position-from-col_offset.
    lv_row_offset = ls_position-from-row_offset.
    CONDENSE lv_col NO-GAPS.
    CONDENSE lv_row NO-GAPS.
    CONDENSE lv_col_offset NO-GAPS.
    CONDENSE lv_row_offset NO-GAPS.

    lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_col_offset ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row ).
    lo_element_from->append_child( new_child = lo_element ).

    lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                     parent = io_document ).
    lo_element->set_value( value = lv_row_offset ).
    lo_element_from->append_child( new_child = lo_element ).
    ep_anchor->append_child( new_child = lo_element_from ).

    IF ls_position-anchor = 'ONE'.

*   ext ******************************
      lo_element_ext = io_document->create_simple_element( name   = lc_xml_node_ext
                                                           parent = io_document ).

      lv_value = io_drawing->get_width_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cx'
                                     value = lv_value ).
      lv_value = io_drawing->get_height_emu_str( ).
      lo_element_ext->set_attribute_ns( name  = 'cy'
                                     value = lv_value ).
      ep_anchor->append_child( new_child = lo_element_ext ).

    ELSEIF ls_position-anchor = 'TWO'.

*   to cell ******************************
      lo_element_to = io_document->create_simple_element( name   = lc_xml_node_to
                                                          parent = io_document ).

      lv_col = ls_position-to-col.
      lv_row = ls_position-to-row.
      lv_col_offset = ls_position-to-col_offset.
      lv_row_offset = ls_position-to-row_offset.
      CONDENSE lv_col NO-GAPS.
      CONDENSE lv_row NO-GAPS.
      CONDENSE lv_col_offset NO-GAPS.
      CONDENSE lv_row_offset NO-GAPS.

      lo_element = io_document->create_simple_element( name = lc_xml_node_col
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_coloff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_col_offset ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_row
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row ).
      lo_element_to->append_child( new_child = lo_element ).

      lo_element = io_document->create_simple_element( name = lc_xml_node_rowoff
                                                       parent = io_document ).
      lo_element->set_value( value = lv_row_offset ).
      lo_element_to->append_child( new_child = lo_element ).
      ep_anchor->append_child( new_child = lo_element_to ).

    ENDIF.

    CASE io_drawing->get_type( ).
      WHEN zcl_excel_drawing=>type_image.
*     pic **********************************
        lo_element_pic = io_document->create_simple_element( name   = lc_xml_node_pic
                                                             parent = io_document ).
*     nvPicPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvpicpr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        CONDENSE lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).

*     cNvPicPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpicpr
                                                          parent = io_document ).

*     picLocks
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_piclocks
                                                          parent = io_document ).
        lo_element3->set_attribute_ns( name  = 'noChangeAspect'
                                       value = '1' ).

        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_pic->append_child( new_child = lo_element ).

*     blipFill
        lv_value = ip_index.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.

        lo_element  = io_document->create_simple_element( name = lc_xml_node_blipfill
                                                          parent = io_document ).
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_ablip
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element2->set_attribute_ns( name  = 'r:embed'
                                       value = lv_value ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element2  = io_document->create_simple_element( name = lc_xml_node_astretch
                                                          parent = io_document ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).

*     spPr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_sppr
                                                          parent = io_document ).

        lo_element2 = io_document->create_simple_element( name = lc_xml_node_apgeom
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'prst'
                                       value = 'rect' ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_aavlst
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).

        lo_element_pic->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_pic ).
      WHEN zcl_excel_drawing=>type_chart.
*     graphicFrame **********************************
        lo_element_graphicframe = io_document->create_simple_element( name   = lc_xml_node_graphicframe
                                                             parent = io_document ).
*     nvGraphicFramePr
        lo_element  = io_document->create_simple_element( name = lc_xml_node_nvgraphicframepr
                                                          parent = io_document ).
*     cNvPr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvpr
                                                          parent = io_document ).
        lv_value = sy-index.
        CONDENSE lv_value.
        lo_element2->set_attribute_ns( name  = 'id'
                                       value = lv_value ).
        lo_element2->set_attribute_ns( name  = 'name'
                                       value = io_drawing->title ).
        lo_element->append_child( new_child = lo_element2 ).
*     cNvGraphicFramePr
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_cnvgraphicframepr
                                                          parent = io_document ).
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_graphicframelocks
                                                          parent = io_document ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     xfrm
        lo_element  = io_document->create_simple_element( name = lc_xml_node_xfrm
                                                          parent = io_document ).
*     off
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aoff
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'y' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'x' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
*     ext
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_aext
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'cy' value = '0' ).
        lo_element2->set_attribute_ns( name  = 'cx' value = '0' ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).

*     graphic
        lo_element  = io_document->create_simple_element( name = lc_xml_node_agraphic
                                                          parent = io_document ).
*     graphicData
        lo_element2 = io_document->create_simple_element( name = lc_xml_node_agraphicdata
                                                          parent = io_document ).
        lo_element2->set_attribute_ns( name  = 'uri' value = lc_xml_node_ns_c ).

*     chart
        lo_element3 = io_document->create_simple_element( name = lc_xml_node_cchart
                                                          parent = io_document ).

        lo_element3->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_ns_r ).
        lo_element3->set_attribute_ns( name  = 'xmlns:c'
                                       value = lc_xml_node_ns_c ).

        lv_value = ip_index.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element3->set_attribute_ns( name  = 'r:id'
                                       value = lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
        lo_element->append_child( new_child = lo_element2 ).
        lo_element_graphicframe->append_child( new_child = lo_element ).
        ep_anchor->append_child( new_child = lo_element_graphicframe ).

    ENDCASE.

*   client data ***************************
    lo_element_clientdata = io_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                parent = io_document ).
    ep_anchor->append_child( new_child = lo_element_clientdata ).

  ENDMETHOD.


  METHOD create_xl_drawing_for_comments.
** Constant node name
    CONSTANTS: lc_xml_node_xml             TYPE string VALUE 'xml',
               lc_xml_node_ns_v            TYPE string VALUE 'urn:schemas-microsoft-com:vml',
               lc_xml_node_ns_o            TYPE string VALUE 'urn:schemas-microsoft-com:office:office',
               lc_xml_node_ns_x            TYPE string VALUE 'urn:schemas-microsoft-com:office:excel',
               " shapelayout
               lc_xml_node_shapelayout     TYPE string VALUE 'o:shapelayout',
               lc_xml_node_idmap           TYPE string VALUE 'o:idmap',
               " shapetype
               lc_xml_node_shapetype       TYPE string VALUE 'v:shapetype',
               lc_xml_node_stroke          TYPE string VALUE 'v:stroke',
               lc_xml_node_path            TYPE string VALUE 'v:path',
               " shape
               lc_xml_node_shape           TYPE string VALUE 'v:shape',
               lc_xml_node_fill            TYPE string VALUE 'v:fill',
               lc_xml_node_shadow          TYPE string VALUE 'v:shadow',
               lc_xml_node_textbox         TYPE string VALUE 'v:textbox',
               lc_xml_node_div             TYPE string VALUE 'div',
               lc_xml_node_clientdata      TYPE string VALUE 'x:ClientData',
               lc_xml_node_movewithcells   TYPE string VALUE 'x:MoveWithCells',
               lc_xml_node_sizewithcells   TYPE string VALUE 'x:SizeWithCells',
               lc_xml_node_anchor          TYPE string VALUE 'x:Anchor',
               lc_xml_node_autofill        TYPE string VALUE 'x:AutoFill',
               lc_xml_node_row             TYPE string VALUE 'x:Row',
               lc_xml_node_column          TYPE string VALUE 'x:Column',
               " attributes,
               lc_xml_attr_vext            TYPE string VALUE 'v:ext',
               lc_xml_attr_data            TYPE string VALUE 'data',
               lc_xml_attr_id              TYPE string VALUE 'id',
               lc_xml_attr_coordsize       TYPE string VALUE 'coordsize',
               lc_xml_attr_ospt            TYPE string VALUE 'o:spt',
               lc_xml_attr_joinstyle       TYPE string VALUE 'joinstyle',
               lc_xml_attr_path            TYPE string VALUE 'path',
               lc_xml_attr_gradientshapeok TYPE string VALUE 'gradientshapeok',
               lc_xml_attr_oconnecttype    TYPE string VALUE 'o:connecttype',
               lc_xml_attr_type            TYPE string VALUE 'type',
               lc_xml_attr_style           TYPE string VALUE 'style',
               lc_xml_attr_fillcolor       TYPE string VALUE 'fillcolor',
               lc_xml_attr_oinsetmode      TYPE string VALUE 'o:insetmode',
               lc_xml_attr_color           TYPE string VALUE 'color',
               lc_xml_attr_color2          TYPE string VALUE 'color2',
               lc_xml_attr_on              TYPE string VALUE 'on',
               lc_xml_attr_obscured        TYPE string VALUE 'obscured',
               lc_xml_attr_objecttype      TYPE string VALUE 'ObjectType',
               " attributes values
               lc_xml_attr_val_edit        TYPE string VALUE 'edit',
               lc_xml_attr_val_rect        TYPE string VALUE 'rect',
               lc_xml_attr_val_t           TYPE string VALUE 't',
               lc_xml_attr_val_miter       TYPE string VALUE 'miter',
               lc_xml_attr_val_auto        TYPE string VALUE 'auto',
               lc_xml_attr_val_black       TYPE string VALUE 'black',
               lc_xml_attr_val_none        TYPE string VALUE 'none',
               lc_xml_attr_val_msodir      TYPE string VALUE 'mso-direction-alt:auto',
               lc_xml_attr_val_note        TYPE string VALUE 'Note'.


    DATA: lo_document              TYPE REF TO if_ixml_document,
          lo_element_root          TYPE REF TO if_ixml_element,
          "shapelayout
          lo_element_shapelayout   TYPE REF TO if_ixml_element,
          lo_element_idmap         TYPE REF TO if_ixml_element,
          "shapetype
          lo_element_shapetype     TYPE REF TO if_ixml_element,
          lo_element_stroke        TYPE REF TO if_ixml_element,
          lo_element_path          TYPE REF TO if_ixml_element,
          "shape
          lo_element_shape         TYPE REF TO if_ixml_element,
          lo_element_fill          TYPE REF TO if_ixml_element,
          lo_element_shadow        TYPE REF TO if_ixml_element,
          lo_element_textbox       TYPE REF TO if_ixml_element,
          lo_element_div           TYPE REF TO if_ixml_element,
          lo_element_clientdata    TYPE REF TO if_ixml_element,
          lo_element_movewithcells TYPE REF TO if_ixml_element,
          lo_element_sizewithcells TYPE REF TO if_ixml_element,
          lo_element_anchor        TYPE REF TO if_ixml_element,
          lo_element_autofill      TYPE REF TO if_ixml_element,
          lo_element_row           TYPE REF TO if_ixml_element,
          lo_element_column        TYPE REF TO if_ixml_element,
          lo_iterator              TYPE REF TO zcl_excel_collection_iterator,
          lo_comments              TYPE REF TO zcl_excel_comments,
          lo_comment               TYPE REF TO zcl_excel_comment,
          lv_row                   TYPE zexcel_cell_row,
          lv_str_column            TYPE zexcel_cell_column_alpha,
          lv_column                TYPE zexcel_cell_column,
          lv_index                 TYPE i,
          lv_attr_id_index         TYPE i,
          lv_attr_id               TYPE string,
          lv_int_value             TYPE i,
          lv_int_value_string      TYPE string.
    DATA: lv_rel_id            TYPE i.


**********************************************************************
* STEP 1: Create XML document
    lo_document = me->ixml->create_document( ).

***********************************************************************
* STEP 2: Create main node relationships
    lo_element_root = lo_document->create_simple_element( name   = lc_xml_node_xml
                                                          parent = lo_document ).
    lo_element_root->set_attribute_ns( : name  = 'xmlns:v'  value = lc_xml_node_ns_v ),
                                         name  = 'xmlns:o'  value = lc_xml_node_ns_o ),
                                         name  = 'xmlns:x'  value = lc_xml_node_ns_x ).

**********************************************************************
* STEP 3: Create o:shapeLayout
* TO-DO: management of several authors
    lo_element_shapelayout = lo_document->create_simple_element( name   = lc_xml_node_shapelayout
                                                                 parent = lo_document ).

    lo_element_shapelayout->set_attribute_ns( name  = lc_xml_attr_vext
                                              value = lc_xml_attr_val_edit ).

    lo_element_idmap = lo_document->create_simple_element( name   = lc_xml_node_idmap
                                                           parent = lo_document ).
    lo_element_idmap->set_attribute_ns( : name  = lc_xml_attr_vext  value = lc_xml_attr_val_edit ),
                                          name  = lc_xml_attr_data  value = '1' ).

    lo_element_shapelayout->append_child( new_child = lo_element_idmap ).

    lo_element_root->append_child( new_child = lo_element_shapelayout ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_element_shapetype = lo_document->create_simple_element( name   = lc_xml_node_shapetype
                                                               parent = lo_document ).

    lo_element_shapetype->set_attribute_ns( : name  = lc_xml_attr_id         value = '_x0000_t202' ),
                                              name  = lc_xml_attr_coordsize  value = '21600,21600' ),
                                              name  = lc_xml_attr_ospt       value = '202' ),
                                              name  = lc_xml_attr_path       value = 'm,l,21600r21600,l21600,xe' ).

    lo_element_stroke = lo_document->create_simple_element( name   = lc_xml_node_stroke
                                                            parent = lo_document ).
    lo_element_stroke->set_attribute_ns( name  = lc_xml_attr_joinstyle       value = lc_xml_attr_val_miter ).

    lo_element_path   = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
    lo_element_path->set_attribute_ns( : name  = lc_xml_attr_gradientshapeok value = lc_xml_attr_val_t ),
                                         name  = lc_xml_attr_oconnecttype    value = lc_xml_attr_val_rect ).

    lo_element_shapetype->append_child( : new_child = lo_element_stroke ),
                                          new_child = lo_element_path ).

    lo_element_root->append_child( new_child = lo_element_shapetype ).

**********************************************************************
* STEP 4: Create v:shapetype

    lo_comments = io_worksheet->get_comments( ).

    lo_iterator = lo_comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lv_index = sy-index.
      lo_comment ?= lo_iterator->get_next( ).

      zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING i_columnrow = lo_comment->get_ref( )
                                                        IMPORTING e_column = lv_str_column
                                                                  e_row    = lv_row ).
      lv_column = zcl_excel_common=>convert_column2int( lv_str_column ).

      lo_element_shape = lo_document->create_simple_element( name   = lc_xml_node_shape
                                                             parent = lo_document ).

      lv_attr_id_index = 1024 + lv_index.
      lv_attr_id = lv_attr_id_index.
      CONCATENATE '_x0000_s' lv_attr_id INTO lv_attr_id.
      lo_element_shape->set_attribute_ns( : name  = lc_xml_attr_id          value = lv_attr_id ),
                                            name  = lc_xml_attr_type        value = '#_x0000_t202' ),
                                            name  = lc_xml_attr_style       value = 'size:auto;width:auto;height:auto;position:absolute;margin-left:117pt;margin-top:172.5pt;z-index:1;visibility:hidden' ),
                                            name  = lc_xml_attr_fillcolor   value = '#ffffe1' ),
                                            name  = lc_xml_attr_oinsetmode  value = lc_xml_attr_val_auto ).

      " Fill
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).
      lo_element_fill->set_attribute_ns( name = lc_xml_attr_color2  value = '#ffffe1' ).
      lo_element_shape->append_child( new_child = lo_element_fill ).
      " Shadow
      lo_element_shadow = lo_document->create_simple_element( name   = lc_xml_node_shadow
                                                              parent = lo_document ).
      lo_element_shadow->set_attribute_ns( : name = lc_xml_attr_on        value = lc_xml_attr_val_t ),
                                             name = lc_xml_attr_color     value = lc_xml_attr_val_black ),
                                             name = lc_xml_attr_obscured  value = lc_xml_attr_val_t ).
      lo_element_shape->append_child( new_child = lo_element_shadow ).
      " Path
      lo_element_path = lo_document->create_simple_element( name   = lc_xml_node_path
                                                            parent = lo_document ).
      lo_element_path->set_attribute_ns( name = lc_xml_attr_oconnecttype  value = lc_xml_attr_val_none ).
      lo_element_shape->append_child( new_child = lo_element_path ).
      " Textbox
      lo_element_textbox = lo_document->create_simple_element( name   = lc_xml_node_textbox
                                                               parent = lo_document ).
      lo_element_textbox->set_attribute_ns( name = lc_xml_attr_style  value = lc_xml_attr_val_msodir ).
      lo_element_div = lo_document->create_simple_element( name   = lc_xml_node_div
                                                           parent = lo_document ).
      lo_element_div->set_attribute_ns( name = lc_xml_attr_style  value = 'text-align:left' ).
      lo_element_textbox->append_child( new_child = lo_element_div ).
      lo_element_shape->append_child( new_child = lo_element_textbox ).
      " ClientData
      lo_element_clientdata = lo_document->create_simple_element( name   = lc_xml_node_clientdata
                                                                  parent = lo_document ).
      lo_element_clientdata->set_attribute_ns( name = lc_xml_attr_objecttype  value = lc_xml_attr_val_note ).
      lo_element_movewithcells = lo_document->create_simple_element( name   = lc_xml_node_movewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_movewithcells ).
      lo_element_sizewithcells = lo_document->create_simple_element( name   = lc_xml_node_sizewithcells
                                                                     parent = lo_document ).
      lo_element_clientdata->append_child( new_child = lo_element_sizewithcells ).
      lo_element_anchor = lo_document->create_simple_element( name   = lc_xml_node_anchor
                                                              parent = lo_document ).
      lo_element_anchor->set_value( '2, 15, 11, 10, 4, 31, 15, 9' ).
      lo_element_clientdata->append_child( new_child = lo_element_anchor ).
      lo_element_autofill = lo_document->create_simple_element( name   = lc_xml_node_autofill
                                                                parent = lo_document ).
      lo_element_autofill->set_value( 'False' ).
      lo_element_clientdata->append_child( new_child = lo_element_autofill ).
      lo_element_row = lo_document->create_simple_element( name   = lc_xml_node_row
                                                           parent = lo_document ).
      lv_int_value = lv_row - 1.
      lv_int_value_string = lv_int_value.
      lo_element_row->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_row ).
      lo_element_column = lo_document->create_simple_element( name   = lc_xml_node_column
                                                                parent = lo_document ).
      lv_int_value = lv_column - 1.
      lv_int_value_string = lv_int_value.
      lo_element_column->set_value( lv_int_value_string ).
      lo_element_clientdata->append_child( new_child = lo_element_column ).

      lo_element_shape->append_child( new_child = lo_element_clientdata ).

      lo_element_root->append_child( new_child = lo_element_shape ).
    ENDWHILE.

**********************************************************************
* STEP 6: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_drawing_for_hdft_im.


    DATA:
      ld_1            TYPE string,
      ld_2            TYPE string,
      ld_3            TYPE string,
      ld_4            TYPE string,
      ld_5            TYPE string,
      ld_7            TYPE string,

      ls_odd_header   TYPE zexcel_s_worksheet_head_foot,
      ls_odd_footer   TYPE zexcel_s_worksheet_head_foot,
      ls_even_header  TYPE zexcel_s_worksheet_head_foot,
      ls_even_footer  TYPE zexcel_s_worksheet_head_foot,
      lv_content      TYPE string,
      lo_xml_document TYPE REF TO cl_xml_document.


* INIT_RESULT
    CLEAR ep_content.


* BODY
    ld_1 = '<xml xmlns:v="urn:schemas-microsoft-com:vml"  xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'.
    ld_2 = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>'.
    ld_3 = '<v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'.
    ld_4 = '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'.


    CONCATENATE ld_1
                ld_2
                ld_3
                ld_4
         INTO lv_content.

    io_worksheet->sheet_setup->get_header_footer( IMPORTING ep_odd_header = ls_odd_header
                                                            ep_odd_footer = ls_odd_footer
                                                            ep_even_header = ls_even_header
                                                            ep_even_footer = ls_even_footer ).

    ld_5 = me->set_vml_shape_header( ls_odd_header ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_header( ls_even_header ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_footer( ls_odd_footer ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_footer( ls_even_footer ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.

    ld_7 = '</xml>'.

    CONCATENATE lv_content
                ld_7
           INTO lv_content.

    CREATE OBJECT lo_xml_document.
    CALL METHOD lo_xml_document->parse_string
      EXPORTING
        stream = lv_content.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = lv_content
      IMPORTING
        buffer = ep_content
      EXCEPTIONS
        failed = 1
        OTHERS = 2.
    IF sy-subrc <> 0.
      CLEAR ep_content.
    ENDIF.

  ENDMETHOD.


  METHOD create_xl_relationships.


** Constant node name
    DATA: lc_xml_node_relationships TYPE string VALUE 'Relationships',
          lc_xml_node_relationship  TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id            TYPE string VALUE 'Id',
          lc_xml_attr_type          TYPE string VALUE 'Type',
          lc_xml_attr_target        TYPE string VALUE 'Target',
          " Node namespace
          lc_xml_node_rels_ns       TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          " Node id
          lc_xml_node_ridx_id       TYPE string VALUE 'rId#',
          " Node type
          lc_xml_node_rid_sheet_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          lc_xml_node_rid_theme_tp  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
          lc_xml_node_rid_styles_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
          lc_xml_node_rid_shared_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
          " Node target
          lc_xml_node_ridx_tg       TYPE string VALUE 'worksheets/sheet#.xml',
          lc_xml_node_rid_shared_tg TYPE string VALUE 'sharedStrings.xml',
          lc_xml_node_rid_styles_tg TYPE string VALUE 'styles.xml',
          lc_xml_node_rid_theme_tg  TYPE string VALUE 'theme/theme1.xml'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element.

    DATA: lv_xml_node_ridx_tg TYPE string,
          lv_xml_node_ridx_id TYPE string,
          lv_size             TYPE i,
          lv_syindex          TYPE string.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lv_size = excel->get_worksheets_size( ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
    parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
    value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
    value = lc_xml_node_rid_theme_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
    value = lc_xml_node_rid_theme_tg ).
    lo_element_root->append_child( new_child = lo_element ).


    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    lv_size = lv_size + 1.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_styles_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_styles_tg ).
    lo_element_root->append_child( new_child = lo_element ).



    lv_size = excel->get_worksheets_size( ).

    DO lv_size TIMES.
      " Relationship node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
      parent = lo_document ).
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      lv_xml_node_ridx_tg = lc_xml_node_ridx_tg.
      lv_syindex = sy-index.
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_tg WITH lv_syindex.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
      value = lv_xml_node_ridx_id ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
      value = lc_xml_node_rid_sheet_tp ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
      value = lv_xml_node_ridx_tg ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDDO.

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    ADD 3 TO lv_size.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_shared_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_shared_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_sharedstrings.


** Constant node name
    DATA: lc_xml_node_sst         TYPE string VALUE 'sst',
          lc_xml_node_si          TYPE string VALUE 'si',
          lc_xml_node_t           TYPE string VALUE 't',
          lc_xml_node_r           TYPE string VALUE 'r',
          lc_xml_node_rpr         TYPE string VALUE 'rPr',
          " Node attributes
          lc_xml_attr_count       TYPE string VALUE 'count',
          lc_xml_attr_uniquecount TYPE string VALUE 'uniqueCount',
          " Node namespace
          lc_xml_node_ns          TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_sub_element  TYPE REF TO if_ixml_element,
          lo_sub2_element TYPE REF TO if_ixml_element,
          lo_font_element TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet    TYPE REF TO zcl_excel_worksheet.

    DATA: lt_cell_data       TYPE zexcel_t_cell_data_unsorted,
          lt_cell_data_rtf   TYPE zexcel_t_cell_data_unsorted,
          lv_value           TYPE string,
          ls_shared_string   TYPE zexcel_s_shared_string,
          lv_count_str       TYPE string,
          lv_uniquecount_str TYPE string,
          lv_sytabix         TYPE i,
          lv_count           TYPE i,
          lv_uniquecount     TYPE i.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                   <fs_rtf>           TYPE zexcel_s_rtf,
                   <fs_sheet_string>  TYPE zexcel_s_shared_string.

**********************************************************************
* STEP 1: Collect strings from each worksheet
    lo_iterator = excel->get_worksheets_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      APPEND LINES OF lo_worksheet->sheet_content TO lt_cell_data.
    ENDWHILE.

    DELETE lt_cell_data WHERE cell_formula IS NOT INITIAL. " delete formula content

    DESCRIBE TABLE lt_cell_data LINES lv_count.
    lv_count_str = lv_count.

    " separating plain and rich text format strings
    lt_cell_data_rtf = lt_cell_data.
    DELETE lt_cell_data WHERE rtf_tab IS NOT INITIAL.
    DELETE lt_cell_data_rtf WHERE rtf_tab IS INITIAL.

    SHIFT lv_count_str RIGHT DELETING TRAILING space.
    SHIFT lv_count_str LEFT DELETING LEADING space.

    SORT lt_cell_data BY cell_value data_type.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data COMPARING cell_value data_type.

    " leave unique rich text format strings
    SORT lt_cell_data_rtf BY cell_value rtf_tab.
    DELETE ADJACENT DUPLICATES FROM lt_cell_data_rtf COMPARING cell_value rtf_tab.
    " merge into single list
    APPEND LINES OF lt_cell_data_rtf TO lt_cell_data.
    SORT lt_cell_data BY cell_value rtf_tab.
    FREE lt_cell_data_rtf.

    DESCRIBE TABLE lt_cell_data LINES lv_uniquecount.
    lv_uniquecount_str = lv_uniquecount.

    SHIFT lv_uniquecount_str RIGHT DELETING TRAILING space.
    SHIFT lv_uniquecount_str LEFT DELETING LEADING space.

    CLEAR lv_count.
    LOOP AT lt_cell_data ASSIGNING <fs_sheet_content> WHERE data_type = 's'.
      lv_sytabix = lv_count.
      ls_shared_string-string_no = lv_sytabix.
      ls_shared_string-string_value = <fs_sheet_content>-cell_value.
      ls_shared_string-string_type = <fs_sheet_content>-data_type.
      ls_shared_string-rtf_tab = <fs_sheet_content>-rtf_tab.
      INSERT ls_shared_string INTO TABLE shared_strings.
      ADD 1 TO lv_count.
    ENDLOOP.


**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_sst
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_count
                                       value = lv_count_str ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_uniquecount
                                       value = lv_uniquecount_str ).

**********************************************************************
* STEP 4: Create subnode
    LOOP AT shared_strings ASSIGNING <fs_sheet_string>.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_si
                                                       parent = lo_document ).
      IF <fs_sheet_string>-rtf_tab IS INITIAL.
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                             parent = lo_document ).
        IF boolc( contains( val = <fs_sheet_string>-string_value start = ` ` ) ) = abap_true
              OR boolc( contains( val = <fs_sheet_string>-string_value end = ` ` ) ) = abap_true.
          lo_sub_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
        ENDIF.
        lo_sub_element->set_value( value = <fs_sheet_string>-string_value ).
      ELSE.
        LOOP AT <fs_sheet_string>-rtf_tab ASSIGNING <fs_rtf>.
          lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_r
                                                               parent = lo_element ).
          TRY.
              lv_value = substring( val = <fs_sheet_string>-string_value
                                    off = <fs_rtf>-offset
                                    len = <fs_rtf>-length ).
            CATCH cx_sy_range_out_of_bounds.
              EXIT.
          ENDTRY.
          IF <fs_rtf>-font IS NOT INITIAL.
            lo_font_element = lo_document->create_simple_element( name   = lc_xml_node_rpr
                                                                  parent = lo_sub_element ).
            create_xl_styles_font_node( io_document = lo_document
                                        io_parent   = lo_font_element
                                        is_font     = <fs_rtf>-font
                                        iv_use_rtf  = abap_true ).
          ENDIF.
          lo_sub2_element = lo_document->create_simple_element( name   = lc_xml_node_t
                                                              parent = lo_sub_element ).
          IF boolc( contains( val = lv_value start = ` ` ) ) = abap_true
                OR boolc( contains( val = lv_value end = ` ` ) ) = abap_true.
            lo_sub2_element->set_attribute( name = 'space' namespace = 'xml' value = 'preserve' ).
          ENDIF.
          lo_sub2_element->set_value( lv_value ).
        ENDLOOP.
      ENDIF.
      lo_element->append_child( new_child = lo_sub_element ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDLOOP.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_sheet.
*--------------------------------------------------------------------*
* issue #330   - Adding ColorScale conditional formatting
*              - Ivan Femia,                                2014-08-25
*--------------------------------------------------------------------*

    TYPES: BEGIN OF colors,
             colorrgb TYPE zexcel_color,
           END OF colors.

*--------------------------------------------------------------------*
* issue #237   - Error writing column-style
*              - Stefan Schmoecker,                          2012-11-01
*--------------------------------------------------------------------*

    TYPES: BEGIN OF cfvo,
             value TYPE zexcel_conditional_value,
             type  TYPE zexcel_conditional_type,
           END OF cfvo.

*--------------------------------------------------------------------*
* issue #220 - If cell in tables-area don't use default from row or column or sheet - Declarations 1 - start
*--------------------------------------------------------------------*
    TYPES: BEGIN OF lty_table_area,
             left   TYPE i,
             right  TYPE i,
             top    TYPE i,
             bottom TYPE i,
           END OF lty_table_area.
*--------------------------------------------------------------------*
* issue #220 - If cell in tables-area don't use default from row or column or sheet - Declarations 1 - end
*--------------------------------------------------------------------*

    TYPES: BEGIN OF ty_condformating_range,
             dimension_range     TYPE string,
             condformatting_node TYPE REF TO if_ixml_element,
           END OF ty_condformating_range,
           ty_condformating_ranges TYPE STANDARD TABLE OF ty_condformating_range.

** Constant node name
    DATA: lc_xml_node_worksheet          TYPE string VALUE 'worksheet',
          lc_xml_node_sheetpr            TYPE string VALUE 'sheetPr',
          lc_xml_node_tabcolor           TYPE string VALUE 'tabColor',
          lc_xml_node_outlinepr          TYPE string VALUE 'outlinePr',
          lc_xml_node_dimension          TYPE string VALUE 'dimension',
          lc_xml_node_sheetviews         TYPE string VALUE 'sheetViews',
          lc_xml_node_sheetview          TYPE string VALUE 'sheetView',
          lc_xml_node_selection          TYPE string VALUE 'selection',
          lc_xml_node_pane               TYPE string VALUE 'pane',
          lc_xml_node_sheetformatpr      TYPE string VALUE 'sheetFormatPr',
          lc_xml_node_cols               TYPE string VALUE 'cols',
          lc_xml_node_col                TYPE string VALUE 'col',
          lc_xml_node_sheetprotection    TYPE string VALUE 'sheetProtection',
          lc_xml_node_pagemargins        TYPE string VALUE 'pageMargins',
          lc_xml_node_pagesetup          TYPE string VALUE 'pageSetup',
          lc_xml_node_pagesetuppr        TYPE string VALUE 'pageSetUpPr',
          lc_xml_node_condformatting     TYPE string VALUE 'conditionalFormatting',
          lc_xml_node_cfrule             TYPE string VALUE 'cfRule',
          lc_xml_node_color              TYPE string VALUE 'color',      " Databar by Albert Lladanosa
          lc_xml_node_databar            TYPE string VALUE 'dataBar',    " Databar by Albert Lladanosa
          lc_xml_node_colorscale         TYPE string VALUE 'colorScale',
          lc_xml_node_iconset            TYPE string VALUE 'iconSet',
          lc_xml_node_cfvo               TYPE string VALUE 'cfvo',
          lc_xml_node_formula            TYPE string VALUE 'formula',
          lc_xml_node_datavalidations    TYPE string VALUE 'dataValidations',
          lc_xml_node_datavalidation     TYPE string VALUE 'dataValidation',
          lc_xml_node_formula1           TYPE string VALUE 'formula1',
          lc_xml_node_formula2           TYPE string VALUE 'formula2',
          lc_xml_node_mergecell          TYPE string VALUE 'mergeCell',
          lc_xml_node_mergecells         TYPE string VALUE 'mergeCells',
          lc_xml_node_drawing            TYPE string VALUE 'drawing',
          lc_xml_node_drawing_for_cmt    TYPE string VALUE 'legacyDrawing',

**********************************************************************
          lc_xml_node_drawing_for_hd_ft  TYPE string VALUE 'legacyDrawingHF',
**********************************************************************


          lc_xml_node_headerfooter       TYPE string VALUE 'headerFooter',
          lc_xml_node_oddheader          TYPE string VALUE 'oddHeader',
          lc_xml_node_oddfooter          TYPE string VALUE 'oddFooter',
          lc_xml_node_evenheader         TYPE string VALUE 'evenHeader',
          lc_xml_node_evenfooter         TYPE string VALUE 'evenFooter',
          lc_xml_node_autofilter         TYPE string VALUE 'autoFilter',
          lc_xml_node_filtercolumn       TYPE string VALUE 'filterColumn',
          lc_xml_node_filters            TYPE string VALUE 'filters',
          lc_xml_node_filter             TYPE string VALUE 'filter',
          " Node attributes
          lc_xml_attr_ref                TYPE string VALUE 'ref',
          lc_xml_attr_summarybelow       TYPE string VALUE 'summaryBelow',
          lc_xml_attr_summaryright       TYPE string VALUE 'summaryRight',
          lc_xml_attr_tabselected        TYPE string VALUE 'tabSelected',
          lc_xml_attr_showzeros          TYPE string VALUE 'showZeros',
          lc_xml_attr_zoomscale          TYPE string VALUE 'zoomScale',
          lc_xml_attr_zoomscalenormal    TYPE string VALUE 'zoomScaleNormal',
          lc_xml_attr_zoomscalepageview  TYPE string VALUE 'zoomScalePageLayoutView',
          lc_xml_attr_zoomscalesheetview TYPE string VALUE 'zoomScaleSheetLayoutView',
          lc_xml_attr_workbookviewid     TYPE string VALUE 'workbookViewId',
          lc_xml_attr_showgridlines      TYPE string VALUE 'showGridLines',
          lc_xml_attr_gridlines          TYPE string VALUE 'gridLines',
          lc_xml_attr_showrowcolheaders  TYPE string VALUE 'showRowColHeaders',
          lc_xml_attr_activecell         TYPE string VALUE 'activeCell',
          lc_xml_attr_sqref              TYPE string VALUE 'sqref',
          lc_xml_attr_min                TYPE string VALUE 'min',
          lc_xml_attr_max                TYPE string VALUE 'max',
          lc_xml_attr_hidden             TYPE string VALUE 'hidden',
          lc_xml_attr_width              TYPE string VALUE 'width',
          lc_xml_attr_defaultwidth       TYPE string VALUE '9.10',
          lc_xml_attr_style              TYPE string VALUE 'style',
          lc_xml_attr_true               TYPE string VALUE 'true',
          lc_xml_attr_bestfit            TYPE string VALUE 'bestFit',
          lc_xml_attr_customheight       TYPE string VALUE 'customHeight',
          lc_xml_attr_customwidth        TYPE string VALUE 'customWidth',
          lc_xml_attr_collapsed          TYPE string VALUE 'collapsed',
          lc_xml_attr_defaultrowheight   TYPE string VALUE 'defaultRowHeight',
          lc_xml_attr_defaultcolwidth    TYPE string VALUE 'defaultColWidth',
          lc_xml_attr_outlinelevelrow    TYPE string VALUE 'x14ac:outlineLevelRow',
          lc_xml_attr_outlinelevelcol    TYPE string VALUE 'x14ac:outlineLevelCol',
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
          lc_xml_attr_left               TYPE string VALUE 'left',
          lc_xml_attr_right              TYPE string VALUE 'right',
          lc_xml_attr_top                TYPE string VALUE 'top',
          lc_xml_attr_bottom             TYPE string VALUE 'bottom',
          lc_xml_attr_header             TYPE string VALUE 'header',
          lc_xml_attr_footer             TYPE string VALUE 'footer',
          lc_xml_attr_type               TYPE string VALUE 'type',
          lc_xml_attr_iconset            TYPE string VALUE 'iconSet',
          lc_xml_attr_showvalue          TYPE string VALUE 'showValue',
          lc_xml_attr_val                TYPE string VALUE 'val',
          lc_xml_attr_dxfid              TYPE string VALUE 'dxfId',
          lc_xml_attr_priority           TYPE string VALUE 'priority',
          lc_xml_attr_operator           TYPE string VALUE 'operator',
          lc_xml_attr_text               TYPE string VALUE 'text',
          lc_xml_attr_notcontainstext    TYPE string VALUE 'notContainsText',
          lc_xml_attr_allowblank         TYPE string VALUE 'allowBlank',
          lc_xml_attr_showinputmessage   TYPE string VALUE 'showInputMessage',
          lc_xml_attr_showerrormessage   TYPE string VALUE 'showErrorMessage',
          lc_xml_attr_showdropdown       TYPE string VALUE 'ShowDropDown', " 'showDropDown' does not work
          lc_xml_attr_errortitle         TYPE string VALUE 'errorTitle',
          lc_xml_attr_error              TYPE string VALUE 'error',
          lc_xml_attr_errorstyle         TYPE string VALUE 'errorStyle',
          lc_xml_attr_prompttitle        TYPE string VALUE 'promptTitle',
          lc_xml_attr_prompt             TYPE string VALUE 'prompt',
          lc_xml_attr_count              TYPE string VALUE 'count',
          lc_xml_attr_blackandwhite      TYPE string VALUE 'blackAndWhite',
          lc_xml_attr_cellcomments       TYPE string VALUE 'cellComments',
          lc_xml_attr_copies             TYPE string VALUE 'copies',
          lc_xml_attr_draft              TYPE string VALUE 'draft',
          lc_xml_attr_errors             TYPE string VALUE 'errors',
          lc_xml_attr_firstpagenumber    TYPE string VALUE 'firstPageNumber',
          lc_xml_attr_fittopage          TYPE string VALUE 'fitToPage',
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
          lc_xml_attr_colid              TYPE string VALUE 'colId',
          lc_xml_attr_filtermode         TYPE string VALUE 'filterMode',
          lc_xml_attr_tabcolor_rgb       TYPE string VALUE 'rgb',
          " Node namespace
          lc_xml_node_ns                 TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          lc_xml_node_comp_ns            TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
          lc_xml_node_comp_pref          TYPE string VALUE 'x14ac',
          lc_xml_node_ig_ns              TYPE string VALUE 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_element         TYPE REF TO if_ixml_element,
          lo_element_2       TYPE REF TO if_ixml_element,
          lo_element_3       TYPE REF TO if_ixml_element,
          lo_element_4       TYPE REF TO if_ixml_element,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_style_cond      TYPE REF TO zcl_excel_style_cond,
          lo_data_validation TYPE REF TO zcl_excel_data_validation,
          lo_table           TYPE REF TO zcl_excel_table,
          lo_column_default  TYPE REF TO zcl_excel_column,
          lo_row_default     TYPE REF TO zcl_excel_row.

    DATA: lv_value                    TYPE string,
          lt_range_merge              TYPE string_table,
          lv_column                   TYPE zexcel_cell_column,
          lv_style_guid               TYPE zexcel_cell_style,
          ls_databar                  TYPE zexcel_conditional_databar,      " Databar by Albert Lladanosa
          ls_colorscale               TYPE zexcel_conditional_colorscale,
          ls_iconset                  TYPE zexcel_conditional_iconset,
          ls_cellis                   TYPE zexcel_conditional_cellis,
          ls_textfunction             TYPE zcl_excel_style_cond=>ts_conditional_textfunction,
          lv_column_start             TYPE zexcel_cell_column_alpha,
          lv_row_start                TYPE zexcel_cell_row,
          lv_cell_coords              TYPE zexcel_cell_coords,
          ls_expression               TYPE zexcel_conditional_expression,
          ls_conditional_top10        TYPE zexcel_conditional_top10,
          ls_conditional_above_avg    TYPE zexcel_conditional_above_avg,
          lt_cfvo                     TYPE TABLE OF cfvo,
          ls_cfvo                     TYPE cfvo,
          lt_colors                   TYPE TABLE OF colors,
          ls_colors                   TYPE colors,
          lv_cell_row_s               TYPE string,
          ls_style_mapping            TYPE zexcel_s_styles_mapping,
          lv_freeze_cell_row          TYPE zexcel_cell_row,
          lv_freeze_cell_column       TYPE zexcel_cell_column,
          lv_freeze_cell_column_alpha TYPE zexcel_cell_column_alpha,
          lo_column_iterator          TYPE REF TO zcl_excel_collection_iterator,
          lo_column                   TYPE REF TO zcl_excel_column,
          lo_row_iterator             TYPE REF TO zcl_excel_collection_iterator,
          ls_style_cond_mapping       TYPE zexcel_s_styles_cond_mapping,
          lv_relation_id              TYPE i VALUE 0,
          outline_level_col           TYPE i VALUE 0,
          lts_row_outlines            TYPE zcl_excel_worksheet=>mty_ts_outlines_row,
          merge_count                 TYPE int4,
          lt_values                   TYPE zexcel_t_autofilter_values,
          ls_values                   TYPE zexcel_s_autofilter_values,
          lo_autofilters              TYPE REF TO zcl_excel_autofilters,
          lo_autofilter               TYPE REF TO zcl_excel_autofilter,
          lv_ref                      TYPE string,
          lt_condformating_ranges     TYPE ty_condformating_ranges,
          ls_condformating_range      TYPE ty_condformating_range,
          ld_first_half               TYPE string,
          ld_second_half              TYPE string.

    FIELD-SYMBOLS: <ls_sheet_content>       TYPE zexcel_s_cell_data,
                   <fs_range_merge>         LIKE LINE OF lt_range_merge,
                   <ls_row_outline>         LIKE LINE OF lts_row_outlines,
                   <ls_condformating_range> TYPE ty_condformating_range.

*--------------------------------------------------------------------*
* issue #220 - If cell in tables-area don't use default from row or column or sheet - Declarations 2 - start
*--------------------------------------------------------------------*
    DATA: lt_table_areas TYPE SORTED TABLE OF lty_table_area WITH NON-UNIQUE KEY left right top bottom.

*--------------------------------------------------------------------*
* issue #220 - If cell in tables-area don't use default from row or column or sheet - Declarations 2 - end
*--------------------------------------------------------------------*



**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_worksheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:mc'
                                       value = lc_xml_node_comp_ns ).
    lo_element_root->set_attribute_ns( name  = 'mc:Ignorable'
                                       value = lc_xml_node_comp_pref ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:x14ac'
                                       value = lc_xml_node_ig_ns ).


**********************************************************************
* STEP 4: Create subnodes
    " sheetPr
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheetpr
                                                     parent = lo_document ).
    " TODO tabColor
    IF io_worksheet->tabcolor IS NOT INITIAL.
      lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_tabcolor
                                                         parent = lo_element ).
* Theme not supported yet - start with RGB
      lv_value = io_worksheet->tabcolor-rgb.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabcolor_rgb
                                      value = lv_value ).
    ENDIF.

    " outlinePr
    lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_outlinepr
                                                       parent = lo_document ).

    lv_value = io_worksheet->zif_excel_sheet_properties~summarybelow.
    CONDENSE lv_value.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_summarybelow
                                    value = lv_value ).

    lv_value = io_worksheet->zif_excel_sheet_properties~summaryright.
    CONDENSE lv_value.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_summaryright
                                    value = lv_value ).

    lo_element->append_child( new_child = lo_element_2 ).

    IF io_worksheet->sheet_setup->fit_to_page IS NOT INITIAL.
      lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_pagesetuppr
                                                       parent = lo_document ).
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_fittopage
                                    value = `1` ).
      lo_element->append_child( new_child = lo_element_2 ). " pageSetupPr node
    ENDIF.

    lo_element_root->append_child( new_child = lo_element ).

    " dimension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_dimension
                                                     parent = lo_document ).
    lv_value = io_worksheet->get_dimension_range( ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_ref
                                  value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " sheetViews node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheetviews
                                                     parent = lo_document ).
    " sheetView node
    lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_sheetview
                                                       parent = lo_document ).
    IF io_worksheet->zif_excel_sheet_properties~show_zeros EQ abap_false.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showzeros
                                      value = '0' ).
    ENDIF.
    IF   iv_active = abap_true
      OR io_worksheet->zif_excel_sheet_properties~selected EQ abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabselected
                                      value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_tabselected
                                      value = '0' ).
    ENDIF.
    " Zoom scale
    IF io_worksheet->zif_excel_sheet_properties~zoomscale GT 400.
      io_worksheet->zif_excel_sheet_properties~zoomscale = 400.
    ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale LT 10.
      io_worksheet->zif_excel_sheet_properties~zoomscale = 10.
    ENDIF.
    lv_value = io_worksheet->zif_excel_sheet_properties~zoomscale.
    CONDENSE lv_value.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscale
                                      value = lv_value ).
    IF io_worksheet->zif_excel_sheet_properties~zoomscale_normal NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_normal GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_normal = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_normal LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_normal = 10.
      ENDIF.
      lv_value = io_worksheet->zif_excel_sheet_properties~zoomscale_normal.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalenormal
                                      value = lv_value ).
    ENDIF.
    IF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview = 10.
      ENDIF.
      lv_value = io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalepageview
                                      value = lv_value ).
    ENDIF.
    IF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview NE 0.
      IF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview GT 400.
        io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 400.
      ELSEIF io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview LT 10.
        io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = 10.
      ENDIF.
      lv_value = io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview.
      CONDENSE lv_value.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_zoomscalesheetview
                                      value = lv_value ).
    ENDIF.
    IF io_worksheet->zif_excel_sheet_properties~get_right_to_left( ) EQ abap_true.
      lo_element_2->set_attribute_ns( name  = 'rightToLeft'
                                      value = '1' ).
    ENDIF.
    lo_element_2->set_attribute_ns( name  = lc_xml_attr_workbookviewid
                                            value = '0' ).
    " showGridLines attribute
    IF io_worksheet->show_gridlines = abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showgridlines
                                              value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showgridlines
                                              value = '0' ).
    ENDIF.

    " showRowColHeaders attribute
    IF io_worksheet->show_rowcolheaders = abap_true.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showrowcolheaders
                                              value = '1' ).
    ELSE.
      lo_element_2->set_attribute_ns( name  = lc_xml_attr_showrowcolheaders
                                              value = '0' ).
    ENDIF.


    " freeze panes
    io_worksheet->get_freeze_cell( IMPORTING ep_row = lv_freeze_cell_row
                                   ep_column = lv_freeze_cell_column ).

    IF lv_freeze_cell_row IS NOT INITIAL AND lv_freeze_cell_column IS NOT INITIAL.
      lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_pane
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

      lv_freeze_cell_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column = lv_freeze_cell_column ).
      lv_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_freeze_cell_row  ).
      CONCATENATE lv_freeze_cell_column_alpha lv_value INTO lv_value.
      lo_element_3->set_attribute_ns( name  = 'topLeftCell'
                                    value = lv_value ).

      lo_element_3->set_attribute_ns( name  = 'activePane'
                                    value = 'bottomRight' ).

      lo_element_3->set_attribute_ns( name  = 'state'
                                    value = 'frozen' ).

      lo_element_2->append_child( new_child = lo_element_3 ).
    ENDIF.
    " selection node
    lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_selection
                                                       parent = lo_document ).
    lv_value = io_worksheet->get_active_cell( ).
    lo_element_3->set_attribute_ns( name  = lc_xml_attr_activecell
                                    value = lv_value ).

    lo_element_3->set_attribute_ns( name  = lc_xml_attr_sqref
                                    value = lv_value ).

    lo_element_2->append_child( new_child = lo_element_3 ). " sheetView node

    lo_element->append_child( new_child = lo_element_2 ). " sheetView node

    lo_element_root->append_child( new_child = lo_element ). " sheetViews node


    lo_column_iterator = io_worksheet->get_columns_iterator( ).
    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    " Calculate col
    IF NOT lo_column_iterator IS BOUND.
      io_worksheet->calculate_column_widths( ).
      lo_column_iterator = io_worksheet->get_columns_iterator( ).
    ENDIF.

    " sheetFormatPr node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheetformatpr
                                                     parent = lo_document ).
    " defaultRowHeight
    lo_row_default = io_worksheet->get_default_row( ).
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
    lo_column_default = io_worksheet->get_default_column( ).
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

    lo_element_root->append_child( new_child = lo_element ). " sheetFormatPr node

* Reset column iterator
    lo_column_iterator = io_worksheet->get_columns_iterator( ).
    IF io_worksheet->zif_excel_sheet_properties~get_style( ) IS NOT INITIAL OR lo_column_iterator->has_next( ) = abap_true.
      " cols node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_cols
                                                       parent = lo_document ).
      " This code have to be enhanced in order to manage also column style properties
      " Now it is an out/out
      IF lo_column_iterator->has_next( ) = abap_true.
        WHILE lo_column_iterator->has_next( ) = abap_true.
          lo_column ?= lo_column_iterator->get_next( ).
          " col node
          lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_col
                                                             parent = lo_document ).
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
          READ TABLE styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
          IF sy-subrc = 0.                                                                                     "ins issue #295
            lv_value = ls_style_mapping-style.                                                                 "ins issue #295
            SHIFT lv_value RIGHT DELETING TRAILING space.
            SHIFT lv_value LEFT DELETING LEADING space.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_style
                                            value = lv_value ).
          ENDIF.                                                                                              "ins issue #237

          lo_element->append_child( new_child = lo_element_2 ). " col node
        ENDWHILE.
*    ELSE.                                                                      "del issue #157  -  set sheet style ( add missing columns
*      IF io_worksheet->zif_excel_sheet_properties~get_style( ) IS NOT INITIAL. "del issue #157  -  set sheet style ( add missing columns
* Begin of insertion issue #157 -  set sheet style ( add missing columns
      ENDIF.
* Always pass through this coding
      IF io_worksheet->zif_excel_sheet_properties~get_style( ) IS NOT INITIAL.
        DATA: lts_sorted_columns TYPE SORTED TABLE OF zexcel_cell_column WITH UNIQUE KEY table_line.
        TYPES: BEGIN OF ty_missing_columns,
                 first_column TYPE zexcel_cell_column,
                 last_column  TYPE zexcel_cell_column,
               END OF ty_missing_columns.
        DATA: t_missing_columns TYPE STANDARD TABLE OF ty_missing_columns WITH NON-UNIQUE DEFAULT KEY,
              missing_column    LIKE LINE OF t_missing_columns.

* First collect columns that were already handled before.  The rest has to be inserted now
        lo_column_iterator = io_worksheet->get_columns_iterator( ).
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
          lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_col
                                                             parent = lo_document ).
*        lv_value = zcl_excel_common=>c_excel_sheet_min_col."del issue #157  -  set sheet style ( add missing columns
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
          lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
          READ TABLE styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
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
      IF io_worksheet->zif_excel_sheet_properties~hide_columns_from IS NOT INITIAL.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_col
                                                           parent = lo_document ).
        lv_value = zcl_excel_common=>convert_column2int( io_worksheet->zif_excel_sheet_properties~hide_columns_from ).
        CONDENSE lv_value NO-GAPS.
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_min
                                        value = lv_value ).
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_max
                                        value = '16384' ).
        lo_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                        value = '1' ).
        lo_element->append_child( new_child = lo_element_2 ). " col node
      ENDIF.

      lo_element_root->append_child( new_child = lo_element ). " cols node
    ENDIF.

*--------------------------------------------------------------------*
* Sheet content - use own method to create this
*--------------------------------------------------------------------*
    lo_element = create_xl_sheet_sheet_data( io_worksheet         = io_worksheet
                                             io_document          = lo_document )  .

    lo_autofilters = excel->get_autofilters_reference( ).
    lo_autofilter  = lo_autofilters->get( io_worksheet = io_worksheet ) .
    lo_element_root->append_child( new_child = lo_element ). " sheetData node

*< Begin of insertion Issue #572 - Protect sheet with filter caused Excel error
* Autofilter must be set AFTER sheet protection in XML
    IF io_worksheet->zif_excel_sheet_protection~protected EQ abap_true.
      " sheetProtection node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheetprotection
                                                       parent = lo_document ).
      lv_value = io_worksheet->zif_excel_sheet_protection~password.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_password
                                      value = lv_value ).
      ENDIF.
      lv_value = io_worksheet->zif_excel_sheet_protection~auto_filter.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_autofilter
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~delete_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_deletecolumns
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~delete_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_deleterows
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~format_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatcells
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~format_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatcolumns
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~format_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_formatrows
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~insert_columns.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_insertcolumns
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~insert_hyperlinks.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_inserthyperlinks
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~insert_rows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_insertrows
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~objects.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_objects
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~pivot_tables.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_pivottables
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~scenarios.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_scenarios
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~select_locked_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_selectlockedcells
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~select_unlocked_cells.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_selectunlockedcell
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~sheet.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_sheet
                                    value = lv_value ).
      lv_value = io_worksheet->zif_excel_sheet_protection~sort.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_sort
                                    value = lv_value ).

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
*> End of insertion Issue #572 - Protect sheet with filter caused Excel error

    IF lo_autofilter IS BOUND.
* Create node autofilter
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_autofilter
                                                       parent = lo_document ).
      lv_ref = lo_autofilter->get_filter_range( ) .
      CONDENSE lv_ref NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_ref
                                    value = lv_ref ).
      lt_values = lo_autofilter->get_values( ) .
      IF lt_values IS NOT INITIAL.
* If we filter we need to set the filter mode to 1.
        lo_element_2 = lo_document->find_from_name( name   = lc_xml_node_sheetpr ).
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
            lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_filtercolumn
                                                               parent = lo_element ).
            lv_column   = ls_values-column - lo_autofilter->filter_area-col_start.
            lv_value = lv_column.
            CONDENSE lv_value NO-GAPS.
            lo_element_2->set_attribute_ns( name  = lc_xml_attr_colid
                                            value = lv_value ).
            lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_filters
                                                               parent = lo_element_2 ).
            lv_column = ls_values-column.
          ENDIF.
          lo_element_4 = lo_document->create_simple_element( name   = lc_xml_node_filter
                                                             parent = lo_element_3 ).
          lo_element_4->set_attribute_ns( name  = lc_xml_attr_val
                                          value = ls_values-value ).
          lo_element_3->append_child( new_child = lo_element_4 ). " value node
        ENDLOOP.
        lo_element_2->append_child( new_child = lo_element_3 ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

*< Comment for Issue #572 - Protect sheet with filter caused Excel error
*  IF io_worksheet->zif_excel_sheet_protection~protected EQ abap_true.
*   " sheetProtection node
*    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheetprotection
*                                                     parent = lo_document ).
*    MOVE io_worksheet->zif_excel_sheet_protection~password TO lv_value.
*    IF lv_value IS NOT INITIAL.
*      lo_element->set_attribute_ns( name  = lc_xml_attr_password
*                                    value = lv_value ).
*    ENDIF.
*    lv_value = io_worksheet->zif_excel_sheet_protection~auto_filter.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_autofilter
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~delete_columns.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_deletecolumns
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~delete_rows.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_deleterows
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~format_cells.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_formatcells
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~format_columns.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_formatcolumns
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~format_rows.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_formatrows
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~insert_columns.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_insertcolumns
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~insert_hyperlinks.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_inserthyperlinks
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~insert_rows.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_insertrows
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~objects.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_objects
*                                  value = lv_value ).
*   lv_value = io_worksheet->zif_excel_sheet_protection~pivot_tables.
*   CONDENSE lv_value NO-GAPS.
*   lo_element->set_attribute_ns( name  = lc_xml_attr_pivottables
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~scenarios.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_scenarios
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~select_locked_cells.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_selectlockedcells
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~select_unlocked_cells.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_selectunlockedcell
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~sheet.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_sheet
*                                  value = lv_value ).
*    lv_value = io_worksheet->zif_excel_sheet_protection~sort.
*    CONDENSE lv_value NO-GAPS.
*    lo_element->set_attribute_ns( name  = lc_xml_attr_sort
*                                  value = lv_value ).
*
*    lo_element_root->append_child( new_child = lo_element ).
*  ENDIF.
*> End of Comment for Issue #572 - Protect sheet with filter caused Excel error
    " Merged cells
    lt_range_merge = io_worksheet->get_merge( ).
    IF lt_range_merge IS NOT INITIAL.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_mergecells
                                                       parent = lo_document ).
      DESCRIBE TABLE lt_range_merge LINES merge_count.
      lv_value = merge_count.
      CONDENSE lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                    value = lv_value ).
      LOOP AT lt_range_merge ASSIGNING <fs_range_merge>.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_mergecell
                                                           parent = lo_document ).

        lo_element_2->set_attribute_ns( name  = lc_xml_attr_ref
                                        value = <fs_range_merge> ).
        lo_element->append_child( new_child = lo_element_2 ).
        lo_element_root->append_child( new_child = lo_element ).
        io_worksheet->delete_merge( ).
      ENDLOOP.
    ENDIF.

    " Conditional formatting node
    lo_iterator = io_worksheet->get_style_cond_iterator( ).
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
        lo_element = lo_document->create_simple_element( name   = lc_xml_node_condformatting
                                                         parent = lo_document ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_sqref
                                      value = lv_value ).

        ls_condformating_range-dimension_range = lv_value.
        ls_condformating_range-condformatting_node = lo_element.
        INSERT ls_condformating_range INTO TABLE lt_condformating_ranges.

      ENDIF.

      " cfRule node
      lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_cfrule
                                                       parent = lo_document ).
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
          lo_element_3 = lo_document->create_simple_element( name = lc_xml_node_databar
                                                             parent = lo_document ).

          ls_cfvo-value = ls_databar-cfvo1_value.
          ls_cfvo-type = ls_databar-cfvo1_type.
          APPEND ls_cfvo TO lt_cfvo.

          ls_cfvo-value = ls_databar-cfvo2_value.
          ls_cfvo-type = ls_databar-cfvo2_type.
          APPEND ls_cfvo TO lt_cfvo.

          LOOP AT lt_cfvo INTO ls_cfvo.
            " cfvo node
            lo_element_4 = lo_document->create_simple_element( name = lc_xml_node_cfvo
                                                               parent = lo_document ).
            lv_value = ls_cfvo-type.
            lo_element_4->set_attribute_ns( name = lc_xml_attr_type
                                            value = lv_value ).
            lv_value = ls_cfvo-value.
            lo_element_4->set_attribute_ns( name = lc_xml_attr_val
                                            value = lv_value ).
            lo_element_3->append_child( new_child = lo_element_4 ). " cfvo node
          ENDLOOP.

          lo_element_4 = lo_document->create_simple_element( name = lc_xml_node_color
                                                             parent = lo_document ).
          lv_value = ls_databar-colorrgb.
          lo_element_4->set_attribute_ns( name = lc_xml_attr_tabcolor_rgb
                                          value = lv_value ).

          lo_element_3->append_child( new_child = lo_element_4 ). " color node

          lo_element_2->append_child( new_child = lo_element_3 ). " databar node
          " End << Databar by Albert Lladanosa

        WHEN zcl_excel_style_cond=>c_rule_colorscale.

          ls_colorscale = lo_style_cond->mode_colorscale.

          CLEAR: lt_cfvo, lt_colors.
          lo_element_3 = lo_document->create_simple_element( name = lc_xml_node_colorscale
                                                             parent = lo_document ).

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
            lo_element_4 = lo_document->create_simple_element( name = lc_xml_node_cfvo
                                                               parent = lo_document ).
            lv_value = ls_cfvo-type.
            lo_element_4->set_attribute_ns( name = lc_xml_attr_type
                                            value = lv_value ).
            lv_value = ls_cfvo-value.
            lo_element_4->set_attribute_ns( name = lc_xml_attr_val
                                            value = lv_value ).
            lo_element_3->append_child( new_child = lo_element_4 ). " cfvo node
          ENDLOOP.
          LOOP AT lt_colors INTO ls_colors.

            IF ls_colors IS INITIAL.
              CONTINUE.
            ENDIF.

            lo_element_4 = lo_document->create_simple_element( name = lc_xml_node_color
                                                               parent = lo_document ).
            lv_value = ls_colors-colorrgb.
            lo_element_4->set_attribute_ns( name = lc_xml_attr_tabcolor_rgb
                                            value = lv_value ).

            lo_element_3->append_child( new_child = lo_element_4 ). " color node
          ENDLOOP.

          lo_element_2->append_child( new_child = lo_element_3 ). " databar node

        WHEN zcl_excel_style_cond=>c_rule_iconset.

          ls_iconset = lo_style_cond->mode_iconset.

          CLEAR lt_cfvo.
          " iconset node
          lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_iconset
                                                             parent = lo_document ).
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
            lo_element_4 = lo_document->create_simple_element( name   = lc_xml_node_cfvo
                                                               parent = lo_document ).
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
          READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_cellis-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          lv_value = ls_cellis-operator.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_operator
                                          value = lv_value ).
          " formula node
          lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula
                                                             parent = lo_document ).
          lv_value = ls_cellis-formula.
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node
          IF ls_cellis-formula2 IS NOT INITIAL.
            lv_value = ls_cellis-formula2.
            lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula
                                                               parent = lo_document ).
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
          READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_cellis-cell_style.
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
          lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula
                                                             parent = lo_document ).
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node

        WHEN zcl_excel_style_cond=>c_rule_expression.
          ls_expression = lo_style_cond->mode_expression.
          READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_expression-cell_style.
          lv_value = ls_style_cond_mapping-dxf.
          CONDENSE lv_value.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_dxfid
                                          value = lv_value ).
          " formula node
          lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula
                                                             parent = lo_document ).
          lv_value = ls_expression-formula.
          lo_element_3->set_value( value = lv_value ).
          lo_element_2->append_child( new_child = lo_element_3 ). " formula node

* begin of ins issue #366 - missing conditional rules: top10
        WHEN zcl_excel_style_cond=>c_rule_top10.
          ls_conditional_top10 = lo_style_cond->mode_top10.
          READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_conditional_top10-cell_style.
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
                                            value ='1' ).
          ENDIF.

        WHEN zcl_excel_style_cond=>c_rule_above_average.
          ls_conditional_above_avg = lo_style_cond->mode_above_average.
          READ TABLE me->styles_cond_mapping INTO ls_style_cond_mapping WITH KEY guid = ls_conditional_above_avg-cell_style.
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

      lo_element_root->append_child( new_child = lo_element ). " Conditional formatting node
    ENDWHILE.

    IF io_worksheet->get_data_validations_size( ) GT 0.
      " dataValidations node
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_datavalidations
                                                       parent = lo_document ).
      " Conditional formatting node
      lo_iterator = io_worksheet->get_data_validations_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_data_validation ?= lo_iterator->get_next( ).
        " dataValidation node
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_datavalidation
                                                           parent = lo_document ).
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
        lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula1
                                                           parent = lo_document ).
        lv_value = lo_data_validation->formula1.
        lo_element_3->set_value( value = lv_value ).

        lo_element_2->append_child( new_child = lo_element_3 ). " formula1 node
        " formula2 node
        IF NOT lo_data_validation->formula2 IS INITIAL.
          lo_element_3 = lo_document->create_simple_element( name   = lc_xml_node_formula2
                                                             parent = lo_document ).
          lv_value = lo_data_validation->formula2.
          lo_element_3->set_value( value = lv_value ).

          lo_element_2->append_child( new_child = lo_element_3 ). " formula2 node
        ENDIF.

        lo_element->append_child( new_child = lo_element_2 ). " dataValidation node
      ENDWHILE.
      lo_element_root->append_child( new_child = lo_element ). " dataValidations node
    ENDIF.

    " Hyperlinks
    DATA: lv_hyperlinks_count TYPE i,
          lo_link             TYPE REF TO zcl_excel_hyperlink.

    lv_hyperlinks_count = io_worksheet->get_hyperlinks_size( ).
    IF lv_hyperlinks_count > 0.
      lo_element = lo_document->create_simple_element( name   = 'hyperlinks'
                                                        parent = lo_document ).

      lo_iterator = io_worksheet->get_hyperlinks_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_link ?= lo_iterator->get_next( ).

        lo_element_2 = lo_document->create_simple_element( name   = 'hyperlink'
                                                      parent = lo_element ).

        lv_value = lo_link->get_ref( ).
        lo_element_2->set_attribute_ns( name  = 'ref'
                                       value = lv_value ).

        IF lo_link->is_internal( ) = abap_true.
          lv_value = lo_link->get_url( ).
          lo_element_2->set_attribute_ns( name  = 'location'
                                       value = lv_value ).
        ELSE.
          ADD 1 TO lv_relation_id.

          lv_value = lv_relation_id.
          CONDENSE lv_value.
          CONCATENATE 'rId' lv_value INTO lv_value.

          lo_element_2->set_attribute_ns( name  = 'r:id'
                                        value = lv_value ).

        ENDIF.

        lo_element->append_child( new_child = lo_element_2 ).
      ENDWHILE.

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.


    " PrintOptions
    IF io_worksheet->print_gridlines                  = abap_true OR
       io_worksheet->sheet_setup->vertical_centered   = abap_true OR
       io_worksheet->sheet_setup->horizontal_centered = abap_true.
      lo_element = lo_document->create_simple_element( name   = 'printOptions'
                                                       parent = lo_document ).

      IF io_worksheet->print_gridlines = abap_true.
        lo_element->set_attribute_ns( name  = lc_xml_attr_gridlines
                                      value = 'true' ).
      ENDIF.

      IF io_worksheet->sheet_setup->horizontal_centered = abap_true.
        lo_element->set_attribute_ns( name  = 'horizontalCentered'
                                      value = 'true' ).
      ENDIF.

      IF io_worksheet->sheet_setup->vertical_centered = abap_true.
        lo_element->set_attribute_ns( name  = 'verticalCentered'
                                      value = 'true' ).
      ENDIF.

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
    " pageMargins node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_pagemargins
                                                     parent = lo_document ).

    lv_value = io_worksheet->sheet_setup->margin_left.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_left
                                  value = lv_value ).
    lv_value = io_worksheet->sheet_setup->margin_right.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_right
                                  value = lv_value ).
    lv_value = io_worksheet->sheet_setup->margin_top.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_top
                                  value = lv_value ).
    lv_value = io_worksheet->sheet_setup->margin_bottom.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_bottom
                                  value = lv_value ).
    lv_value = io_worksheet->sheet_setup->margin_header.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_header
                                  value = lv_value ).
    lv_value = io_worksheet->sheet_setup->margin_footer.
    CONDENSE lv_value NO-GAPS.
    lo_element->set_attribute_ns( name  = lc_xml_attr_footer
                                  value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ). " pageMargins node

* pageSetup node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_pagesetup
                                                     parent = lo_document ).

    IF io_worksheet->sheet_setup->black_and_white IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_blackandwhite
                                    value = `1` ).
    ENDIF.

    IF io_worksheet->sheet_setup->cell_comments IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_cellcomments
                                    value = io_worksheet->sheet_setup->cell_comments ).
    ENDIF.

    IF io_worksheet->sheet_setup->copies IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->copies.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_copies
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->draft IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_draft
                                    value = `1` ).
    ENDIF.

    IF io_worksheet->sheet_setup->errors IS NOT INITIAL.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_errors
                                    value = io_worksheet->sheet_setup->errors ).
    ENDIF.

    IF io_worksheet->sheet_setup->first_page_number IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->first_page_number.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_firstpagenumber
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->fit_to_page IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->fit_to_height.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fittoheight
                                    value = lv_value ).
      lv_value = io_worksheet->sheet_setup->fit_to_width.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fittowidth
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->horizontal_dpi IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->horizontal_dpi.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_horizontaldpi
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->orientation IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->orientation.
      lo_element->set_attribute_ns( name  = lc_xml_attr_orientation
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->page_order IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_pageorder
                                    value = io_worksheet->sheet_setup->page_order ).
    ENDIF.

    IF io_worksheet->sheet_setup->paper_height IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->paper_height.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_paperheight
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->paper_size IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->paper_size.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_papersize
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->paper_width IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->paper_width.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_paperwidth
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->scale IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->scale.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_scale
                                    value = lv_value ).
    ENDIF.

    IF io_worksheet->sheet_setup->use_first_page_num IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_usefirstpagenumber
                                    value = `1` ).
    ENDIF.

    IF io_worksheet->sheet_setup->use_printer_defaults IS NOT INITIAL.
      lo_element->set_attribute_ns( name  = lc_xml_attr_useprinterdefaults
                                    value = `1` ).
    ENDIF.

    IF io_worksheet->sheet_setup->vertical_dpi IS NOT INITIAL.
      lv_value = io_worksheet->sheet_setup->vertical_dpi.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_verticaldpi
                                    value = lv_value ).
    ENDIF.

    lo_element_root->append_child( new_child = lo_element ). " pageSetup node

* { headerFooter necessary?   >
    IF    io_worksheet->sheet_setup->odd_header IS NOT INITIAL
       OR io_worksheet->sheet_setup->odd_footer IS NOT INITIAL
       OR io_worksheet->sheet_setup->diff_oddeven_headerfooter = abap_true.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_headerfooter
                                                       parent = lo_document ).

      " Different header/footer for odd/even pages?
      IF io_worksheet->sheet_setup->diff_oddeven_headerfooter = abap_true.
        lo_element->set_attribute_ns( name  = lc_xml_attr_differentoddeven
                                      value = '1' ).
      ENDIF.

      " OddHeader
      CLEAR: lv_value.
      io_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_header = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_oddheader
                                                           parent = lo_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " OddFooter
      CLEAR: lv_value.
      io_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_odd_footer = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_oddfooter
                                                           parent = lo_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " evenHeader
      CLEAR: lv_value.
      io_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_even_header = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_evenheader
                                                           parent = lo_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.

      " evenFooter
      CLEAR: lv_value.
      io_worksheet->sheet_setup->get_header_footer_string( IMPORTING ep_even_footer = lv_value ) .
      IF lv_value IS NOT INITIAL.
        lo_element_2 = lo_document->create_simple_element( name   = lc_xml_node_evenfooter
                                                           parent = lo_document ).
        lo_element_2->set_value( value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).
      ENDIF.


      lo_element_root->append_child( new_child = lo_element ). " headerFooter

    ENDIF.

* issue #377 pagebreaks
    TRY.
        create_xl_sheet_pagebreaks( io_document  = lo_document
                                    io_parent    = lo_element_root
                                    io_worksheet = io_worksheet     )  .
      CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    ENDTRY.

* drawing
    DATA: lo_drawings TYPE REF TO zcl_excel_drawings.

    lo_drawings = io_worksheet->get_drawings( ).
    IF lo_drawings->is_empty( ) = abap_false.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_drawing
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name = 'r:id'
                                 value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

* Begin - Add - Issue #180
    " (Legacy) drawings for comments
    DATA: lo_drawing_for_comments  TYPE REF TO zcl_excel_comments.

    lo_drawing_for_comments = io_worksheet->get_comments( ).
    IF lo_drawing_for_comments->is_empty( ) = abap_false.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_drawing_for_cmt
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.  " +1 for legacyDrawings

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name = 'r:id'
                                 value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).

      ADD 1 TO lv_relation_id.  " +1 for comments (not referenced in XL sheet but let's reserve the rId)
    ENDIF.
* End   - Add - Issue #180

* Header/Footer Image
    DATA: lt_drawings TYPE zexcel_t_drawings.
    lt_drawings = io_worksheet->get_header_footer_drawings( ).
    IF lines( lt_drawings ) > 0. "Header or footer image exist
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_drawing_for_hd_ft
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.  " +1 for legacyDrawings
      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute( name = 'r:id'
                                 value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
      ADD 1 TO lv_relation_id.  " +1 for comments (not referenced in XL sheet but let's reserve the rId)
    ENDIF.
*

* ignoredErrors
    create_xl_sheet_ignored_errors( io_worksheet = io_worksheet io_document = lo_document io_element_root = lo_element_root ).


* tables
    DATA lv_table_count TYPE i.

    lv_table_count = io_worksheet->get_tables_size( ).
    IF lv_table_count > 0.
      lo_element = lo_document->create_simple_element( name   = 'tableParts'
                                                        parent = lo_document ).
      lv_value = lv_table_count.
      CONDENSE lv_value.
      lo_element->set_attribute_ns( name  = 'count'
                                     value = lv_value ).

      lo_iterator = io_worksheet->get_tables_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_iterator->get_next( ).
        ADD 1 TO lv_relation_id.

        lv_value = lv_relation_id.
        CONDENSE lv_value.
        CONCATENATE 'rId' lv_value INTO lv_value.
        lo_element_2 = lo_document->create_simple_element( name   = 'tablePart'
                                                        parent = lo_element ).
        lo_element_2->set_attribute_ns( name  = 'r:id'
                                       value = lv_value ).
        lo_element->append_child( new_child = lo_element_2 ).

      ENDWHILE.

      lo_element_root->append_child( new_child = lo_element ).

    ENDIF.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_sheet_column_formula.

    TYPES: ls_column_formula_used     TYPE mty_column_formula_used,
           lv_column_alpha            TYPE zexcel_cell_column_alpha,
           lv_top_cell_coords         TYPE zexcel_cell_coords,
           lv_bottom_cell_coords      TYPE zexcel_cell_coords,
           lv_cell_coords             TYPE zexcel_cell_coords,
           lv_ref_value               TYPE string,
           lv_test_shared             TYPE string,
           lv_si                      TYPE i,
           lv_1st_line_shared_formula TYPE abap_bool.
    DATA: lv_value                   TYPE string,
          ls_column_formula_used     TYPE mty_column_formula_used,
          lv_column_alpha            TYPE zexcel_cell_column_alpha,
          lv_top_cell_coords         TYPE zexcel_cell_coords,
          lv_bottom_cell_coords      TYPE zexcel_cell_coords,
          lv_cell_coords             TYPE zexcel_cell_coords,
          lv_ref_value               TYPE string,
          lv_1st_line_shared_formula TYPE abap_bool.
    FIELD-SYMBOLS: <ls_column_formula>      TYPE zcl_excel_worksheet=>mty_s_column_formula,
                   <ls_column_formula_used> TYPE mty_column_formula_used.


    READ TABLE it_column_formulas WITH TABLE KEY id = is_sheet_content-column_formula_id ASSIGNING <ls_column_formula>.
    ASSERT sy-subrc = 0.

    lv_value = <ls_column_formula>-formula.
    lv_1st_line_shared_formula = abap_false.
    eo_element = io_document->create_simple_element( name   = 'f'
                                                     parent = io_document ).
    READ TABLE ct_column_formulas_used WITH TABLE KEY id = is_sheet_content-column_formula_id ASSIGNING <ls_column_formula_used>.
    IF sy-subrc <> 0.
      CLEAR ls_column_formula_used.
      ls_column_formula_used-id = is_sheet_content-column_formula_id.
      IF is_formula_shareable( ip_formula = lv_value ) = abap_true.
        ls_column_formula_used-t = 'shared'.
        ls_column_formula_used-si = cv_si.
        CONDENSE ls_column_formula_used-si.
        cv_si = cv_si + 1.
        lv_1st_line_shared_formula = abap_true.
      ENDIF.
      INSERT ls_column_formula_used INTO TABLE ct_column_formulas_used ASSIGNING <ls_column_formula_used>.
    ENDIF.

    IF lv_1st_line_shared_formula = abap_true OR <ls_column_formula_used>-t <> 'shared'.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column = is_sheet_content-cell_column ).
      lv_top_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_top_left_row + 1 }|.
      lv_bottom_cell_coords = |{ lv_column_alpha }{ <ls_column_formula>-table_bottom_right_row + 1 }|.
      lv_cell_coords = |{ lv_column_alpha }{ is_sheet_content-cell_row }|.
      IF lv_top_cell_coords = lv_cell_coords.
        lv_ref_value = |{ lv_top_cell_coords }:{ lv_bottom_cell_coords }|.
      ELSE.
        lv_ref_value = |{ lv_cell_coords }:{ lv_bottom_cell_coords }|.
        lv_value = zcl_excel_common=>shift_formula(
            iv_reference_formula = lv_value
            iv_shift_cols        = 0
            iv_shift_rows        = is_sheet_content-cell_row - <ls_column_formula>-table_top_left_row - 1 ).
      ENDIF.
    ENDIF.

    IF <ls_column_formula_used>-t = 'shared'.
      eo_element->set_attribute( name  = 't'
                                 value = <ls_column_formula_used>-t ).
      eo_element->set_attribute( name  = 'si'
                                 value = <ls_column_formula_used>-si ).
      IF lv_1st_line_shared_formula = abap_true.
        eo_element->set_attribute( name  = 'ref'
                                   value = lv_ref_value ).
        eo_element->set_value( value = lv_value ).
      ENDIF.
    ELSE.
      eo_element->set_value( value = lv_value ).
    ENDIF.

  ENDMETHOD.


  METHOD create_xl_sheet_ignored_errors.
    DATA: lo_element        TYPE REF TO if_ixml_element,
          lo_element2       TYPE REF TO if_ixml_element,
          lt_ignored_errors TYPE zcl_excel_worksheet=>mty_th_ignored_errors.
    FIELD-SYMBOLS: <ls_ignored_errors> TYPE zcl_excel_worksheet=>mty_s_ignored_errors.

    lt_ignored_errors = io_worksheet->get_ignored_errors( ).

    IF lt_ignored_errors IS NOT INITIAL.
      lo_element = io_document->create_simple_element( name   = 'ignoredErrors'
                                                       parent = io_document ).


      LOOP AT lt_ignored_errors ASSIGNING <ls_ignored_errors>.

        lo_element2 = io_document->create_simple_element( name   = 'ignoredError'
                                                          parent = io_document ).

        lo_element2->set_attribute_ns( name  = 'sqref'
                                       value = <ls_ignored_errors>-cell_coords ).

        IF <ls_ignored_errors>-eval_error = abap_true.
          lo_element2->set_attribute_ns( name  = 'evalError'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-two_digit_text_year = abap_true.
          lo_element2->set_attribute_ns( name  = 'twoDigitTextYear'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-number_stored_as_text = abap_true.
          lo_element2->set_attribute_ns( name  = 'numberStoredAsText'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'formula'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-formula_range = abap_true.
          lo_element2->set_attribute_ns( name  = 'formulaRange'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-unlocked_formula = abap_true.
          lo_element2->set_attribute_ns( name  = 'unlockedFormula'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-empty_cell_reference = abap_true.
          lo_element2->set_attribute_ns( name  = 'emptyCellReference'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-list_data_validation = abap_true.
          lo_element2->set_attribute_ns( name  = 'listDataValidation'
                                         value = '1' ).
        ENDIF.
        IF <ls_ignored_errors>-calculated_column = abap_true.
          lo_element2->set_attribute_ns( name  = 'calculatedColumn'
                                         value = '1' ).
        ENDIF.

        lo_element->append_child( lo_element2 ).

      ENDLOOP.

      io_element_root->append_child( lo_element ).

    ENDIF.

  ENDMETHOD.


  METHOD create_xl_sheet_pagebreaks.
    DATA: lo_pagebreaks     TYPE REF TO zcl_excel_worksheet_pagebreaks,
          lt_pagebreaks     TYPE zcl_excel_worksheet_pagebreaks=>tt_pagebreak_at,
          lt_rows           TYPE HASHED TABLE OF int4 WITH UNIQUE KEY table_line,
          lt_columns        TYPE HASHED TABLE OF int4 WITH UNIQUE KEY table_line,

          lo_node_rowbreaks TYPE REF TO if_ixml_element,
          lo_node_colbreaks TYPE REF TO if_ixml_element,
          lo_node_break     TYPE REF TO if_ixml_element,

          lv_value          TYPE string.


    FIELD-SYMBOLS: <ls_pagebreak> LIKE LINE OF lt_pagebreaks.

    lo_pagebreaks = io_worksheet->get_pagebreaks( ).
    CHECK lo_pagebreaks IS BOUND.

    lt_pagebreaks = lo_pagebreaks->get_all_pagebreaks( ).
    CHECK lt_pagebreaks IS NOT INITIAL.  " No need to proceed if don't have any pagebreaks.

    lo_node_rowbreaks = io_document->create_simple_element( name   = 'rowBreaks'
                                                            parent = io_document ).

    lo_node_colbreaks = io_document->create_simple_element( name   = 'colBreaks'
                                                            parent = io_document ).


    LOOP AT lt_pagebreaks ASSIGNING <ls_pagebreak>.

* Count how many rows and columns need to be broken
      INSERT <ls_pagebreak>-cell_row    INTO TABLE lt_rows.
      IF sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_row.
        CONDENSE lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).      " Manual break
        lo_node_break->set_attribute( name = 'max' value = '16383' ).  " Max columns

        lo_node_rowbreaks->append_child( new_child = lo_node_break ).
      ENDIF.

      INSERT <ls_pagebreak>-cell_column INTO TABLE lt_columns.
      IF sy-subrc = 0. " New
        lv_value = <ls_pagebreak>-cell_column.
        CONDENSE lv_value.

        lo_node_break = io_document->create_simple_element( name   = 'brk'
                                                            parent = io_document ).
        lo_node_break->set_attribute( name = 'id'  value = lv_value ).
        lo_node_break->set_attribute( name = 'man' value = '1' ).        " Manual break
        lo_node_break->set_attribute( name = 'max' value = '1048575' ).  " Max rows

        lo_node_colbreaks->append_child( new_child = lo_node_break ).
      ENDIF.


    ENDLOOP.

    lv_value = lines( lt_rows ).
    CONDENSE lv_value.
    lo_node_rowbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_rowbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    lv_value = lines( lt_rows ).
    CONDENSE lv_value.
    lo_node_colbreaks->set_attribute( name = 'count'             value = lv_value ).
    lo_node_colbreaks->set_attribute( name = 'manualBreakCount'  value = lv_value ).

    io_parent->append_child( new_child = lo_node_rowbreaks ).
    io_parent->append_child( new_child = lo_node_colbreaks ).

  ENDMETHOD.


  METHOD create_xl_sheet_rels.


** Constant node name
    DATA: lc_xml_node_relationships      TYPE string VALUE 'Relationships',
          lc_xml_node_relationship       TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id                 TYPE string VALUE 'Id',
          lc_xml_attr_type               TYPE string VALUE 'Type',
          lc_xml_attr_target             TYPE string VALUE 'Target',
          lc_xml_attr_target_mode        TYPE string VALUE 'TargetMode',
          lc_xml_val_external            TYPE string VALUE 'External',
          " Node namespace
          lc_xml_node_rels_ns            TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
          lc_xml_node_rid_table_tp       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
          lc_xml_node_rid_printer_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
          lc_xml_node_rid_drawing_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          lc_xml_node_rid_comment_tp     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',        " (+) Issue #180
          lc_xml_node_rid_drawing_cmt_tp TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',      " (+) Issue #180
          lc_xml_node_rid_link_tp        TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_table        TYPE REF TO zcl_excel_table,
          lo_link         TYPE REF TO zcl_excel_hyperlink.

    DATA: lv_value         TYPE string,
          lv_relation_id   TYPE i,
          lv_index_str     TYPE string,
          lv_comment_index TYPE i.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_relationships
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_rels_ns ).

**********************************************************************
* STEP 4: Create subnodes

    " Add sheet Relationship nodes here
    lv_relation_id = 0.
    lo_iterator = io_worksheet->get_hyperlinks_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_link ?= lo_iterator->get_next( ).
      CHECK lo_link->is_internal( ) = abap_false.  " issue #340 - don't put internal links here
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_link_tp ).

      lv_value = lo_link->get_url( ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_target_mode
                                    value = lc_xml_val_external ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

* drawing
    DATA: lo_drawings TYPE REF TO zcl_excel_drawings.

    lo_drawings = io_worksheet->get_drawings( ).
    IF lo_drawings->is_empty( ) = abap_false.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_tp ).

      lv_index_str = iv_drawing_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->c_xl_drawings.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

* Begin - Add - Issue #180
    DATA: lo_comments  TYPE REF TO zcl_excel_comments.

    lv_comment_index = iv_comment_index.

    lo_comments = io_worksheet->get_comments( ).
    IF lo_comments->is_empty( ) = abap_false.
      " Drawing for comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).

      ADD 1 TO lv_relation_id.
      ADD 1 TO lv_comment_index.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = iv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->cl_xl_drawing_for_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).

      " Comment
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_comment_tp ).

      lv_index_str = iv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->c_xl_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
* End   - Add - Issue #180

**********************************************************************
* header footer image
    DATA: lt_drawings TYPE zexcel_t_drawings.
    lt_drawings = io_worksheet->get_header_footer_drawings( ).
    IF lines( lt_drawings ) > 0. "Header or footer image exist
      ADD 1 TO lv_relation_id.
      " Drawing for comment/header/footer
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_drawing_cmt_tp ).

      lv_index_str = lv_comment_index.
      CONDENSE lv_index_str NO-GAPS.
      lv_value = me->cl_xl_drawing_for_comments.
      REPLACE 'xl' WITH '..' INTO lv_value.
      REPLACE '#' WITH lv_index_str INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.
*** End Header Footer
**********************************************************************


    lo_iterator = io_worksheet->get_tables_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ADD 1 TO lv_relation_id.

      lv_value = lv_relation_id.
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.

      lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                       parent = lo_document ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                    value = lv_value ).
      lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                    value = lc_xml_node_rid_table_tp ).

      lv_value = lo_table->get_name( ).
      CONCATENATE '../tables/' lv_value '.xml' INTO lv_value.
      lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDWHILE.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_sheet_sheet_data.

    TYPES: BEGIN OF lty_table_area,
             left   TYPE i,
             right  TYPE i,
             top    TYPE i,
             bottom TYPE i,
           END OF lty_table_area.

    CONSTANTS: lc_dummy_cell_content       TYPE zexcel_s_cell_data-cell_value VALUE '})~~~ This is a dummy value for ABAP2XLSX and you should never find this in a real excelsheet Ihope'.

    CONSTANTS: lc_xml_node_sheetdata TYPE string VALUE 'sheetData',   " SheetData tag
               lc_xml_node_row       TYPE string VALUE 'row',         " Row tag
               lc_xml_attr_r         TYPE string VALUE 'r',           " Cell:  row-attribute
               lc_xml_attr_spans     TYPE string VALUE 'spans',       " Cell: spans-attribute
               lc_xml_node_c         TYPE string VALUE 'c',           " Cell tag
               lc_xml_node_v         TYPE string VALUE 'v',           " Cell: value
               lc_xml_node_f         TYPE string VALUE 'f',           " Cell: formula
               lc_xml_attr_s         TYPE string VALUE 's',           " Cell: style
               lc_xml_attr_t         TYPE string VALUE 't'.           " Cell: type

    DATA: col_count              TYPE int4,
          lo_autofilters         TYPE REF TO zcl_excel_autofilters,
          lo_autofilter          TYPE REF TO zcl_excel_autofilter,
          l_autofilter_hidden    TYPE flag,
          lt_values              TYPE zexcel_t_autofilter_values,
          ls_values              TYPE zexcel_s_autofilter_values,
          ls_area                TYPE zexcel_s_autofilter_area,

          lo_iterator            TYPE REF TO zcl_excel_collection_iterator,
          lo_table               TYPE REF TO zcl_excel_table,
          lt_table_areas         TYPE SORTED TABLE OF lty_table_area WITH NON-UNIQUE KEY left right top bottom,
          ls_table_area          LIKE LINE OF lt_table_areas,
          lo_column              TYPE REF TO zcl_excel_column,

          ls_sheet_content       LIKE LINE OF io_worksheet->sheet_content,
          ls_sheet_content_empty LIKE LINE OF io_worksheet->sheet_content,
          lv_current_row         TYPE i,
          lv_next_row            TYPE i,
          lv_last_row            TYPE i,

*        lts_row_dimensions     TYPE zexcel_t_worksheet_rowdimensio,
          lo_row_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_row                 TYPE REF TO zcl_excel_row,
          lo_row_empty           TYPE REF TO zcl_excel_row,
          lts_row_outlines       TYPE zcl_excel_worksheet=>mty_ts_outlines_row,

          ls_last_row            TYPE zexcel_s_cell_data,
          ls_style_mapping       TYPE zexcel_s_styles_mapping,

          lo_element_2           TYPE REF TO if_ixml_element,
          lo_element_3           TYPE REF TO if_ixml_element,
          lo_element_4           TYPE REF TO if_ixml_element,

          lv_value               TYPE string,
          lv_style_guid          TYPE zexcel_cell_style.
    DATA: lt_column_formulas_used TYPE mty_column_formulas_used,
          lv_si                   TYPE i.

    FIELD-SYMBOLS: <ls_sheet_content> TYPE zexcel_s_cell_data,
                   <ls_row_outline>   LIKE LINE OF lts_row_outlines.


    " sheetData node
    rv_ixml_sheet_data_root = io_document->create_simple_element( name   = lc_xml_node_sheetdata
                                                                  parent = io_document ).

    " Get column count
    col_count      = io_worksheet->get_highest_column( ).
    " Get autofilter
    lo_autofilters = excel->get_autofilters_reference( ).
    lo_autofilter  = lo_autofilters->get( io_worksheet = io_worksheet ) .
    IF lo_autofilter IS BOUND.
      lt_values           = lo_autofilter->get_values( ) .
      ls_area             = lo_autofilter->get_filter_area( ) .
      l_autofilter_hidden = abap_true. " First defautl is not showing
    ENDIF.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - start
*--------------------------------------------------------------------*
*Build table to hold all table-areas attached to this sheet
    lo_iterator = io_worksheet->get_tables_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_table ?= lo_iterator->get_next( ).
      ls_table_area-left   = zcl_excel_common=>convert_column2int( lo_table->settings-top_left_column ).
      ls_table_area-right  = lo_table->get_right_column_integer( ).
      ls_table_area-top    = lo_table->settings-top_left_row.
      ls_table_area-bottom = lo_table->get_bottom_row_integer( ).
      INSERT ls_table_area INTO TABLE lt_table_areas.
    ENDWHILE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 1 - end
*--------------------------------------------------------------------*
*We have problems when the first rows or trailing rows are not set but we have rowinformation
*to solve this we add dummycontent into first and last line that will not be set
*Set first line if necessary
    READ TABLE io_worksheet->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = 1.
    IF sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = 1.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      INSERT ls_sheet_content_empty INTO TABLE io_worksheet->sheet_content.
    ENDIF.
*Set last line if necessary
*Last row with cell content
    lv_last_row = io_worksheet->get_highest_row( ).
*Last line with row-information set directly ( like line height, hidden-status ... )

    lo_row_iterator = io_worksheet->get_rows_iterator( ).
    WHILE lo_row_iterator->has_next( ) = abap_true.
      lo_row ?= lo_row_iterator->get_next( ).
      IF lo_row->get_row_index( ) > lv_last_row.
        lv_last_row = lo_row->get_row_index( ).
      ENDIF.
    ENDWHILE.

*Last line with row-information set indirectly by row outline
    lts_row_outlines = io_worksheet->get_row_outlines( ).
    LOOP AT lts_row_outlines ASSIGNING <ls_row_outline>.
      IF <ls_row_outline>-collapsed = 'X'.
        lv_current_row = <ls_row_outline>-row_to + 1.  " collapsed-status may be set on following row
      ELSE.
        lv_current_row = <ls_row_outline>-row_to.  " collapsed-status may be set on following row
      ENDIF.
      IF lv_current_row > lv_last_row.
        lv_last_row = lv_current_row.
      ENDIF.
    ENDLOOP.
    READ TABLE io_worksheet->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = lv_last_row.
    IF sy-subrc <> 0.
      ls_sheet_content_empty-cell_row      = lv_last_row.
      ls_sheet_content_empty-cell_column   = 1.
      ls_sheet_content_empty-cell_value    = lc_dummy_cell_content.
      INSERT ls_sheet_content_empty INTO TABLE io_worksheet->sheet_content.
    ENDIF.

    CLEAR ls_sheet_content.
    LOOP AT io_worksheet->sheet_content INTO ls_sheet_content.
      IF lt_values IS INITIAL. " no values attached to autofilter  " issue #368 autofilter filtering too much
        CLEAR l_autofilter_hidden.
      ELSE.
        READ TABLE lt_values INTO ls_values WITH KEY column = ls_last_row-cell_column.
        IF sy-subrc = 0 AND ls_values-value = ls_last_row-cell_value.
          CLEAR l_autofilter_hidden.
        ENDIF.
      ENDIF.
      CLEAR ls_style_mapping.
*Create row element
*issues #346,#154, #195  - problems when we have information in row_dimension but no cell content in that row
*Get next line that may have to be added.  If we have empty lines this is the next line after previous cell content
*Otherwise it is the line of the current cell content
      lv_current_row = ls_last_row-cell_row + 1.
      IF lv_current_row > ls_sheet_content-cell_row.
        lv_current_row = ls_sheet_content-cell_row.
      ENDIF.
*Fill in empty lines if necessary - assign an emtpy sheet content
      lv_next_row = lv_current_row.
      WHILE lv_next_row <= ls_sheet_content-cell_row.
        lv_current_row = lv_next_row.
        lv_next_row = lv_current_row + 1.
        IF lv_current_row = ls_sheet_content-cell_row. " cell value found in this row
          ASSIGN ls_sheet_content TO <ls_sheet_content>.
        ELSE.
*Check if empty row is really necessary - this is basically the case when we have information in row_dimension
          lo_row_empty = io_worksheet->get_row( lv_current_row ).
          CHECK lo_row_empty->get_row_height( )                 >= 0          OR
                lo_row_empty->get_collapsed( io_worksheet )      = abap_true  OR
                lo_row_empty->get_outline_level( io_worksheet )  > 0          OR
                lo_row_empty->get_xf_index( )                   <> 0.
          " Dummyentry A1
          ls_sheet_content_empty-cell_row      = lv_current_row.
          ls_sheet_content_empty-cell_column   = 1.
          ASSIGN ls_sheet_content_empty TO <ls_sheet_content>.
        ENDIF.

        IF ls_last_row-cell_row NE <ls_sheet_content>-cell_row.
          IF lo_autofilter IS BOUND.
            IF ls_area-row_start >=  ls_last_row-cell_row OR " One less for header
              ls_area-row_end   < ls_last_row-cell_row .
              CLEAR l_autofilter_hidden.
            ENDIF.
          ELSE.
            CLEAR l_autofilter_hidden.
          ENDIF.
          IF ls_last_row-cell_row IS NOT INITIAL.
            " Row visibility of previos row.
            IF lo_row->get_visible( io_worksheet ) = abap_false OR
               l_autofilter_hidden = abap_true.
              lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
            ENDIF.
            rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
          ENDIF.
          " Add new row
          lo_element_2 = io_document->create_simple_element( name   = lc_xml_node_row
                                                             parent = io_document ).
          " r
          lv_value = <ls_sheet_content>-cell_row.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.

          lo_element_2->set_attribute_ns( name  = lc_xml_attr_r
                                          value = lv_value ).
          " Spans
          lv_value = col_count.
          CONCATENATE '1:' lv_value INTO lv_value.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_element_2->set_attribute_ns( name  = lc_xml_attr_spans
                                          value = lv_value ).
          lo_row = io_worksheet->get_row( <ls_sheet_content>-cell_row ).
          " Row dimensions
          IF lo_row->get_custom_height( ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'customHeight' value = '1' ).
          ENDIF.
          IF lo_row->get_row_height( ) > 0.
            lv_value = lo_row->get_row_height( ).
            lo_element_2->set_attribute_ns( name  = 'ht' value = lv_value ).
          ENDIF.
          " Collapsed
          IF lo_row->get_collapsed( io_worksheet ) = abap_true.
            lo_element_2->set_attribute_ns( name  = 'collapsed' value = 'true' ).
          ENDIF.
          " Outline level
          IF lo_row->get_outline_level( io_worksheet ) > 0.
            lv_value = lo_row->get_outline_level( io_worksheet ).
            SHIFT lv_value RIGHT DELETING TRAILING space.
            SHIFT lv_value LEFT DELETING LEADING space.
            lo_element_2->set_attribute_ns( name  = 'outlineLevel' value = lv_value ).
          ENDIF.
          " Style
          IF lo_row->get_xf_index( ) <> 0.
            lv_value = lo_row->get_xf_index( ).
            lo_element_2->set_attribute_ns( name  = 's' value = lv_value ).
            lo_element_2->set_attribute_ns( name  = 'customFormat'  value = '1' ).
          ENDIF.
          IF lt_values IS INITIAL. " no values attached to autofilter  " issue #368 autofilter filtering too much
            CLEAR l_autofilter_hidden.
          ELSE.
            l_autofilter_hidden = abap_true. " First default is not showing
          ENDIF.
        ELSE.

        ENDIF.
      ENDWHILE.

      lo_element_3 = io_document->create_simple_element( name   = lc_xml_node_c
                                                         parent = io_document ).

      lo_element_3->set_attribute_ns( name  = lc_xml_attr_r
                                      value = <ls_sheet_content>-cell_coords ).

*begin of change issue #157 - allow column cellstyle
*if no cellstyle is set, look into column, then into sheet
      IF <ls_sheet_content>-cell_style IS NOT INITIAL.
        lv_style_guid = <ls_sheet_content>-cell_style.
      ELSE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - start
*--------------------------------------------------------------------*
*Check if cell in any of the table areas
        LOOP AT lt_table_areas TRANSPORTING NO FIELDS WHERE top    <= <ls_sheet_content>-cell_row
                                                        AND bottom >= <ls_sheet_content>-cell_row
                                                        AND left   <= <ls_sheet_content>-cell_column
                                                        AND right  >= <ls_sheet_content>-cell_column.
          EXIT.
        ENDLOOP.
        IF sy-subrc = 0.
          CLEAR lv_style_guid.     " No style --> EXCEL will use built-in-styles as declared in the tables-section
        ELSE.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 2 - end
*--------------------------------------------------------------------*
          lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
          lo_column ?= io_worksheet->get_column( <ls_sheet_content>-cell_column ).
          IF lo_column->get_column_index( ) = <ls_sheet_content>-cell_column.
            lv_style_guid = lo_column->get_column_style_guid( ).
            IF lv_style_guid IS INITIAL.
              lv_style_guid = io_worksheet->zif_excel_sheet_properties~get_style( ).
            ENDIF.
          ENDIF.

*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - start
*--------------------------------------------------------------------*
        ENDIF.
*--------------------------------------------------------------------*
*issue #220 - If cell in tables-area don't use default from row or column or sheet - Coding 3 - end
*--------------------------------------------------------------------*
      ENDIF.
      IF lv_style_guid IS NOT INITIAL.
        READ TABLE styles_mapping INTO ls_style_mapping WITH KEY guid = lv_style_guid.
*end of change issue #157 - allow column cellstyles
        lv_value = ls_style_mapping-style.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element_3->set_attribute_ns( name  = lc_xml_attr_s
                                        value = lv_value ).
      ENDIF.

      " For cells with formula ignore the value - Excel will calculate it
      IF <ls_sheet_content>-cell_formula IS NOT INITIAL.
        " fomula node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_f
                                                           parent = io_document ).
        lo_element_4->set_value( value = <ls_sheet_content>-cell_formula ).
        lo_element_3->append_child( new_child = lo_element_4 ). " formula node
      ELSEIF <ls_sheet_content>-column_formula_id <> 0.
        create_xl_sheet_column_formula(
          EXPORTING
            io_document             = io_document
            it_column_formulas      = io_worksheet->column_formulas
            is_sheet_content        = <ls_sheet_content>
          IMPORTING
            eo_element              = lo_element_4
          CHANGING
            ct_column_formulas_used = lt_column_formulas_used
            cv_si                   = lv_si ).
        lo_element_3->append_child( new_child = lo_element_4 ).
      ELSEIF <ls_sheet_content>-cell_value IS NOT INITIAL           "cell can have just style or formula
         AND <ls_sheet_content>-cell_value <> lc_dummy_cell_content.
        IF <ls_sheet_content>-data_type IS NOT INITIAL.
          IF <ls_sheet_content>-data_type EQ 's_leading_blanks'.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = 's' ).
          ELSE.
            lo_element_3->set_attribute_ns( name  = lc_xml_attr_t
                                            value = <ls_sheet_content>-data_type ).
          ENDIF.
        ENDIF.

        " value node
        lo_element_4 = io_document->create_simple_element( name   = lc_xml_node_v
                                                           parent = io_document ).

        IF <ls_sheet_content>-data_type EQ 's' OR <ls_sheet_content>-data_type EQ 's_leading_blanks'.
          lv_value = me->get_shared_string_index( ip_cell_value = <ls_sheet_content>-cell_value
                                                  it_rtf        = <ls_sheet_content>-rtf_tab ).
          CONDENSE lv_value.
          lo_element_4->set_value( value = lv_value ).
        ELSE.
          lv_value = <ls_sheet_content>-cell_value.
          CONDENSE lv_value.
          lo_element_4->set_value( value = lv_value ).
        ENDIF.

        lo_element_3->append_child( new_child = lo_element_4 ). " value node
      ENDIF.

      lo_element_2->append_child( new_child = lo_element_3 ). " column node
      ls_last_row = <ls_sheet_content>.
    ENDLOOP.
    IF sy-subrc = 0.
      READ TABLE lt_values INTO ls_values WITH KEY column = ls_last_row-cell_column.
      IF sy-subrc = 0 AND ls_values-value = ls_last_row-cell_value.
        CLEAR l_autofilter_hidden.
      ENDIF.
      IF lo_autofilter IS BOUND.
        IF ls_area-row_start >=  ls_last_row-cell_row OR " One less for header
          ls_area-row_end   < ls_last_row-cell_row .
          CLEAR l_autofilter_hidden.
        ENDIF.
      ELSE.
        CLEAR l_autofilter_hidden.
      ENDIF.
      " Row visibility of previos row.
      IF lo_row->get_visible( ) = abap_false OR
         l_autofilter_hidden = abap_true.
        lo_element_2->set_attribute_ns( name  = 'hidden' value = 'true' ).
      ENDIF.
      rv_ixml_sheet_data_root->append_child( new_child = lo_element_2 ). " row node
    ENDIF.
    DELETE io_worksheet->sheet_content WHERE cell_value = lc_dummy_cell_content.  " Get rid of dummyentries

  ENDMETHOD.


  METHOD create_xl_styles.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   dxfs-cellstyles are used in conditional formats:
*                CellIs, Expression, top10 ( forthcoming above average as well )
*                create own method to write dsfx-cellstyle to be reuseable by all these
*--------------------------------------------------------------------*


** Constant node name
    CONSTANTS: lc_xml_node_stylesheet        TYPE string VALUE 'styleSheet',
               " font
               lc_xml_node_fonts             TYPE string VALUE 'fonts',
               lc_xml_node_font              TYPE string VALUE 'font',
               lc_xml_node_color             TYPE string VALUE 'color',
               " fill
               lc_xml_node_fills             TYPE string VALUE 'fills',
               lc_xml_node_fill              TYPE string VALUE 'fill',
               lc_xml_node_patternfill       TYPE string VALUE 'patternFill',
               lc_xml_node_fgcolor           TYPE string VALUE 'fgColor',
               lc_xml_node_bgcolor           TYPE string VALUE 'bgColor',
               lc_xml_node_gradientfill      TYPE string VALUE 'gradientFill',
               lc_xml_node_stop              TYPE string VALUE 'stop',
               " borders
               lc_xml_node_borders           TYPE string VALUE 'borders',
               lc_xml_node_border            TYPE string VALUE 'border',
               lc_xml_node_left              TYPE string VALUE 'left',
               lc_xml_node_right             TYPE string VALUE 'right',
               lc_xml_node_top               TYPE string VALUE 'top',
               lc_xml_node_bottom            TYPE string VALUE 'bottom',
               lc_xml_node_diagonal          TYPE string VALUE 'diagonal',
               " numfmt
               lc_xml_node_numfmts           TYPE string VALUE 'numFmts',
               lc_xml_node_numfmt            TYPE string VALUE 'numFmt',
               " Styles
               lc_xml_node_cellstylexfs      TYPE string VALUE 'cellStyleXfs',
               lc_xml_node_xf                TYPE string VALUE 'xf',
               lc_xml_node_cellxfs           TYPE string VALUE 'cellXfs',
               lc_xml_node_cellstyles        TYPE string VALUE 'cellStyles',
               lc_xml_node_cellstyle         TYPE string VALUE 'cellStyle',
               lc_xml_node_dxfs              TYPE string VALUE 'dxfs',
               lc_xml_node_tablestyles       TYPE string VALUE 'tableStyles',
               " Colors
               lc_xml_node_colors            TYPE string VALUE 'colors',
               lc_xml_node_indexedcolors     TYPE string VALUE 'indexedColors',
               lc_xml_node_rgbcolor          TYPE string VALUE 'rgbColor',
               lc_xml_node_mrucolors         TYPE string VALUE 'mruColors',
               " Alignment
               lc_xml_node_alignment         TYPE string VALUE 'alignment',
               " Protection
               lc_xml_node_protection        TYPE string VALUE 'protection',
               " Node attributes
               lc_xml_attr_count             TYPE string VALUE 'count',
               lc_xml_attr_val               TYPE string VALUE 'val',
               lc_xml_attr_theme             TYPE string VALUE 'theme',
               lc_xml_attr_rgb               TYPE string VALUE 'rgb',
               lc_xml_attr_indexed           TYPE string VALUE 'indexed',
               lc_xml_attr_tint              TYPE string VALUE 'tint',
               lc_xml_attr_style             TYPE string VALUE 'style',
               lc_xml_attr_position          TYPE string VALUE 'position',
               lc_xml_attr_degree            TYPE string VALUE 'degree',
               lc_xml_attr_patterntype       TYPE string VALUE 'patternType',
               lc_xml_attr_numfmtid          TYPE string VALUE 'numFmtId',
               lc_xml_attr_fontid            TYPE string VALUE 'fontId',
               lc_xml_attr_fillid            TYPE string VALUE 'fillId',
               lc_xml_attr_borderid          TYPE string VALUE 'borderId',
               lc_xml_attr_xfid              TYPE string VALUE 'xfId',
               lc_xml_attr_applynumberformat TYPE string VALUE 'applyNumberFormat',
               lc_xml_attr_applyprotection   TYPE string VALUE 'applyProtection',
               lc_xml_attr_applyfont         TYPE string VALUE 'applyFont',
               lc_xml_attr_applyfill         TYPE string VALUE 'applyFill',
               lc_xml_attr_applyborder       TYPE string VALUE 'applyBorder',
               lc_xml_attr_name              TYPE string VALUE 'name',
               lc_xml_attr_builtinid         TYPE string VALUE 'builtinId',
               lc_xml_attr_defaulttablestyle TYPE string VALUE 'defaultTableStyle',
               lc_xml_attr_defaultpivotstyle TYPE string VALUE 'defaultPivotStyle',
               lc_xml_attr_applyalignment    TYPE string VALUE 'applyAlignment',
               lc_xml_attr_horizontal        TYPE string VALUE 'horizontal',
               lc_xml_attr_formatcode        TYPE string VALUE 'formatCode',
               lc_xml_attr_vertical          TYPE string VALUE 'vertical',
               lc_xml_attr_wraptext          TYPE string VALUE 'wrapText',
               lc_xml_attr_textrotation      TYPE string VALUE 'textRotation',
               lc_xml_attr_shrinktofit       TYPE string VALUE 'shrinkToFit',
               lc_xml_attr_indent            TYPE string VALUE 'indent',
               lc_xml_attr_locked            TYPE string VALUE 'locked',
               lc_xml_attr_hidden            TYPE string VALUE 'hidden',
               lc_xml_attr_diagonalup        TYPE string VALUE 'diagonalUp',
               lc_xml_attr_diagonaldown      TYPE string VALUE 'diagonalDown',
               " Node namespace
               lc_xml_node_ns                TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               lc_xml_attr_type              TYPE string VALUE 'type',
               lc_xml_attr_bottom            TYPE string VALUE 'bottom',
               lc_xml_attr_top               TYPE string VALUE 'top',
               lc_xml_attr_right             TYPE string VALUE 'right',
               lc_xml_attr_left              TYPE string VALUE 'left'.

    DATA: lo_document        TYPE REF TO if_ixml_document,
          lo_element_root    TYPE REF TO if_ixml_element,
          lo_element_fonts   TYPE REF TO if_ixml_element,
          lo_element_font    TYPE REF TO if_ixml_element,
          lo_element_fills   TYPE REF TO if_ixml_element,
          lo_element_fill    TYPE REF TO if_ixml_element,
          lo_element_borders TYPE REF TO if_ixml_element,
          lo_element_border  TYPE REF TO if_ixml_element,
          lo_element_numfmts TYPE REF TO if_ixml_element,
          lo_element_numfmt  TYPE REF TO if_ixml_element,
          lo_element_cellxfs TYPE REF TO if_ixml_element,
          lo_element         TYPE REF TO if_ixml_element,
          lo_sub_element     TYPE REF TO if_ixml_element,
          lo_sub_element_2   TYPE REF TO if_ixml_element,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lo_iterator2       TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lo_style_cond      TYPE REF TO zcl_excel_style_cond,
          lo_style           TYPE REF TO zcl_excel_style.


    DATA: lt_fonts          TYPE zexcel_t_style_font,
          ls_font           TYPE zexcel_s_style_font,
          lt_fills          TYPE zexcel_t_style_fill,
          ls_fill           TYPE zexcel_s_style_fill,
          lt_borders        TYPE zexcel_t_style_border,
          ls_border         TYPE zexcel_s_style_border,
          lt_numfmts        TYPE zexcel_t_style_numfmt,
          ls_numfmt         TYPE zexcel_s_style_numfmt,
          lt_protections    TYPE zexcel_t_style_protection,
          ls_protection     TYPE zexcel_s_style_protection,
          lt_alignments     TYPE zexcel_t_style_alignment,
          ls_alignment      TYPE zexcel_s_style_alignment,
          lt_cellxfs        TYPE zexcel_t_cellxfs,
          ls_cellxfs        TYPE zexcel_s_cellxfs,
          ls_styles_mapping TYPE zexcel_s_styles_mapping,
          lt_colors         TYPE zexcel_t_style_color_argb,
          ls_color          LIKE LINE OF lt_colors.

    DATA: lv_value         TYPE string,
          lv_dfx_count     TYPE i,
          lv_fonts_count   TYPE i,
          lv_fills_count   TYPE i,
          lv_borders_count TYPE i,
          lv_cellxfs_count TYPE i.

    TYPES: BEGIN OF ts_built_in_format,
             num_format TYPE zexcel_number_format,
             id         TYPE i,
           END OF ts_built_in_format.

    DATA: lt_built_in_num_formats TYPE HASHED TABLE OF ts_built_in_format WITH UNIQUE KEY num_format,
          ls_built_in_num_format  LIKE LINE OF lt_built_in_num_formats.
    FIELD-SYMBOLS: <ls_built_in_format> LIKE LINE OF lt_built_in_num_formats,
                   <ls_reader_built_in> LIKE LINE OF zcl_excel_style_number_format=>mt_built_in_num_formats.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

***********************************************************************
* STEP 3: Create main node relationships
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_stylesheet
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).

**********************************************************************
* STEP 4: Create subnodes

    lo_element_fonts = lo_document->create_simple_element( name   = lc_xml_node_fonts
                                                           parent = lo_document ).

    lo_element_fills = lo_document->create_simple_element( name   = lc_xml_node_fills
                                                           parent = lo_document ).

    lo_element_borders = lo_document->create_simple_element( name   = lc_xml_node_borders
                                                             parent = lo_document ).

    lo_element_cellxfs = lo_document->create_simple_element( name   = lc_xml_node_cellxfs
                                                             parent = lo_document ).

    lo_element_numfmts = lo_document->create_simple_element( name   = lc_xml_node_numfmts
                                                             parent = lo_document ).

* Prepare built-in number formats.
    LOOP AT zcl_excel_style_number_format=>mt_built_in_num_formats ASSIGNING <ls_reader_built_in>.
      ls_built_in_num_format-id         = <ls_reader_built_in>-id.
      ls_built_in_num_format-num_format = <ls_reader_built_in>-format->format_code.
      INSERT ls_built_in_num_format INTO TABLE lt_built_in_num_formats.
    ENDLOOP.
* Compress styles
    lo_iterator = excel->get_styles_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_style ?= lo_iterator->get_next( ).
      ls_font       = lo_style->font->get_structure( ).
      ls_fill       = lo_style->fill->get_structure( ).
      ls_border     = lo_style->borders->get_structure( ).
      ls_alignment  = lo_style->alignment->get_structure( ).
      ls_protection = lo_style->protection->get_structure( ).
      ls_numfmt     = lo_style->number_format->get_structure( ).

      CLEAR ls_cellxfs.


* Compress fonts
      READ TABLE lt_fonts FROM ls_font TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-fontid = sy-tabix.
      ELSE.
        APPEND ls_font TO lt_fonts.
        DESCRIBE TABLE lt_fonts LINES ls_cellxfs-fontid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-fontid.

* Compress alignment
      READ TABLE lt_alignments FROM ls_alignment TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-alignmentid = sy-tabix.
      ELSE.
        APPEND ls_alignment TO lt_alignments.
        DESCRIBE TABLE lt_alignments LINES ls_cellxfs-alignmentid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-alignmentid.

* Compress fills
      READ TABLE lt_fills FROM ls_fill TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-fillid = sy-tabix.
      ELSE.
        APPEND ls_fill TO lt_fills.
        DESCRIBE TABLE lt_fills LINES ls_cellxfs-fillid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-fillid.

* Compress borders
      READ TABLE lt_borders FROM ls_border TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_cellxfs-borderid = sy-tabix.
      ELSE.
        APPEND ls_border TO lt_borders.
        DESCRIBE TABLE lt_borders LINES ls_cellxfs-borderid.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-borderid.

* Compress protection
      IF ls_protection-locked EQ c_on AND ls_protection-hidden EQ c_off.
        ls_cellxfs-applyprotection    = 0.
      ELSE.
        READ TABLE lt_protections FROM ls_protection TRANSPORTING NO FIELDS.
        IF sy-subrc EQ 0.
          ls_cellxfs-protectionid = sy-tabix.
        ELSE.
          APPEND ls_protection TO lt_protections.
          DESCRIBE TABLE lt_protections LINES ls_cellxfs-protectionid.
        ENDIF.
        ls_cellxfs-applyprotection    = 1.
      ENDIF.
      SUBTRACT 1 FROM ls_cellxfs-protectionid.

* Compress number formats

      "-----------
      IF ls_numfmt-numfmt NE zcl_excel_style_number_format=>c_format_date_std." and ls_numfmt-NUMFMT ne 'STD_NDEC'. " ALE Changes on going
        "---
        IF ls_numfmt IS NOT INITIAL.
* issue  #389 - Problem with built-in format ( those are not being taken account of )
* There are some internal number formats built-in into EXCEL
* Use these instead of duplicating the entries here, since they seem to be language-dependant and adjust to user settings in excel
          READ TABLE lt_built_in_num_formats ASSIGNING <ls_built_in_format> WITH TABLE KEY num_format = ls_numfmt-numfmt.
          IF sy-subrc = 0.
            ls_cellxfs-numfmtid = <ls_built_in_format>-id.
          ELSE.
            READ TABLE lt_numfmts FROM ls_numfmt TRANSPORTING NO FIELDS.
            IF sy-subrc EQ 0.
              ls_cellxfs-numfmtid = sy-tabix.
            ELSE.
              APPEND ls_numfmt TO lt_numfmts.
              DESCRIBE TABLE lt_numfmts LINES ls_cellxfs-numfmtid.
            ENDIF.
            ADD zcl_excel_common=>c_excel_numfmt_offset TO ls_cellxfs-numfmtid. " Add OXML offset for custom styles
          ENDIF.
          ls_cellxfs-applynumberformat    = 1.
        ELSE.
          ls_cellxfs-applynumberformat    = 0.
        ENDIF.
        "----------- " ALE changes on going
      ELSE.
        ls_cellxfs-applynumberformat    = 1.
        IF ls_numfmt-numfmt EQ zcl_excel_style_number_format=>c_format_date_std.
          ls_cellxfs-numfmtid = 14.
        ENDIF.
      ENDIF.
      "---

      IF ls_cellxfs-fontid NE 0.
        ls_cellxfs-applyfont    = 1.
      ELSE.
        ls_cellxfs-applyfont    = 0.
      ENDIF.
      IF ls_cellxfs-alignmentid NE 0.
        ls_cellxfs-applyalignment = 1.
      ELSE.
        ls_cellxfs-applyalignment = 0.
      ENDIF.
      IF ls_cellxfs-fillid NE 0.
        ls_cellxfs-applyfill    = 1.
      ELSE.
        ls_cellxfs-applyfill    = 0.
      ENDIF.
      IF ls_cellxfs-borderid NE 0.
        ls_cellxfs-applyborder    = 1.
      ELSE.
        ls_cellxfs-applyborder    = 0.
      ENDIF.

* Remap styles
      READ TABLE lt_cellxfs FROM ls_cellxfs TRANSPORTING NO FIELDS.
      IF sy-subrc EQ 0.
        ls_styles_mapping-style = sy-tabix.
      ELSE.
        APPEND ls_cellxfs TO lt_cellxfs.
        DESCRIBE TABLE lt_cellxfs LINES ls_styles_mapping-style.
      ENDIF.
      SUBTRACT 1 FROM ls_styles_mapping-style.
      ls_styles_mapping-guid = lo_style->get_guid( ).
      APPEND ls_styles_mapping TO me->styles_mapping.
    ENDWHILE.

    " create numfmt elements
    LOOP AT lt_numfmts INTO ls_numfmt.
      lo_element_numfmt = lo_document->create_simple_element( name   = lc_xml_node_numfmt
                                                              parent = lo_document ).
      lv_value = sy-tabix + zcl_excel_common=>c_excel_numfmt_offset.
      CONDENSE lv_value.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                        value = lv_value ).
      lv_value = ls_numfmt-numfmt.
      lo_element_numfmt->set_attribute_ns( name  = lc_xml_attr_formatcode
                                           value = lv_value ).
      lo_element_numfmts->append_child( new_child = lo_element_numfmt ).
    ENDLOOP.

    " create font elements
    LOOP AT lt_fonts INTO ls_font.
      lo_element_font = lo_document->create_simple_element( name   = lc_xml_node_font
                                                            parent = lo_document ).
      create_xl_styles_font_node( io_document = lo_document
                                  io_parent   = lo_element_font
                                  is_font     = ls_font ).
      lo_element_fonts->append_child( new_child = lo_element_font ).
    ENDLOOP.

    " create fill elements
    LOOP AT lt_fills INTO ls_fill.
      lo_element_fill = lo_document->create_simple_element( name   = lc_xml_node_fill
                                                            parent = lo_document ).

      IF ls_fill-gradtype IS NOT INITIAL.
        "gradient

        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_gradientfill
                                                            parent = lo_document ).
        IF ls_fill-gradtype-degree IS NOT INITIAL.
          lv_value = ls_fill-gradtype-degree.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_degree  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-type IS NOT INITIAL.
          lv_value = ls_fill-gradtype-type.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_type  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-bottom IS NOT INITIAL.
          lv_value = ls_fill-gradtype-bottom.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_bottom  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-top IS NOT INITIAL.
          lv_value = ls_fill-gradtype-top.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_top  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-right IS NOT INITIAL.
          lv_value = ls_fill-gradtype-right.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_right  value = lv_value ).
        ENDIF.
        IF ls_fill-gradtype-left IS NOT INITIAL.
          lv_value = ls_fill-gradtype-left.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_left  value = lv_value ).
        ENDIF.

        IF ls_fill-gradtype-position3 IS NOT INITIAL.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.

          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position3.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

        ELSE.
          "create <stop> elements for gradients, we can have 2 or 3 stops in each gradient
          lo_sub_element_2 =  lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                  parent = lo_sub_element ).
          lv_value = ls_fill-gradtype-position1.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).

          lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_stop
                                                                 parent = lo_sub_element ).

          lv_value = ls_fill-gradtype-position2.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_position
                                              value = lv_value ).

          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element_2
              is_color           = ls_fill-fgcolor
              iv_color_elem_name = lc_xml_node_color ).
          lo_sub_element->append_child( new_child = lo_sub_element_2 ).
        ENDIF.

      ELSE.
        "pattern
        lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_patternfill
                                                             parent = lo_document ).
        lv_value = ls_fill-filltype.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_patterntype
                                          value = lv_value ).
        " fgcolor
        create_xl_styles_color_node(
            io_document        = lo_document
            io_parent          = lo_sub_element
            is_color           = ls_fill-fgcolor
            iv_color_elem_name = lc_xml_node_fgcolor ).

        IF  ls_fill-fgcolor-rgb IS INITIAL AND
            ls_fill-fgcolor-indexed EQ zcl_excel_style_color=>c_indexed_not_set AND
            ls_fill-fgcolor-theme EQ zcl_excel_style_color=>c_theme_not_set AND
            ls_fill-fgcolor-tint IS INITIAL AND ls_fill-bgcolor-indexed EQ zcl_excel_style_color=>c_indexed_sys_foreground.

          " bgcolor
          create_xl_styles_color_node(
              io_document        = lo_document
              io_parent          = lo_sub_element
              is_color           = ls_fill-bgcolor
              iv_color_elem_name = lc_xml_node_bgcolor ).

        ENDIF.
      ENDIF.

      lo_element_fill->append_child( new_child = lo_sub_element )."pattern
      lo_element_fills->append_child( new_child = lo_element_fill ).
    ENDLOOP.

    " create border elements
    LOOP AT lt_borders INTO ls_border.
      lo_element_border = lo_document->create_simple_element( name   = lc_xml_node_border
                                                              parent = lo_document ).

      IF ls_border-diagonalup IS NOT INITIAL.
        lv_value = ls_border-diagonalup.
        CONDENSE lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonalup
                                          value = lv_value ).
      ENDIF.

      IF ls_border-diagonaldown IS NOT INITIAL.
        lv_value = ls_border-diagonaldown.
        CONDENSE lv_value.
        lo_element_border->set_attribute_ns( name  = lc_xml_attr_diagonaldown
                                          value = lv_value ).
      ENDIF.

      "left
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_left
                                                           parent = lo_document ).
      IF ls_border-left_style IS NOT INITIAL.
        lv_value = ls_border-left_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-left_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "right
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_right
                                                           parent = lo_document ).
      IF ls_border-right_style IS NOT INITIAL.
        lv_value = ls_border-right_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-right_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "top
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_top
                                                           parent = lo_document ).
      IF ls_border-top_style IS NOT INITIAL.
        lv_value = ls_border-top_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-top_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "bottom
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_bottom
                                                           parent = lo_document ).
      IF ls_border-bottom_style IS NOT INITIAL.
        lv_value = ls_border-bottom_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-bottom_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).

      "diagonal
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_diagonal
                                                           parent = lo_document ).
      IF ls_border-diagonal_style IS NOT INITIAL.
        lv_value = ls_border-diagonal_style.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_style
                                          value = lv_value ).
      ENDIF.

      create_xl_styles_color_node(
          io_document        = lo_document
          io_parent          = lo_sub_element
          is_color           = ls_border-diagonal_color ).

      lo_element_border->append_child( new_child = lo_sub_element ).
      lo_element_borders->append_child( new_child = lo_element_border ).
    ENDLOOP.

    " update attribute "count"
    DESCRIBE TABLE lt_fonts LINES lv_fonts_count.
    lv_value = lv_fonts_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_fonts->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    DESCRIBE TABLE lt_fills LINES lv_fills_count.
    lv_value = lv_fills_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_fills->set_attribute_ns( name  = lc_xml_attr_count
                                        value = lv_value ).
    DESCRIBE TABLE lt_borders LINES lv_borders_count.
    lv_value = lv_borders_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_borders->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).
    DESCRIBE TABLE lt_cellxfs LINES lv_cellxfs_count.
    lv_value = lv_cellxfs_count.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_element_cellxfs->set_attribute_ns( name  = lc_xml_attr_count
                                          value = lv_value ).

    " Append to root node
    lo_element_root->append_child( new_child = lo_element_numfmts ).
    lo_element_root->append_child( new_child = lo_element_fonts ).
    lo_element_root->append_child( new_child = lo_element_fills ).
    lo_element_root->append_child( new_child = lo_element_borders ).

    " cellstylexfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstylexfs
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    LOOP AT lt_cellxfs INTO ls_cellxfs.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_xf
                                                          parent = lo_document ).
      lv_value = ls_cellxfs-numfmtid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_numfmtid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fontid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fontid
                                    value = lv_value ).
      lv_value = ls_cellxfs-fillid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_fillid
                                    value = lv_value ).
      lv_value = ls_cellxfs-borderid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_borderid
                                    value = lv_value ).
      lv_value = ls_cellxfs-xfid.
      SHIFT lv_value RIGHT DELETING TRAILING space.
      SHIFT lv_value LEFT DELETING LEADING space.
      lo_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                    value = lv_value ).
      IF ls_cellxfs-applynumberformat EQ 1.
        lv_value = ls_cellxfs-applynumberformat.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applynumberformat
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyfont EQ 1.
        lv_value = ls_cellxfs-applyfont.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfont
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyfill EQ 1.
        lv_value = ls_cellxfs-applyfill.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyfill
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyborder EQ 1.
        lv_value = ls_cellxfs-applyborder.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyborder
                                      value = lv_value ).
      ENDIF.
      IF ls_cellxfs-applyalignment EQ 1. " depends on each style not for all the sheet
        lv_value = ls_cellxfs-applyalignment.
        SHIFT lv_value RIGHT DELETING TRAILING space.
        SHIFT lv_value LEFT DELETING LEADING space.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyalignment
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_alignment
                                                               parent = lo_document ).
        ADD 1 TO ls_cellxfs-alignmentid. "Table index starts from 1
        READ TABLE lt_alignments INTO ls_alignment INDEX ls_cellxfs-alignmentid.
        SUBTRACT 1 FROM ls_cellxfs-alignmentid.
        IF ls_alignment-horizontal IS NOT INITIAL.
          lv_value = ls_alignment-horizontal.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_horizontal
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-vertical IS NOT INITIAL.
          lv_value = ls_alignment-vertical.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_vertical
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-wraptext EQ abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_wraptext
                                              value = c_on ).
        ENDIF.
        IF ls_alignment-textrotation IS NOT INITIAL.
          lv_value = ls_alignment-textrotation.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_textrotation
                                              value = lv_value ).
        ENDIF.
        IF ls_alignment-shrinktofit EQ abap_true.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_shrinktofit
                                              value = c_on ).
        ENDIF.
        IF ls_alignment-indent IS NOT INITIAL.
          lv_value = ls_alignment-indent.
          SHIFT lv_value RIGHT DELETING TRAILING space.
          SHIFT lv_value LEFT DELETING LEADING space.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_indent
                                              value = lv_value ).
        ENDIF.

        lo_element->append_child( new_child = lo_sub_element_2 ).
      ENDIF.
      IF ls_cellxfs-applyprotection EQ 1.
        lv_value = ls_cellxfs-applyprotection.
        CONDENSE lv_value NO-GAPS.
        lo_element->set_attribute_ns( name  = lc_xml_attr_applyprotection
                                      value = lv_value ).
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_protection
                                                               parent = lo_document ).
        ADD 1 TO ls_cellxfs-protectionid. "Table index starts from 1
        READ TABLE lt_protections INTO ls_protection INDEX ls_cellxfs-protectionid.
        SUBTRACT 1 FROM ls_cellxfs-protectionid.
        IF ls_protection-locked IS NOT INITIAL.
          lv_value = ls_protection-locked.
          CONDENSE lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_locked
                                              value = lv_value ).
        ENDIF.
        IF ls_protection-hidden IS NOT INITIAL.
          lv_value = ls_protection-hidden.
          CONDENSE lv_value.
          lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_hidden
                                              value = lv_value ).
        ENDIF.
        lo_element->append_child( new_child = lo_sub_element_2 ).
      ENDIF.
      lo_element_cellxfs->append_child( new_child = lo_element ).
    ENDLOOP.

    lo_element_root->append_child( new_child = lo_element_cellxfs ).

    " cellStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '1' ).
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_cellstyle
                                                         parent = lo_document ).

    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                      value = 'Normal' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xfid
                                      value = c_off ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_builtinid
                                      value = c_off ).

    lo_element->append_child( new_child = lo_sub_element ).
    lo_element_root->append_child( new_child = lo_element ).

    " dxfs node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_dxfs
                                                     parent = lo_document ).

    lo_iterator = me->excel->get_worksheets_iterator( ).
    " get sheets
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      " Conditional formatting styles into exch sheet
      lo_iterator2 = lo_worksheet->get_style_cond_iterator( ).
      WHILE lo_iterator2->has_next( ) EQ abap_true.
        lo_style_cond ?= lo_iterator2->get_next( ).
        CASE lo_style_cond->rule.
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method
          WHEN zcl_excel_style_cond=>c_rule_cellis.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_cellis-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).

          WHEN zcl_excel_style_cond=>c_rule_expression.
            me->create_dxf_style( EXPORTING
                          iv_cell_style    = lo_style_cond->mode_expression-cell_style
                          io_dxf_element   = lo_element
                          io_ixml_document = lo_document
                          it_cellxfs       = lt_cellxfs
                          it_fonts         = lt_fonts
                          it_fills         = lt_fills
                        CHANGING
                          cv_dfx_count     = lv_dfx_count ).



          WHEN zcl_excel_style_cond=>c_rule_top10.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_top10-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).

          WHEN zcl_excel_style_cond=>c_rule_above_average.
            me->create_dxf_style( EXPORTING
                                    iv_cell_style    = lo_style_cond->mode_above_average-cell_style
                                    io_dxf_element   = lo_element
                                    io_ixml_document = lo_document
                                    it_cellxfs       = lt_cellxfs
                                    it_fonts         = lt_fonts
                                    it_fills         = lt_fills
                                  CHANGING
                                    cv_dfx_count     = lv_dfx_count ).
* begin of change issue #366 - missing conditional rules: top10, move dfx-styles to own method

          WHEN OTHERS.
            CONTINUE.
        ENDCASE.
      ENDWHILE.
    ENDWHILE.

    lv_value = lv_dfx_count.
    CONDENSE lv_value.
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = lv_value ).
    lo_element_root->append_child( new_child = lo_element ).

    " tableStyles node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_tablestyles
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_count
                                  value = '0' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaulttablestyle
                                  value = zcl_excel_table=>builtinstyle_medium9 ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_defaultpivotstyle
                                  value = zcl_excel_table=>builtinstyle_pivot_light16 ).
    lo_element_root->append_child( new_child = lo_element ).

    "write legacy color palette in case any indexed color was changed
    IF excel->legacy_palette->is_modified( ) = abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_colors
                                                     parent   = lo_document ).
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_indexedcolors
                                                         parent   = lo_document ).
      lo_element->append_child( new_child = lo_sub_element ).

      lt_colors = excel->legacy_palette->get_colors( ).
      LOOP AT lt_colors INTO ls_color.
        lo_sub_element_2 = lo_document->create_simple_element( name   = lc_xml_node_rgbcolor
                                                               parent = lo_document ).
        lv_value = ls_color.
        lo_sub_element_2->set_attribute_ns( name  = lc_xml_attr_rgb
                                            value = lv_value ).
        lo_sub_element->append_child( new_child = lo_sub_element_2 ).
      ENDLOOP.

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_styles_color_node.
    DATA: lo_sub_element TYPE REF TO if_ixml_element,
          lv_value       TYPE string.

    CONSTANTS: lc_xml_attr_theme   TYPE string VALUE 'theme',
               lc_xml_attr_rgb     TYPE string VALUE 'rgb',
               lc_xml_attr_indexed TYPE string VALUE 'indexed',
               lc_xml_attr_tint    TYPE string VALUE 'tint'.

    "add node only if at least one attribute is set
    CHECK is_color-rgb IS NOT INITIAL OR
          is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set OR
          is_color-theme <> zcl_excel_style_color=>c_theme_not_set OR
          is_color-tint IS NOT INITIAL.

    lo_sub_element = io_document->create_simple_element(
        name      = iv_color_elem_name
        parent    = io_parent ).

    IF is_color-rgb IS NOT INITIAL.
      lv_value = is_color-rgb.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_rgb
                                        value = lv_value ).
    ENDIF.

    IF is_color-indexed <> zcl_excel_style_color=>c_indexed_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-indexed ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_indexed
                                        value = lv_value ).
    ENDIF.

    IF is_color-theme <> zcl_excel_style_color=>c_theme_not_set.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-theme ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_theme
                                        value = lv_value ).
    ENDIF.

    IF is_color-tint IS NOT INITIAL.
      lv_value = zcl_excel_common=>number_to_excel_string( is_color-tint ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_tint
                                        value = lv_value ).
    ENDIF.

    io_parent->append_child( new_child = lo_sub_element ).
  ENDMETHOD.


  METHOD create_xl_styles_font_node.

    CONSTANTS: lc_xml_node_b      TYPE string VALUE 'b',            "bold
               lc_xml_node_i      TYPE string VALUE 'i',            "italic
               lc_xml_node_u      TYPE string VALUE 'u',            "underline
               lc_xml_node_strike TYPE string VALUE 'strike',       "strikethrough
               lc_xml_node_sz     TYPE string VALUE 'sz',
               lc_xml_node_name   TYPE string VALUE 'name',
               lc_xml_node_rfont  TYPE string VALUE 'rFont',
               lc_xml_node_family TYPE string VALUE 'family',
               lc_xml_node_scheme TYPE string VALUE 'scheme',
               lc_xml_attr_val    TYPE string VALUE 'val'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_font TYPE REF TO if_ixml_element,
          ls_font         TYPE zexcel_s_style_font,
          lo_sub_element  TYPE REF TO if_ixml_element,
          lv_value        TYPE string.

    lo_document = io_document.
    lo_element_font = io_parent.
    ls_font = is_font.

    IF ls_font-bold EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_b
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-italic EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_i
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-underline EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_u
                                                           parent = lo_document ).
      lv_value = ls_font-underline_mode.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    IF ls_font-strikethrough EQ abap_true.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_strike
                                                           parent = lo_document ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.
    "size
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_sz
                                                         parent = lo_document ).
    lv_value = ls_font-size.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "color
    create_xl_styles_color_node(
        io_document        = lo_document
        io_parent          = lo_element_font
        is_color           = ls_font-color ).

    "name
    IF iv_use_rtf = abap_false.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_name
                                                           parent = lo_document ).
    ELSE.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_rfont
                                                           parent = lo_document ).
    ENDIF.
    lv_value = ls_font-name.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "family
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_family
                                                         parent = lo_document ).
    lv_value = ls_font-family.
    SHIFT lv_value RIGHT DELETING TRAILING space.
    SHIFT lv_value LEFT DELETING LEADING space.
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                      value = lv_value ).
    lo_element_font->append_child( new_child = lo_sub_element ).
    "scheme
    IF ls_font-scheme IS NOT INITIAL.
      lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_scheme
                                                           parent = lo_document ).
      lv_value = ls_font-scheme.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_val
                                        value = lv_value ).
      lo_element_font->append_child( new_child = lo_sub_element ).
    ENDIF.

  ENDMETHOD.


  METHOD create_xl_table.

    DATA: lc_xml_node_table        TYPE string VALUE 'table',
          lc_xml_node_relationship TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id           TYPE string VALUE 'id',
          lc_xml_attr_name         TYPE string VALUE 'name',
          lc_xml_attr_display_name TYPE string VALUE 'displayName',
          lc_xml_attr_ref          TYPE string VALUE 'ref',
          lc_xml_attr_totals       TYPE string VALUE 'totalsRowShown',
          " Node namespace
          lc_xml_node_table_ns     TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          " Node id
          lc_xml_node_ridx_id      TYPE string VALUE 'rId#'.

    DATA: lo_document     TYPE REF TO if_ixml_document,
          lo_element_root TYPE REF TO if_ixml_element,
          lo_element      TYPE REF TO if_ixml_element,
          lo_element2     TYPE REF TO if_ixml_element,
          lo_element3     TYPE REF TO if_ixml_element,
          lv_table_name   TYPE string,
          lv_id           TYPE i,
          lv_match        TYPE i,
          lv_ref          TYPE string,
          lv_value        TYPE string,
          lv_num_columns  TYPE i,
          ls_fieldcat     TYPE zexcel_s_fieldcatalog.


**********************************************************************
* STEP 1: Create xml
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node table
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_table
                                                           parent = lo_document ).

    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_table_ns  ).

    lv_id = io_table->get_id( ).
    lv_value = zcl_excel_common=>number_to_excel_string( ip_value = lv_id ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_id
                                       value = lv_value ).

    FIND ALL OCCURRENCES OF REGEX '[^_a-zA-Z0-9]' IN io_table->settings-table_name IGNORING CASE MATCH COUNT lv_match.
    IF io_table->settings-table_name IS NOT INITIAL AND lv_match EQ 0.
      " Name rules (https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114)
      "   - You can't use "C", "c", "R", or "r" for the name, because they're already designated as a shortcut for selecting the column or row for the active cell when you enter them in the Name or Go To box.
      "   - Don't use cell references — Names can't be the same as a cell reference, such as Z$100 or R1C1
      IF ( strlen( io_table->settings-table_name ) = 1 AND io_table->settings-table_name CO 'CcRr' )
         OR zcl_excel_common=>shift_formula(
              iv_reference_formula = io_table->settings-table_name
              iv_shift_cols        = 0
              iv_shift_rows        = 1 ) <> io_table->settings-table_name.
        lv_table_name = io_table->get_name( ).
      ELSE.
        lv_table_name = io_table->settings-table_name.
      ENDIF.
    ELSE.
      lv_table_name = io_table->get_name( ).
    ENDIF.
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_name
                                       value = lv_table_name ).

    lo_element_root->set_attribute_ns( name  = lc_xml_attr_display_name
                                       value = lv_table_name ).

    lv_ref = io_table->get_reference( ).
    lo_element_root->set_attribute_ns( name  = lc_xml_attr_ref
                                       value = lv_ref ).
    IF io_table->has_totals( ) = abap_true.
      lo_element_root->set_attribute_ns( name  = 'totalsRowCount'
                                             value = '1' ).
    ELSE.
      lo_element_root->set_attribute_ns( name  = lc_xml_attr_totals
                                           value = '0' ).
    ENDIF.

**********************************************************************
* STEP 4: Create subnodes

    " autoFilter
    IF io_table->settings-nofilters EQ abap_false.
      lo_element = lo_document->create_simple_element( name   = 'autoFilter'
                                                       parent = lo_document ).

      lv_ref = io_table->get_reference( ip_include_totals_row = abap_false ).
      lo_element->set_attribute_ns( name  = 'ref'
                                    value = lv_ref ).

      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    "columns
    lo_element = lo_document->create_simple_element( name   = 'tableColumns'
                                                     parent = lo_document ).

    LOOP AT io_table->fieldcat INTO ls_fieldcat WHERE dynpfld = abap_true.
      ADD 1 TO lv_num_columns.
    ENDLOOP.

    lv_value = lv_num_columns.
    CONDENSE lv_value.
    lo_element->set_attribute_ns( name  = 'count'
                                  value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).

    LOOP AT io_table->fieldcat INTO ls_fieldcat WHERE dynpfld = abap_true.
      lo_element2 = lo_document->create_simple_element_ns( name   = 'tableColumn'
                                                                  parent = lo_element ).

      lv_value = ls_fieldcat-position.
      SHIFT lv_value LEFT DELETING LEADING '0'.
      lo_element2->set_attribute_ns( name  = 'id'
                                    value = lv_value ).
      lv_value = ls_fieldcat-scrtext_l.

      " The text "_x...._", with "_x" not "_X", with exactly 4 ".", each being 0-9 a-f or A-F (case insensitive), is interpreted
      " like Unicode character U+.... (e.g. "_x0041_" is rendered like "A") is for characters.
      " To not interpret it, Excel replaces the first "_" is to be replaced with "_x005f_".
      IF lv_value CS '_x'.
        REPLACE ALL OCCURRENCES OF REGEX '_(x[0-9a-fA-F]{4}_)' IN lv_value WITH '_x005f_$1' RESPECTING CASE.
      ENDIF.

      " XML chapter 2.2: Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
      " NB: although Excel supports _x0009_, it's not rendered except if you edit the text.
      " Excel considers _x000d_ as being an error (_x000a_ is sufficient and rendered).
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN lv_value WITH '_x000a_'.
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf(1) IN lv_value WITH ``.
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>horizontal_tab IN lv_value WITH '_x0009_'.

      lo_element2->set_attribute_ns( name  = 'name'
                                    value = lv_value ).

      IF ls_fieldcat-totals_function IS NOT INITIAL.
        lo_element2->set_attribute_ns( name  = 'totalsRowFunction'
                                          value = ls_fieldcat-totals_function ).
      ENDIF.

      IF ls_fieldcat-column_formula IS NOT INITIAL.
        lv_value = ls_fieldcat-column_formula.
        CONDENSE lv_value.
        lo_element3 = lo_document->create_simple_element_ns( name   = 'calculatedColumnFormula'
                                                             parent = lo_element2 ).
        lo_element3->set_value( lv_value ).
        lo_element2->append_child( new_child = lo_element3 ).
      ENDIF.

      lo_element->append_child( new_child = lo_element2 ).
    ENDLOOP.


    lo_element  = lo_document->create_simple_element( name   = 'tableStyleInfo'
                                                           parent = lo_element_root ).

    lo_element->set_attribute_ns( name  = 'name'
                                       value = io_table->settings-table_style  ).

    lo_element->set_attribute_ns( name  = 'showFirstColumn'
                                       value = '0' ).

    lo_element->set_attribute_ns( name  = 'showLastColumn'
                                       value = '0' ).

    IF io_table->settings-show_row_stripes = abap_true.
      lv_value = '1'.
    ELSE.
      lv_value = '0'.
    ENDIF.

    lo_element->set_attribute_ns( name  = 'showRowStripes'
                                       value = lv_value ).

    IF io_table->settings-show_column_stripes = abap_true.
      lv_value = '1'.
    ELSE.
      lv_value = '0'.
    ENDIF.

    lo_element->set_attribute_ns( name  = 'showColumnStripes'
                                       value = lv_value ).

    lo_element_root->append_child( new_child = lo_element ).
**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xl_theme.
    DATA: lo_theme TYPE REF TO zcl_excel_theme.

    excel->get_theme(
    IMPORTING
      eo_theme = lo_theme
      ).
    IF lo_theme IS INITIAL.
      CREATE OBJECT lo_theme.
    ENDIF.
    ep_content = lo_theme->write_theme( ).

  ENDMETHOD.


  METHOD create_xl_workbook.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-01
* changes:  correction of pointer to localSheetId
*--------------------------------------------------------------------*

** Constant node name
    DATA: lc_xml_node_workbook           TYPE string VALUE 'workbook',
          lc_xml_node_fileversion        TYPE string VALUE 'fileVersion',
          lc_xml_node_workbookpr         TYPE string VALUE 'workbookPr',
          lc_xml_node_bookviews          TYPE string VALUE 'bookViews',
          lc_xml_node_workbookview       TYPE string VALUE 'workbookView',
          lc_xml_node_sheets             TYPE string VALUE 'sheets',
          lc_xml_node_sheet              TYPE string VALUE 'sheet',
          lc_xml_node_calcpr             TYPE string VALUE 'calcPr',
          lc_xml_node_workbookprotection TYPE string VALUE 'workbookProtection',
          lc_xml_node_definednames       TYPE string VALUE 'definedNames',
          lc_xml_node_definedname        TYPE string VALUE 'definedName',
          " Node attributes
          lc_xml_attr_appname            TYPE string VALUE 'appName',
          lc_xml_attr_lastedited         TYPE string VALUE 'lastEdited',
          lc_xml_attr_lowestedited       TYPE string VALUE 'lowestEdited',
          lc_xml_attr_rupbuild           TYPE string VALUE 'rupBuild',
          lc_xml_attr_xwindow            TYPE string VALUE 'xWindow',
          lc_xml_attr_ywindow            TYPE string VALUE 'yWindow',
          lc_xml_attr_windowwidth        TYPE string VALUE 'windowWidth',
          lc_xml_attr_windowheight       TYPE string VALUE 'windowHeight',
          lc_xml_attr_activetab          TYPE string VALUE 'activeTab',
          lc_xml_attr_name               TYPE string VALUE 'name',
          lc_xml_attr_sheetid            TYPE string VALUE 'sheetId',
          lc_xml_attr_state              TYPE string VALUE 'state',
          lc_xml_attr_id                 TYPE string VALUE 'id',
          lc_xml_attr_calcid             TYPE string VALUE 'calcId',
          lc_xml_attr_lockrevision       TYPE string VALUE 'lockRevision',
          lc_xml_attr_lockstructure      TYPE string VALUE 'lockStructure',
          lc_xml_attr_lockwindows        TYPE string VALUE 'lockWindows',
          lc_xml_attr_revisionspassword  TYPE string VALUE 'revisionsPassword',
          lc_xml_attr_workbookpassword   TYPE string VALUE 'workbookPassword',
          lc_xml_attr_hidden             TYPE string VALUE 'hidden',
          lc_xml_attr_localsheetid       TYPE string VALUE 'localSheetId',
          " Node namespace
          lc_r_ns                        TYPE string VALUE 'r',
          lc_xml_node_ns                 TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          " Node id
          lc_xml_node_ridx_id            TYPE string VALUE 'rId#'.

    DATA: lo_document       TYPE REF TO if_ixml_document,
          lo_element_root   TYPE REF TO if_ixml_element,
          lo_element        TYPE REF TO if_ixml_element,
          lo_element_range  TYPE REF TO if_ixml_element,
          lo_sub_element    TYPE REF TO if_ixml_element,
          lo_iterator       TYPE REF TO zcl_excel_collection_iterator,
          lo_iterator_range TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          lo_range          TYPE REF TO zcl_excel_range,
          lo_autofilters    TYPE REF TO zcl_excel_autofilters,
          lo_autofilter     TYPE REF TO zcl_excel_autofilter.

    DATA: lv_xml_node_ridx_id TYPE string,
          lv_value            TYPE string,
          lv_syindex          TYPE string,
          lv_active_sheet     TYPE zexcel_active_worksheet.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_workbook
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).

**********************************************************************
* STEP 4: Create subnode
    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_fileversion
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_appname
                                  value = 'xl' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lastedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lowestedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_rupbuild
                                  value = '4506' ).
    lo_element_root->append_child( new_child = lo_element ).

    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookpr
                                                     parent = lo_document ).
    lo_element_root->append_child( new_child = lo_element ).

    " workbookProtection node
    IF me->excel->zif_excel_book_protection~protected EQ abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookprotection
                                                       parent = lo_document ).
      lv_value = me->excel->zif_excel_book_protection~workbookpassword.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_workbookpassword
                                      value = lv_value ).
      ENDIF.
      lv_value = me->excel->zif_excel_book_protection~revisionspassword.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_revisionspassword
                                      value = lv_value ).
      ENDIF.
      lv_value = me->excel->zif_excel_book_protection~lockrevision.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockrevision
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockstructure.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockstructure
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockwindows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockwindows
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    " bookviews node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_bookviews
                                                     parent = lo_document ).
    " bookview node
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_workbookview
                                                         parent = lo_document ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xwindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_ywindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowwidth
                                      value = '19035' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowheight
                                      value = '8445' ).
    " Set Active Sheet
    lv_active_sheet = excel->get_active_sheet_index( ).
* issue #365 - test if sheet exists - otherwise set active worksheet to 1
    lo_worksheet = excel->get_worksheet_by_index( lv_active_sheet ).
    IF lo_worksheet IS NOT BOUND.
      lv_active_sheet = 1.
      excel->set_active_sheet_index( lv_active_sheet ).
    ENDIF.
    IF lv_active_sheet > 1.
      lv_active_sheet = lv_active_sheet - 1.
      lv_value = lv_active_sheet.
      CONDENSE lv_value.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_activetab
                                        value = lv_value ).
    ENDIF.
    lo_element->append_child( new_child = lo_sub_element )." bookview node
    lo_element_root->append_child( new_child = lo_element )." bookviews node

    " sheets node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheets
                                                     parent = lo_document ).
    lo_iterator = excel->get_worksheets_iterator( ).

    " ranges node
    lo_element_range = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 +
                                                           parent = lo_document ).           " issue 163 +

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " sheet node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_sheet
                                                              parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_syindex = sy-index.                                                                  " question by Stefan SchmÃ¶cker 2012-12-02:  sy-index seems to do the job - but is it proven to work or purely coincedence
      lv_value = lo_worksheet->get_title( ).
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_sheetid
                                        value = lv_syindex ).
      IF lo_worksheet->zif_excel_sheet_properties~hidden EQ zif_excel_sheet_properties=>c_hidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'hidden' ).
      ELSEIF lo_worksheet->zif_excel_sheet_properties~hidden EQ zif_excel_sheet_properties=>c_veryhidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'veryHidden' ).
      ENDIF.
      lo_sub_element->set_attribute_ns( name    = lc_xml_attr_id
                                        prefix  = lc_r_ns
                                        value   = lv_xml_node_ridx_id ).
      lo_element->append_child( new_child = lo_sub_element ). " sheet node

      " issue 163 >>>
      lo_iterator_range = lo_worksheet->get_ranges_iterator( ).

*--------------------------------------------------------------------*
* Defined names sheetlocal:  Ranges, Repeat rows and columns
*--------------------------------------------------------------------*
      WHILE lo_iterator_range->has_next( ) EQ abap_true.
        " range node
        lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                parent = lo_document ).
        lo_range ?= lo_iterator_range->get_next( ).
        lv_value = lo_range->name.

        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                          value = lv_value ).

*      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid           "del #235 Repeat rows/cols - EXCEL starts couting from zero
*                                        value = lv_xml_node_ridx_id ).             "del #235 Repeat rows/cols - and needs absolute referencing to localSheetId
        lv_value   = lv_syindex - 1.                                                  "ins #235 Repeat rows/cols
        CONDENSE lv_value NO-GAPS.                                                    "ins #235 Repeat rows/cols
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                          value = lv_value ).

        lv_value = lo_range->get_value( ).
        lo_sub_element->set_value( value = lv_value ).
        lo_element_range->append_child( new_child = lo_sub_element ). " range node

      ENDWHILE.
      " issue 163 <<<

    ENDWHILE.
    lo_element_root->append_child( new_child = lo_element )." sheets node


*--------------------------------------------------------------------*
* Defined names workbookgolbal:  Ranges
*--------------------------------------------------------------------*
*  " ranges node
*  lo_element = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 -
*                                                   parent = lo_document ).           " issue 163 -
    lo_iterator = excel->get_ranges_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " range node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                              parent = lo_document ).
      lo_range ?= lo_iterator->get_next( ).
      lv_value = lo_range->name.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lv_value = lo_range->get_value( ).
      lo_sub_element->set_value( value = lv_value ).
      lo_element_range->append_child( new_child = lo_sub_element ). " range node

    ENDWHILE.

*--------------------------------------------------------------------*
* Defined names - Autofilters ( also sheetlocal )
*--------------------------------------------------------------------*
    lo_autofilters = excel->get_autofilters_reference( ).
    IF lo_autofilters->is_empty( ) = abap_false.
      lo_iterator = excel->get_worksheets_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.

        lo_worksheet ?= lo_iterator->get_next( ).
        lv_syindex = sy-index - 1 .
        lo_autofilter = lo_autofilters->get( io_worksheet = lo_worksheet ).
        IF lo_autofilter IS BOUND.
          lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                  parent = lo_document ).
          lv_value = lo_autofilters->c_autofilter.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                            value = lv_value ).
          lv_value = lv_syindex.
          CONDENSE lv_value NO-GAPS.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                            value = lv_value ).
          lv_value = '1'. " Always hidden
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_hidden
                                            value = lv_value ).
          lv_value = lo_autofilter->get_filter_reference( ).
          lo_sub_element->set_value( value = lv_value ).
          lo_element_range->append_child( new_child = lo_sub_element ). " range node
        ENDIF.

      ENDWHILE.
    ENDIF.
    lo_element_root->append_child( new_child = lo_element_range ).                      " ranges node


    " calcPr node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_calcpr
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_calcid
                                  value = '125725' ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.


  METHOD create_xml_document.
    DATA lo_encoding TYPE REF TO if_ixml_encoding.
    lo_encoding = me->ixml->create_encoding( byte_order = if_ixml_encoding=>co_platform_endian
                                             character_set = 'utf-8' ).
    ro_document = me->ixml->create_document( ).
    ro_document->set_encoding( lo_encoding ).
    ro_document->set_standalone( abap_true ).
  ENDMETHOD.


  METHOD flag2bool.


    IF ip_flag EQ abap_true.
      ep_boolean = 'true'.
    ELSE.
      ep_boolean = 'false'.
    ENDIF.
  ENDMETHOD.


  METHOD get_shared_string_index.


    DATA ls_shared_string TYPE zexcel_s_shared_string.

    IF it_rtf IS INITIAL.
      READ TABLE shared_strings INTO ls_shared_string WITH TABLE KEY string_value = ip_cell_value.
      ep_index = ls_shared_string-string_no.
    ELSE.
      LOOP AT shared_strings INTO ls_shared_string WHERE string_value = ip_cell_value
                                                     AND rtf_tab = it_rtf.

        ep_index = ls_shared_string-string_no.
        EXIT.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.


  METHOD is_formula_shareable.
    DATA: lv_test_shared TYPE string.

    ep_shareable = abap_false.
    IF ip_formula NA '!'.
      lv_test_shared = zcl_excel_common=>shift_formula(
          iv_reference_formula = ip_formula
          iv_shift_cols        = 1
          iv_shift_rows        = 1 ).
      IF lv_test_shared <> ip_formula.
        ep_shareable = abap_true.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD render_xml_document.
    DATA lo_streamfactory TYPE REF TO if_ixml_stream_factory.
    DATA lo_ostream       TYPE REF TO if_ixml_ostream.
    DATA lo_renderer      TYPE REF TO if_ixml_renderer.
    DATA lv_string        TYPE string.

    " So that the rendering of io_document to a XML text in UTF-8 XSTRING works for all Unicode characters (Chinese,
    " emoticons, etc.) the method CREATE_OSTREAM_CSTRING must be used instead of CREATE_OSTREAM_XSTRING as explained
    " in note 2922674 below (original there: https://launchpad.support.sap.com/#/notes/2922674), and then the STRING
    " variable can be converted into UTF-8.
    "
    " Excerpt from Note 2922674 - Support for Unicode Characters U+10000 to U+10FFFF in the iXML kernel library / ABAP package SIXML.
    "
    "   You are running a unicode system with SAP Netweaver / SAP_BASIS release equal or lower than 7.51.
    "
    "   Some functions in the iXML kernel library / ABAP package SIXML does not fully or incorrectly support unicode
    "   characters of the supplementary planes. This is caused by using UCS-2 in codepage conversion functions.
    "   Therefore, when reading from iXML input steams, the characters from the supplementary planes, that are not
    "   supported by UCS-2, might be replaced by the character #. When writing to iXML output streams, UTF-16 surrogate
    "   pairs, representing characters from the supplementary planes, might be incorrectly encoded in UTF-8.
    "
    "   The characters incorrectly encoded in UTF-8, might be accepted as input for the iXML parser or external parsers,
    "   but might also be rejected.
    "
    "   Support for unicode characters of the supplementary planes was introduced for SAP_BASIS 7.51 or lower with note
    "   2220720, but later withdrawn with note 2346627 for functional issues.
    "
    "   Characters of the supplementary planes are supported with ABAP Platform 1709 / SAP_BASIS 7.52 and higher.
    "
    "   Please note, that the iXML runtime behaves like the ABAP runtime concerning the handling of unicode characters of
    "   the supplementary planes. In iXML and ABAP, these characters have length 2 (as returned by ABAP build-in function
    "   STRLEN), and string processing functions like SUBSTRING might split these characters into 2 invalid characters
    "   with length 1. These invalid characters are commonly referred to as broken surrogate pairs.
    "
    "   A workaround for the incorrect UTF-8 encoding in SAP_BASIS 7.51 or lower is to render the document to an ABAP
    "   variable with type STRING using a output stream created with factory method IF_IXML_STREAM_FACTORY=>CREATE_OSTREAM_CSTRING
    "   and then to convert the STRING variable to UTF-8 using method CL_ABAP_CODEPAGE=>CONVERT_TO.

    " 1) RENDER TO XML STRING
    lo_streamfactory = me->ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_cstring( string = lv_string ).
    lo_renderer = me->ixml->create_renderer( ostream  = lo_ostream document = io_document ).
    lo_renderer->render( ).

    " 2) CONVERT IT TO UTF-8
    "-----------------
    " The beginning of the XML string has these 57 characters:
    "   X<?xml version="1.0" encoding="utf-16" standalone="yes"?>
    "   (where "X" is the special character corresponding to the utf-16 BOM, hexadecimal FFFE or FEFF,
    "   but there's no "X" in non-Unicode SAP systems)
    " The encoding must be removed otherwise Excel would fail to decode correctly the UTF-8 XML.
    " For a better performance, it's assumed that "encoding" is in the first 100 characters.
    IF strlen( lv_string ) < 100.
      REPLACE REGEX 'encoding="[^"]+"' IN lv_string WITH ``.
    ELSE.
      REPLACE REGEX 'encoding="[^"]+"' IN SECTION LENGTH 100 OF lv_string WITH ``.
    ENDIF.
    " Convert XML text to UTF-8 (NB: if 2 first bytes are the UTF-16 BOM, they are converted into 3 bytes of UTF-8 BOM)
    ep_content = cl_abap_codepage=>convert_to( source = lv_string ).
    " Add the UTF-8 Byte Order Mark if missing (NB: that serves as substitute of "encoding")
    IF xstrlen( ep_content ) >= 3 AND ep_content(3) <> cl_abap_char_utilities=>byte_order_mark_utf8.
      CONCATENATE cl_abap_char_utilities=>byte_order_mark_utf8 ep_content INTO ep_content IN BYTE MODE.
    ENDIF.

  ENDMETHOD.


  METHOD set_vml_shape_footer.

    CONSTANTS: lc_shape               TYPE string VALUE '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:{WIDTH}pt;height:{HEIGHT}pt; z-index:1''>',
               lc_shape_image         TYPE string VALUE '<v:imagedata o:relid="{RID}" o:title="Logo Title"/><o:lock v:ext="edit" rotation="t"/></v:shape>',
               lc_shape_header_center TYPE string VALUE 'CH',
               lc_shape_header_left   TYPE string VALUE 'LH',
               lc_shape_header_right  TYPE string VALUE 'RH',
               lc_shape_footer_center TYPE string VALUE 'CF',
               lc_shape_footer_left   TYPE string VALUE 'LF',
               lc_shape_footer_right  TYPE string VALUE 'RF'.

    DATA: lv_content_left         TYPE string,
          lv_content_center       TYPE string,
          lv_content_right        TYPE string,
          lv_content_image_left   TYPE string,
          lv_content_image_center TYPE string,
          lv_content_image_right  TYPE string,
          lv_value                TYPE string,
          ls_drawing_position     TYPE zexcel_drawing_position.

    IF is_footer-left_image IS NOT INITIAL.
      lv_content_left = lc_shape.
      REPLACE '{ID}' IN lv_content_left WITH lc_shape_footer_left.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_left WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_left WITH lv_value.
      lv_content_image_left = lc_shape_image.
      lv_value = is_footer-left_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_left WITH lv_value.
    ENDIF.
    IF is_footer-center_image IS NOT INITIAL.
      lv_content_center = lc_shape.
      REPLACE '{ID}' IN lv_content_center WITH lc_shape_footer_center.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_center WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_center WITH lv_value.
      lv_content_image_center = lc_shape_image.
      lv_value = is_footer-center_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_center WITH lv_value.
    ENDIF.
    IF is_footer-right_image IS NOT INITIAL.
      lv_content_right = lc_shape.
      REPLACE '{ID}' IN lv_content_right WITH lc_shape_footer_right.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_right WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_right WITH lv_value.
      lv_content_image_right = lc_shape_image.
      lv_value = is_footer-right_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_right WITH lv_value.
    ENDIF.

    CONCATENATE lv_content_left
                lv_content_image_left
                lv_content_center
                lv_content_image_center
                lv_content_right
                lv_content_image_right
           INTO ep_content.

  ENDMETHOD.


  METHOD set_vml_shape_header.

*  CONSTANTS: lc_shape TYPE string VALUE '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:198.75pt;height:48.75pt; z-index:1''>',
    CONSTANTS: lc_shape               TYPE string VALUE '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:{WIDTH}pt;height:{HEIGHT}pt; z-index:1''>',
               lc_shape_image         TYPE string VALUE '<v:imagedata o:relid="{RID}" o:title="Logo Title"/><o:lock v:ext="edit" rotation="t"/></v:shape>',
               lc_shape_header_center TYPE string VALUE 'CH',
               lc_shape_header_left   TYPE string VALUE 'LH',
               lc_shape_header_right  TYPE string VALUE 'RH',
               lc_shape_footer_center TYPE string VALUE 'CF',
               lc_shape_footer_left   TYPE string VALUE 'LF',
               lc_shape_footer_right  TYPE string VALUE 'RF'.

    DATA: lv_content_left         TYPE string,
          lv_content_center       TYPE string,
          lv_content_right        TYPE string,
          lv_content_image_left   TYPE string,
          lv_content_image_center TYPE string,
          lv_content_image_right  TYPE string,
          lv_value                TYPE string,
          ls_drawing_position     TYPE zexcel_drawing_position.

    CLEAR ep_content.

    IF is_header-left_image IS NOT INITIAL.
      lv_content_left = lc_shape.
      REPLACE '{ID}' IN lv_content_left WITH lc_shape_header_left.
      ls_drawing_position = is_header-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_left WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_left WITH lv_value.
      lv_content_image_left = lc_shape_image.
      lv_value = is_header-left_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_left WITH lv_value.
    ENDIF.
    IF is_header-center_image IS NOT INITIAL.
      lv_content_center = lc_shape.
      REPLACE '{ID}' IN lv_content_center WITH lc_shape_header_center.
      ls_drawing_position = is_header-center_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_center WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_center WITH lv_value.
      lv_content_image_center = lc_shape_image.
      lv_value = is_header-center_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_center WITH lv_value.
    ENDIF.
    IF is_header-right_image IS NOT INITIAL.
      lv_content_right = lc_shape.
      REPLACE '{ID}' IN lv_content_right WITH lc_shape_header_right.
      ls_drawing_position = is_header-right_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_right WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_right WITH lv_value.
      lv_content_image_right = lc_shape_image.
      lv_value = is_header-right_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_right WITH lv_value.
    ENDIF.

    CONCATENATE lv_content_left
                lv_content_image_left
                lv_content_center
                lv_content_image_center
                lv_content_right
                lv_content_image_right
           INTO ep_content.

  ENDMETHOD.


  METHOD set_vml_string.

    DATA:
      ld_1           TYPE string,
      ld_2           TYPE string,
      ld_3           TYPE string,
      ld_4           TYPE string,
      ld_5           TYPE string,
      ld_7           TYPE string,

      lv_relation_id TYPE i,
      lo_iterator    TYPE REF TO zcl_excel_collection_iterator,
      lo_worksheet   TYPE REF TO zcl_excel_worksheet,
      ls_odd_header  TYPE zexcel_s_worksheet_head_foot,
      ls_odd_footer  TYPE zexcel_s_worksheet_head_foot,
      ls_even_header TYPE zexcel_s_worksheet_head_foot,
      ls_even_footer TYPE zexcel_s_worksheet_head_foot.


* INIT_RESULT
    CLEAR ep_content.


* BODY
    ld_1 = '<xml xmlns:v="urn:schemas-microsoft-com:vml"  xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'.
    ld_2 = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>'.
    ld_3 = '<v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'.
    ld_4 = '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'.


    CONCATENATE ld_1
                ld_2
                ld_3
                ld_4
         INTO ep_content.

    lv_relation_id = 0.
    lo_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lo_worksheet->sheet_setup->get_header_footer( IMPORTING ep_odd_header = ls_odd_header
                                                              ep_odd_footer = ls_odd_footer
                                                              ep_even_header = ls_even_header
                                                              ep_even_footer = ls_even_footer ).

      ld_5 = me->set_vml_shape_header( ls_odd_header ).
      CONCATENATE ep_content
                  ld_5
             INTO ep_content.
      ld_5 = me->set_vml_shape_header( ls_even_header ).
      CONCATENATE ep_content
                  ld_5
             INTO ep_content.
      ld_5 = me->set_vml_shape_footer( ls_odd_footer ).
      CONCATENATE ep_content
                  ld_5
             INTO ep_content.
      ld_5 = me->set_vml_shape_footer( ls_even_footer ).
      CONCATENATE ep_content
                  ld_5
             INTO ep_content.
    ENDWHILE.

    ld_7 = '</xml>'.

    CONCATENATE ep_content
                ld_7
           INTO ep_content.

  ENDMETHOD.


  METHOD zif_excel_writer~write_file.
    me->excel = io_excel.

    ep_file = me->create( ).
  ENDMETHOD.
ENDCLASS.
