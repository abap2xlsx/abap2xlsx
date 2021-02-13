class ZCL_EXCEL_READER_2007 definition
  public
  create public .

public section.
*"* public components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!
  type-pools IXML .

  interfaces ZIF_EXCEL_READER .

  class-methods FILL_STRUCT_FROM_ATTRIBUTES
    importing
      !IP_ELEMENT type ref to IF_IXML_ELEMENT
    changing
      !CP_STRUCTURE type ANY .
protected section.

  types:
*"* protected components of class ZCL_EXCEL_READER_2007
*"* do not include other source files here!!!
    BEGIN OF t_relationship,
      id         TYPE string,
      type       TYPE string,
      target     TYPE string,
      targetmode TYPE string,
      worksheet  TYPE REF TO zcl_excel_worksheet,
      sheetid    TYPE string,     "ins #235 - repeat rows/cols - needed to identify correct sheet
    END OF t_relationship .
  types:
    BEGIN OF t_fileversion,
      appname      TYPE string,
      lastedited   TYPE string,
      lowestedited TYPE string,
      rupbuild     TYPE string,
      codename     TYPE string,
    END OF t_fileversion .
  types:
    BEGIN OF t_sheet,
      name    TYPE string,
      sheetid TYPE string,
      id      TYPE string,
      state   TYPE string,
    END OF t_sheet .
  types:
    BEGIN OF t_workbookpr,
      codename            TYPE string,
      defaultthemeversion TYPE string,
    END OF t_workbookpr .
  types:
    BEGIN OF t_sheetpr,
      codename TYPE string,
    END OF t_sheetpr .
  types:
    BEGIN OF t_range,
      name         TYPE string,
      hidden       TYPE string,       "inserted with issue #235 because Autofilters didn't passthrough
      localsheetid TYPE string,       " issue #163
    END OF t_range .
  types:
    t_fills   TYPE STANDARD TABLE OF REF TO zcl_excel_style_fill WITH NON-UNIQUE DEFAULT KEY .
  types:
    t_borders TYPE STANDARD TABLE OF REF TO zcl_excel_style_borders WITH NON-UNIQUE DEFAULT KEY .
  types:
    t_fonts   TYPE STANDARD TABLE OF REF TO zcl_excel_style_font WITH NON-UNIQUE DEFAULT KEY .
  types:
    t_style_refs TYPE STANDARD TABLE OF REF TO zcl_excel_style WITH NON-UNIQUE DEFAULT KEY .
  types:
    BEGIN OF t_color,
      indexed TYPE string,
      rgb     TYPE string,
      theme   TYPE string,
      tint    TYPE string,
    END OF t_color .
  types:
    BEGIN OF t_rel_drawing,
      id          TYPE string,
      content     TYPE xstring,
      file_ext    TYPE string,
      content_xml TYPE REF TO if_ixml_document,
    END OF t_rel_drawing .
  types:
    t_rel_drawings TYPE STANDARD TABLE OF t_rel_drawing WITH NON-UNIQUE DEFAULT KEY .
  types:
    BEGIN OF gts_external_hyperlink,
      id     TYPE string,
      target TYPE string,
    END OF gts_external_hyperlink .
  types:
    gtt_external_hyperlinks TYPE HASHED TABLE OF gts_external_hyperlink WITH UNIQUE KEY id .
  types:
    BEGIN OF ty_ref_formulae,
      sheet   TYPE REF TO zcl_excel_worksheet,
      row     TYPE i,
      column  TYPE i,
      si      TYPE i,
      ref     TYPE string,
      formula TYPE string,
    END   OF ty_ref_formulae .
  types:
    tyt_ref_formulae TYPE HASHED TABLE OF ty_ref_formulae WITH UNIQUE KEY sheet row column .

  data SHARED_STRINGS type STRINGTAB .
  data STYLES type T_STYLE_REFS .
  data MT_REF_FORMULAE type TYT_REF_FORMULAE .
  data MT_DXF_STYLES type ZEXCEL_T_STYLES_COND_MAPPING .

  methods FILL_ROW_OUTLINES
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods GET_FROM_ZIP_ARCHIVE
    importing
      !I_FILENAME type STRING
    returning
      value(R_CONTENT) type XSTRING
    raising
      ZCX_EXCEL .
  methods GET_IXML_FROM_ZIP_ARCHIVE
    importing
      !I_FILENAME type STRING
      !IS_NORMALIZING type abap_bool default 'X'
    returning
      value(R_IXML) type ref to IF_IXML_DOCUMENT
    raising
      ZCX_EXCEL .
  methods LOAD_DRAWING_ANCHOR
    importing
      !IO_ANCHOR_ELEMENT type ref to IF_IXML_ELEMENT
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IT_RELATED_DRAWINGS type T_REL_DRAWINGS .
  methods LOAD_SHARED_STRINGS
    importing
      !IP_PATH type STRING
    raising
      ZCX_EXCEL .
  methods LOAD_STYLES
    importing
      !IP_PATH type STRING
      !IP_EXCEL type ref to ZCL_EXCEL
    raising
      ZCX_EXCEL .
  methods LOAD_DXF_STYLES
    importing
      !IV_PATH type STRING
      !IO_EXCEL type ref to ZCL_EXCEL
    raising
      ZCX_EXCEL .
  methods LOAD_STYLE_BORDERS
    importing
      !IP_XML type ref to IF_IXML_DOCUMENT
    returning
      value(EP_BORDERS) type T_BORDERS .
  methods LOAD_STYLE_FILLS
    importing
      !IP_XML type ref to IF_IXML_DOCUMENT
    returning
      value(EP_FILLS) type T_FILLS .
  methods LOAD_STYLE_FONTS
    importing
      !IP_XML type ref to IF_IXML_DOCUMENT
    returning
      value(EP_FONTS) type T_FONTS .
  methods LOAD_STYLE_NUM_FORMATS
    importing
      !IP_XML type ref to IF_IXML_DOCUMENT
    returning
      value(EP_NUM_FORMATS) type ZCL_EXCEL_STYLE_NUMBER_FORMAT=>T_NUM_FORMATS .
  methods LOAD_WORKBOOK
    importing
      !IV_WORKBOOK_FULL_FILENAME type STRING
      !IO_EXCEL type ref to ZCL_EXCEL
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET
    importing
      !IP_PATH type STRING
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET_COND_FORMAT
    importing
      !IO_IXML_WORKSHEET type ref to IF_IXML_DOCUMENT
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET_COND_FORMAT_AA
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND.
  methods LOAD_WORKSHEET_COND_FORMAT_CI
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_COND_FORMAT_CS
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_COND_FORMAT_EX
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_COND_FORMAT_IS
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_COND_FORMAT_DB
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_COND_FORMAT_T10
    importing
      !IO_IXML_RULE type ref to IF_IXML_ELEMENT
      !IO_STYLE_COND type ref to ZCL_EXCEL_STYLE_COND .
  methods LOAD_WORKSHEET_DRAWING
    importing
      !IP_PATH type STRING
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET_HYPERLINKS
    importing
      !IO_IXML_WORKSHEET type ref to IF_IXML_DOCUMENT
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
      !IT_EXTERNAL_HYPERLINKS type GTT_EXTERNAL_HYPERLINKS
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET_PAGEBREAKS
    importing
      !IO_IXML_WORKSHEET type ref to IF_IXML_DOCUMENT
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods LOAD_WORKSHEET_PAGEMARGINS
    importing
      !IO_IXML_WORKSHEET type ref to IF_IXML_DOCUMENT
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  class-methods RESOLVE_PATH
    importing
      !IP_PATH type STRING
    returning
      value(RP_RESULT) type STRING .
  methods RESOLVE_REFERENCED_FORMULAE .
  methods GET_DXF_STYLE_GUID
    importing
      !IO_IXML_DXF type ref to IF_IXML_ELEMENT
      !IO_EXCEL type ref to ZCL_EXCEL
    returning
      value(RV_STYLE_GUID) type ZEXCEL_CELL_STYLE .
  methods LOAD_THEME
    importing
      value(IV_PATH) type STRING
      !IP_EXCEL type ref to ZCL_EXCEL .
private section.

  data ZIP type ref to LCL_ZIP_ARCHIVE .

  methods CREATE_ZIP_ARCHIVE
    importing
      !I_XLSX_BINARY type XSTRING
      !I_USE_ALTERNATE_ZIP type SEOCLSNAME optional
    returning
      value(E_ZIP) type ref to LCL_ZIP_ARCHIVE
    raising
      ZCX_EXCEL .
  methods READ_FROM_APPLSERVER
    importing
      !I_FILENAME type CSEQUENCE
    returning
      value(R_EXCEL_DATA) type XSTRING
    raising
      ZCX_EXCEL.
  methods READ_FROM_LOCAL_FILE
    importing
      !I_FILENAME type CSEQUENCE
    returning
      value(R_EXCEL_DATA) type XSTRING
    raising
      ZCX_EXCEL .
ENDCLASS.



CLASS ZCL_EXCEL_READER_2007 IMPLEMENTATION.


METHOD create_zip_archive.
  CASE i_use_alternate_zip.
    WHEN space.
      e_zip = lcl_abap_zip_archive=>create( i_xlsx_binary ).
    WHEN OTHERS.
      e_zip = lcl_alternate_zip_archive=>create( i_data                = i_xlsx_binary
                                                 i_alternate_zip_class = i_use_alternate_zip ).
  ENDCASE.
ENDMETHOD.


METHOD fill_row_outlines.

  TYPES: BEGIN OF lts_row_data,
           row           TYPE i,
           outline_level TYPE i,
         END OF lts_row_data,
         ltt_row_data TYPE SORTED TABLE OF lts_row_data WITH UNIQUE KEY row.

  DATA: lt_row_data             TYPE ltt_row_data,
        ls_row_data             LIKE LINE OF lt_row_data,
        lt_collapse_rows        TYPE HASHED TABLE OF i WITH UNIQUE KEY table_line,

        lv_collapsed            TYPE abap_bool,

        lv_outline_level        TYPE i,
        lv_next_consecutive_row TYPE i,
        lt_outline_rows         TYPE zcl_excel_worksheet=>mty_ts_outlines_row,
        ls_outline_row          LIKE LINE OF lt_outline_rows,
        lo_row                  TYPE REF TO zcl_excel_row,
        lo_row_iterator         TYPE REF TO cl_object_collection_iterator,
        lv_row_offset           TYPE i,
        lv_row_collapse_flag    TYPE i.


  FIELD-SYMBOLS: <ls_row_data>      LIKE LINE OF lt_row_data.

* First collect information about outlines ( outline leven and collapsed state )
  lo_row_iterator = io_worksheet->get_rows_iterator( ).
  WHILE lo_row_iterator->has_next( ) = abap_true.
    lo_row ?= lo_row_iterator->get_next( ).
    ls_row_data-row           = lo_row->get_row_index( ).
    ls_row_data-outline_level = lo_row->get_outline_level( ).
    IF ls_row_data-outline_level IS NOT INITIAL.
      INSERT ls_row_data INTO TABLE lt_row_data.
    ENDIF.

    lv_collapsed = lo_row->get_collapsed( ).
    IF lv_collapsed = abap_true.
      INSERT lo_row->get_row_index( ) INTO TABLE lt_collapse_rows.
    ENDIF.
  ENDWHILE.

* Now parse this information - we need consecutive rows - any gap will create a new outline
  DO 7 TIMES.  " max number of outlines allowed
    lv_outline_level = sy-index.
    CLEAR lv_next_consecutive_row.
    CLEAR ls_outline_row.
    LOOP AT lt_row_data ASSIGNING <ls_row_data> WHERE outline_level >= lv_outline_level.

      IF lv_next_consecutive_row    <> <ls_row_data>-row   " A gap --> close all open outlines
        AND lv_next_consecutive_row IS NOT INITIAL.        " First time in loop.
        INSERT ls_outline_row INTO TABLE lt_outline_rows.
        CLEAR: ls_outline_row.
      ENDIF.

      IF ls_outline_row-row_from IS INITIAL.
        ls_outline_row-row_from = <ls_row_data>-row.
      ENDIF.
      ls_outline_row-row_to = <ls_row_data>-row.

      lv_next_consecutive_row = <ls_row_data>-row + 1.

    ENDLOOP.
    IF ls_outline_row-row_from IS NOT INITIAL.
      INSERT ls_outline_row INTO TABLE lt_outline_rows.
    ENDIF.
  ENDDO.

* lt_outline_rows holds all outline information
* we now need to determine whether the outline is collapsed or not
  LOOP AT lt_outline_rows INTO ls_outline_row.

    IF io_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_off.
      lv_row_collapse_flag = ls_outline_row-row_from - 1.
    ELSE.
      lv_row_collapse_flag = ls_outline_row-row_to + 1.
    ENDIF.
    READ TABLE lt_collapse_rows TRANSPORTING NO FIELDS WITH TABLE KEY table_line = lv_row_collapse_flag.
    IF sy-subrc = 0.
      ls_outline_row-collapsed = abap_true.
    ENDIF.
    io_worksheet->set_row_outline( iv_row_from  = ls_outline_row-row_from
                                   iv_row_to    = ls_outline_row-row_to
                                   iv_collapsed = ls_outline_row-collapsed ).

  ENDLOOP.

* Finally purge outline information ( collapsed state, outline leve)  from row_dimensions, since we want to keep these in the outline-table
  lo_row_iterator = io_worksheet->get_rows_iterator( ).
  WHILE lo_row_iterator->has_next( ) = abap_true.
    lo_row ?= lo_row_iterator->get_next( ).

    lo_row->set_outline_level( 0 ).
    lo_row->set_collapsed( abap_false ).

  ENDWHILE.

ENDMETHOD.


method fill_struct_from_attributes.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

  data:       lv_name                         type string,
              lo_attributes                   type ref to if_ixml_named_node_map,
              lo_attribute                    type ref to if_ixml_attribute,
              lo_iterator                     type ref to if_ixml_node_iterator.

  field-symbols: <component>                  type any.

*--------------------------------------------------------------------*
* The values of named attributes of a tag are being read and moved into corresponding
* fields of given structure
* Behaves like move-corresonding tag to structure

* Example:
*     <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*   Here the attributes are Target, Type and Id.  Thus if the passed
*   structure has fieldnames Id and Target these would be filled with
*   "rId3" and "docProps/app.xml" respectively
*--------------------------------------------------------------------*
  clear cp_structure.

  lo_attributes  = ip_element->get_attributes( ).
  lo_iterator    = lo_attributes->create_iterator( ).
  lo_attribute  ?= lo_iterator->get_next( ).
  while lo_attribute is bound.

    lv_name = lo_attribute->get_name( ).
    translate lv_name to upper case.
    assign component lv_name of structure cp_structure to <component>.
    if sy-subrc = 0.
      <component> = lo_attribute->get_value( ).
    endif.
    lo_attribute ?= lo_iterator->get_next( ).

  endwhile.


endmethod.


METHOD get_dxf_style_guid.
  DATA: lo_ixml_dxf_children          TYPE REF TO if_ixml_node_list,
        lo_ixml_iterator_dxf_children TYPE REF TO if_ixml_node_iterator,
        lo_ixml_dxf_child             TYPE REF TO if_ixml_element,

        lv_dxf_child_type             TYPE string,

        lo_ixml_element               TYPE REF TO if_ixml_element,
        lo_ixml_element2              TYPE REF TO if_ixml_element,
        lv_val                        TYPE string.

  DATA: ls_cstyle  TYPE zexcel_s_cstyle_complete,
        ls_cstylex TYPE zexcel_s_cstylex_complete.



  lo_ixml_dxf_children = io_ixml_dxf->get_children( ).
  lo_ixml_iterator_dxf_children = lo_ixml_dxf_children->create_iterator( ).
  lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).
  WHILE lo_ixml_dxf_child IS BOUND.

    lv_dxf_child_type = lo_ixml_dxf_child->get_name( ).
    CASE lv_dxf_child_type.

      WHEN 'font'.
*--------------------------------------------------------------------*
* italic
*--------------------------------------------------------------------*
        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'i' ).
        IF lo_ixml_element IS BOUND.
          CLEAR lv_val.
          lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
          IF lv_val <> '0'.
            ls_cstyle-font-italic  = 'X'.
            ls_cstylex-font-italic = 'X'.
          ENDIF.

        ENDIF.
*--------------------------------------------------------------------*
* bold
*--------------------------------------------------------------------*
        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'b' ).
        IF lo_ixml_element IS BOUND.
          CLEAR lv_val.
          lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
          IF lv_val <> '0'.
            ls_cstyle-font-bold  = 'X'.
            ls_cstylex-font-bold = 'X'.
          ENDIF.

        ENDIF.
*--------------------------------------------------------------------*
* strikethrough
*--------------------------------------------------------------------*
        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'strike' ).
        IF lo_ixml_element IS BOUND.
          CLEAR lv_val.
          lv_val  = lo_ixml_element->get_attribute_ns( 'val' ).
          IF lv_val <> '0'.
            ls_cstyle-font-strikethrough  = 'X'.
            ls_cstylex-font-strikethrough = 'X'.
          ENDIF.

        ENDIF.
*--------------------------------------------------------------------*
* color
*--------------------------------------------------------------------*
        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'color' ).
        IF lo_ixml_element IS BOUND.
          CLEAR lv_val.
          lv_val  = lo_ixml_element->get_attribute_ns( 'rgb' ).
          ls_cstyle-font-color-rgb  = lv_val.
          ls_cstylex-font-color-rgb = 'X'.
        ENDIF.

      WHEN 'fill'.
        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'patternFill' ).
        IF lo_ixml_element IS BOUND.
          lo_ixml_element2 = lo_ixml_element->find_from_name( 'bgColor' ).
          IF lo_ixml_element2 IS BOUND.
            CLEAR lv_val.
            lv_val  = lo_ixml_element2->get_attribute_ns( 'rgb' ).
            IF lv_val IS NOT INITIAL.
              ls_cstyle-fill-filltype       = zcl_excel_style_fill=>c_fill_solid.
              ls_cstyle-fill-bgcolor-rgb    = lv_val.
              ls_cstylex-fill-filltype      = 'X'.
              ls_cstylex-fill-bgcolor-rgb   = 'X'.
            ENDIF.
            CLEAR lv_val.
            lv_val  = lo_ixml_element2->get_attribute_ns( 'theme' ).
            IF lv_val IS NOT INITIAL.
              ls_cstyle-fill-filltype         = zcl_excel_style_fill=>c_fill_solid.
              ls_cstyle-fill-bgcolor-theme    = lv_val.
              ls_cstylex-fill-filltype        = 'X'.
              ls_cstylex-fill-bgcolor-theme   = 'X'.
            ENDIF.
            CLEAR lv_val.
          ENDIF.
        ENDIF.

* 2do - borders into dxf-styles.  Here and in writerclass
*      WHEN 'border'.
*        lo_ixml_element = lo_ixml_dxf_child->find_from_name( 'left' ).
*        IF lo_ixml_element IS BOUND.
*          CLEAR lv_val.
*          lv_val  = lo_ixml_element2->get_attribute_ns( 'style' ).
*          IF lv_val IS NOT INITIAL.
*            ls_cstyle-borders-left-border_style  = lv_val.
*            ls_cstylex-borders-left-border_style = 'X'.
*          ENDIF.
*        ENDIF.

    ENDCASE.

    lo_ixml_dxf_child ?= lo_ixml_iterator_dxf_children->get_next( ).

  ENDWHILE.

  rv_style_guid = io_excel->get_static_cellstyle_guid( ip_cstyle_complete  = ls_cstyle
                                                       ip_cstylex_complete = ls_cstylex  ).


ENDMETHOD.


METHOD get_from_zip_archive.

  ASSERT zip IS BOUND. " zip object has to exist at this point

  r_content = zip->read(  i_filename ).

ENDMETHOD.


METHOD get_ixml_from_zip_archive.

  DATA:       lv_content                     TYPE xstring,
              lo_ixml                        TYPE REF TO if_ixml,
              lo_streamfactory               TYPE REF TO if_ixml_stream_factory,
              lo_istream                     TYPE REF TO if_ixml_istream,
              lo_parser                      TYPE REF TO if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*
  lv_content        = me->get_from_zip_archive( i_filename ).
  lo_ixml           = cl_ixml=>create( ).
  lo_streamfactory  = lo_ixml->create_stream_factory( ).
  lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
  r_ixml            = lo_ixml->create_document( ).
  lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                              istream        = lo_istream
                                              document       = r_ixml ).
  lo_parser->set_normalizing( is_normalizing ).
  lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
  lo_parser->parse( ).

ENDMETHOD.


METHOD load_drawing_anchor.

  TYPES: BEGIN OF t_c_nv_pr,
           name TYPE string,
           id   TYPE string,
         END OF t_c_nv_pr.

  TYPES: BEGIN OF t_blip,
           cstate TYPE string,
           embed  TYPE string,
         END OF t_blip.

  TYPES: BEGIN OF t_chart,
           id TYPE string,
         END OF t_chart.

  TYPES: BEGIN OF t_ext,
           cx TYPE string,
           cy TYPE string,
         END OF t_ext.

  CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
             lc_xml_attr_true_int TYPE string VALUE '1'.
  CONSTANTS: lc_rel_chart TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
             lc_rel_image TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.

  DATA: lo_drawing     TYPE REF TO zcl_excel_drawing,
        node           TYPE REF TO if_ixml_element,
        node2          TYPE REF TO if_ixml_element,
        node3          TYPE REF TO if_ixml_element,
        node4          TYPE REF TO if_ixml_element,

        ls_upper       TYPE zexcel_drawing_location,
        ls_lower       TYPE zexcel_drawing_location,
        ls_size        TYPE zexcel_drawing_size,
        ext            TYPE t_ext,
        lv_content     TYPE xstring,
        lv_relation_id TYPE string,
        lv_title       TYPE string,

        cnvpr          TYPE t_c_nv_pr,
        blip           TYPE t_blip,
        chart          TYPE t_chart,
        drawing_type   TYPE zexcel_drawing_type,

        rel_drawing    TYPE t_rel_drawing.

  node ?= io_anchor_element->find_from_name( name = 'from' namespace = 'xdr' ).
  CHECK node IS NOT INITIAL.
  node2 ?= node->find_from_name( name = 'col' namespace = 'xdr' ).
  ls_upper-col = node2->get_value( ).
  node2 ?= node->find_from_name( name = 'row' namespace = 'xdr' ).
  ls_upper-row = node2->get_value( ).
  node2 ?= node->find_from_name( name = 'colOff' namespace = 'xdr' ).
  ls_upper-col_offset = node2->get_value( ).
  node2 ?= node->find_from_name( name = 'rowOff' namespace = 'xdr' ).
  ls_upper-row_offset = node2->get_value( ).

  node ?= io_anchor_element->find_from_name( name = 'ext' namespace = 'xdr' ).
  IF node IS INITIAL.
    CLEAR ls_size.
  ELSE.
    me->fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = ext ).
    ls_size-width = ext-cx.
    ls_size-height = ext-cy.
    TRY.
        ls_size-width  = zcl_excel_drawing=>emu2pixel( ls_size-width ).
      CATCH cx_root.
    ENDTRY.
    TRY.
        ls_size-height = zcl_excel_drawing=>emu2pixel( ls_size-height ).
      CATCH cx_root.
    ENDTRY.
  ENDIF.

  node ?= io_anchor_element->find_from_name( name = 'to' namespace = 'xdr' ).
  IF node IS INITIAL.
    CLEAR ls_lower.
  ELSE.
    node2 ?= node->find_from_name( name = 'col' namespace = 'xdr' ).
    ls_lower-col = node2->get_value( ).
    node2 ?= node->find_from_name( name = 'row' namespace = 'xdr' ).
    ls_lower-row = node2->get_value( ).
    node2 ?= node->find_from_name( name = 'colOff' namespace = 'xdr' ).
    ls_lower-col_offset = node2->get_value( ).
    node2 ?= node->find_from_name( name = 'rowOff' namespace = 'xdr' ).
    ls_lower-row_offset = node2->get_value( ).
  ENDIF.

  node ?= io_anchor_element->find_from_name( name = 'pic' namespace = 'xdr' ).
  IF node IS NOT INITIAL.
    node2 ?= node->find_from_name( name = 'nvPicPr' namespace = 'xdr' ).
    CHECK node2 IS NOT INITIAL.
    node3 ?= node2->find_from_name( name = 'cNvPr' namespace = 'xdr' ).
    CHECK node3 IS NOT INITIAL.
    me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = cnvpr ).
    lv_title = cnvpr-name.

    node2 ?= node->find_from_name( name = 'blipFill' namespace = 'xdr' ).
    CHECK node2 IS NOT INITIAL.
    node3 ?= node2->find_from_name( name = 'blip' namespace = 'a' ).
    CHECK node3 IS NOT INITIAL.
    me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = blip ).
    lv_relation_id = blip-embed.

    drawing_type = zcl_excel_drawing=>type_image.
  ENDIF.

  node ?= io_anchor_element->find_from_name( name = 'graphicFrame' namespace = 'xdr' ).
  IF node IS NOT INITIAL.
    node2 ?= node->find_from_name( name = 'nvGraphicFramePr' namespace = 'xdr' ).
    CHECK node2 IS NOT INITIAL.
    node3 ?= node2->find_from_name( name = 'cNvPr' namespace = 'xdr' ).
    CHECK node3 IS NOT INITIAL.
    me->fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = cnvpr ).
    lv_title = cnvpr-name.

    node2 ?= node->find_from_name( name = 'graphic' namespace = 'a' ).
    CHECK node2 IS NOT INITIAL.
    node3 ?= node2->find_from_name( name = 'graphicData' namespace = 'a' ).
    CHECK node3 IS NOT INITIAL.
    node4 ?= node2->find_from_name( name = 'chart' namespace = 'c' ).
    CHECK node4 IS NOT INITIAL.
    me->fill_struct_from_attributes( EXPORTING ip_element = node4 CHANGING cp_structure = chart ).
    lv_relation_id = chart-id.

    drawing_type = zcl_excel_drawing=>type_chart.
  ENDIF.

  lo_drawing = io_worksheet->excel->add_new_drawing(
                    ip_type  = drawing_type
                    ip_title = lv_title ).
  io_worksheet->add_drawing( lo_drawing ).

  lo_drawing->set_position2(
    EXPORTING
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  READ TABLE it_related_drawings INTO rel_drawing
        WITH KEY id = lv_relation_id.

  lo_drawing->set_media(
    EXPORTING
      ip_media = rel_drawing-content
      ip_media_type = rel_drawing-file_ext
      ip_width = ls_size-width
      ip_height = ls_size-height ).

  IF drawing_type = zcl_excel_drawing=>type_chart.
*  Begin fix for Issue #551
    DATA: lo_tmp_node_2                TYPE REF TO if_ixml_element.
    lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name( name = 'pieChart' namespace = 'c' ).
    if lo_tmp_node_2 is not initial.
        lo_drawing->graph_type = zcl_excel_drawing=>c_graph_pie.
    else.
        lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name( name = 'barChart' namespace = 'c' ).
          if lo_tmp_node_2 is not initial.
                lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.
          else.
              lo_tmp_node_2 ?= rel_drawing-content_xml->find_from_name( name = 'lineChart' namespace = 'c' ).
              if lo_tmp_node_2 is not initial.
                   lo_drawing->graph_type = zcl_excel_drawing=>c_graph_line.
              endif.
          endif.
    endif.
* End fix for issue #551
    "-------------Added by Alessandro Iannacci - Should load chart attributes
    lo_drawing->load_chart_attributes( rel_drawing-content_xml ).
  ENDIF.

ENDMETHOD.


METHOD load_dxf_styles.

  DATA:        lo_styles_xml   TYPE REF TO if_ixml_document,
               lo_node_dxfs    TYPE REF TO if_ixml_element,

               lo_nodes_dxf    TYPE REF TO if_ixml_node_collection,
               lo_iterator_dxf TYPE REF TO if_ixml_node_iterator,
               lo_node_dxf     TYPE REF TO if_ixml_element,

               lv_dxf_count    TYPE i.

  FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF mt_dxf_styles.

*--------------------------------------------------------------------*
* Look for dxfs-node
*--------------------------------------------------------------------*
  lo_styles_xml = me->get_ixml_from_zip_archive( iv_path ).
  lo_node_dxfs  = lo_styles_xml->find_from_name( 'dxfs' ).
  CHECK lo_node_dxfs IS BOUND.


*--------------------------------------------------------------------*
* loop through all dxf-nodes and create style for each
*--------------------------------------------------------------------*
  lo_nodes_dxf ?= lo_node_dxfs->get_elements_by_tag_name( 'dxf' ).
  lo_iterator_dxf = lo_nodes_dxf->create_iterator( ).
  lo_node_dxf ?= lo_iterator_dxf->get_next( ).
  WHILE lo_node_dxf IS BOUND.

    APPEND INITIAL LINE TO mt_dxf_styles ASSIGNING <ls_dxf_style>.
    <ls_dxf_style>-dxf = lv_dxf_count. " We start counting at 0
    ADD 1 TO lv_dxf_count.             " prepare next entry

    <ls_dxf_style>-guid = get_dxf_style_guid( io_ixml_dxf = lo_node_dxf
                                              io_excel    = io_excel ).
    lo_node_dxf ?= lo_iterator_dxf->get_next( ).

  ENDWHILE.


ENDMETHOD.


method LOAD_SHARED_STRINGS.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support partial formatting of strings in cells
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: renaming variables to naming conventions
*          renaming variables to indicate what they are used for
*          aligning code
*          adding comments to explain what we are trying to achieve
*          rewriting code for better readibility
*--------------------------------------------------------------------*



  DATA:
              lo_shared_strings_xml           TYPE REF TO if_ixml_document,
              lo_node_si                      TYPE REF TO if_ixml_element,
              lo_node_si_child                TYPE REF TO if_ixml_element,
              lo_node_r_child_t               TYPE REF TO if_ixml_element,
              lv_tag_name                     TYPE string,
              lv_node_value                   TYPE string.

  FIELD-SYMBOLS: <lv_shared_string>           LIKE LINE OF me->shared_strings.

*--------------------------------------------------------------------*

* §1  Parse shared strings file and get into internal table
*   So far I have encountered 2 ways how a string can be represented in the shared strings file
*   §1.1 - "simple" strings
*   §1.2 - rich text formatted strings

*     Following is an example how this file could be set up; 2 strings in simple formatting, 3rd string rich textformatted


*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <sst uniqueCount="6" count="6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <si>
*                <t>This is a teststring 1</t>
*            </si>
*            <si>
*                <t>This is a teststring 2</t>
*            </si>
*            <si>
*                <r>
*                  <t>T</t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color rgb="FFFF0000"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t xml:space="preserve">his is a </t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color theme="1"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t>teststring 3</t>
*                </r>
*            </si>
*        </sst>
*--------------------------------------------------------------------*

  lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
                                                         is_normalizing = space ).  " NO!!! normalizing - otherwise leading blanks will be omitted and that is not really desired for the stringtable
  lo_node_si ?= lo_shared_strings_xml->find_from_name( 'si' ).
  WHILE lo_node_si IS BOUND.

    APPEND INITIAL LINE TO me->shared_strings ASSIGNING <lv_shared_string>.            " Each <si>-entry in the xml-file must lead to an entry in our stringtable
    lo_node_si_child ?= lo_node_si->get_first_child( ).
    IF lo_node_si_child IS BOUND.
      lv_tag_name = lo_node_si_child->get_name( ).
      IF lv_tag_name = 't'.
*--------------------------------------------------------------------*
*   §1.1 - "simple" strings
*                Example:  see above
*--------------------------------------------------------------------*
        <lv_shared_string> = lo_node_si_child->get_value( ).
      ELSE.
*--------------------------------------------------------------------*
*   §1.2 - rich text formatted strings
*       it is sufficient to strip the <t>...</t> tag from each <r>-tag and concatenate these
*       as long as rich text formatting is not supported (2do§1) ignore all info about formatting
*                Example:  see above
*--------------------------------------------------------------------*
        WHILE lo_node_si_child IS BOUND.                                             " actually these children of <si> are <r>-tags

          lo_node_r_child_t ?= lo_node_si_child->find_from_name( 't' ).              " extract the <t>...</t> part of each <r>-tag
          IF lo_node_r_child_t IS BOUND.
            lv_node_value = lo_node_r_child_t->get_value( ).
            CONCATENATE <lv_shared_string> lv_node_value INTO <lv_shared_string> RESPECTING BLANKS.
          ENDIF.

          lo_node_si_child ?= lo_node_si_child->get_next( ).

        ENDWHILE.
      ENDIF.
    ENDIF.

    lo_node_si ?= lo_node_si->get_next( ).
  ENDWHILE.

  endmethod.


METHOD load_styles.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
  TYPES:  BEGIN OF lty_xf,
            applyalignment    TYPE string,
            applyborder       TYPE string,
            applyfill         TYPE string,
            applyfont         TYPE string,
            applynumberformat TYPE string,
            applyprotection   TYPE string,
            borderid          TYPE string,
            fillid            TYPE string,
            fontid            TYPE string,
            numfmtid          TYPE string,
            pivotbutton       TYPE string,
            quoteprefix       TYPE string,
            xfid              TYPE string,
          END OF lty_xf.

  TYPES:  BEGIN OF lty_alignment,
            horizontal      TYPE string,
            indent          TYPE string,
            justifylastline TYPE string,
            readingorder    TYPE string,
            relativeindent  TYPE string,
            shrinktofit     TYPE string,
            textrotation    TYPE string,
            vertical        TYPE string,
            wraptext        TYPE string,
          END OF lty_alignment.

  TYPES:  BEGIN OF lty_protection,
            hidden TYPE string,
            locked TYPE string,
          END OF lty_protection.

  DATA:       lo_styles_xml                 TYPE REF TO if_ixml_document,
              lo_style                      TYPE REF TO zcl_excel_style,

              lt_num_formats                TYPE zcl_excel_style_number_format=>t_num_formats,
              lt_fills                      TYPE t_fills,
              lt_borders                    TYPE t_borders,
              lt_fonts                      TYPE t_fonts,

              ls_num_format                 TYPE zcl_excel_style_number_format=>t_num_format ,
              ls_fill                       TYPE REF TO zcl_excel_style_fill,
              ls_cell_border                TYPE REF TO zcl_excel_style_borders,
              ls_font                       TYPE REF TO zcl_excel_style_font,

              lo_node_cellxfs               TYPE REF TO if_ixml_element,
              lo_node_cellxfs_xf            TYPE REF TO if_ixml_element,
              lo_node_cellxfs_xf_alignment  TYPE REF TO if_ixml_element,
              lo_node_cellxfs_xf_protection TYPE REF TO if_ixml_element,

              lo_nodes_xf                   TYPE REF TO if_ixml_node_collection,
              lo_iterator_cellxfs           TYPE REF TO if_ixml_node_iterator,

              ls_xf                         TYPE lty_xf,
              ls_alignment                  TYPE lty_alignment,
              ls_protection                 TYPE lty_protection,
              lv_index                      TYPE i.

*--------------------------------------------------------------------*
* To build a complete style that fully describes how a cell looks like
* we need the various parts
* §1 - Numberformat
* §2 - Fillstyle
* §3 - Borders
* §4 - Font
* §5 - Alignment
* §6 - Protection

*          Following is an example how this part of a file could be set up
*              ...
*              parts with various formatinformation - see §1,§2,§3,§4
*              ...
*          <cellXfs count="26">
*              <xf numFmtId="0" borderId="0" fillId="0" fontId="0" xfId="0"/>
*              <xf numFmtId="0" borderId="0" fillId="2" fontId="0" xfId="0" applyFill="1"/>
*              <xf numFmtId="0" borderId="1" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="2" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="3" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="4" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              <xf numFmtId="0" borderId="0" fillId="3" fontId="0" xfId="0" applyFill="1" applyBorder="1"/>
*              ...
*          </cellXfs>
*--------------------------------------------------------------------*

  lo_styles_xml = me->get_ixml_from_zip_archive( ip_path ).

*--------------------------------------------------------------------*
* The styles are build up from
* §1 number formats
* §2 fill styles
* §3 border styles
* §4 fonts
* These need to be read before we can try to build up a complete
* style that describes the look of a cell
*--------------------------------------------------------------------*
  lt_num_formats   = load_style_num_formats( lo_styles_xml ).   " §1
  lt_fills         = load_style_fills( lo_styles_xml ).         " §2
  lt_borders       = load_style_borders( lo_styles_xml ).       " §3
  lt_fonts         = load_style_fonts( lo_styles_xml ).         " §4

*--------------------------------------------------------------------*
* Now everything is prepared to build a "full" style
*--------------------------------------------------------------------*
  lo_node_cellxfs  = lo_styles_xml->find_from_name( name = 'cellXfs' ).
  IF lo_node_cellxfs IS BOUND.
    lo_nodes_xf         = lo_node_cellxfs->get_elements_by_tag_name( name = 'xf' ).
    lo_iterator_cellxfs = lo_nodes_xf->create_iterator( ).
    lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).
    WHILE lo_node_cellxfs_xf IS BOUND.

      lo_style = ip_excel->add_new_style( ).
      fill_struct_from_attributes( EXPORTING
                                     ip_element   =  lo_node_cellxfs_xf
                                   CHANGING
                                     cp_structure = ls_xf ).
*--------------------------------------------------------------------*
* §2 fill style
*--------------------------------------------------------------------*
      IF ls_xf-applyfill = '1' AND ls_xf-fillid IS NOT INITIAL.
        lv_index = ls_xf-fillid + 1.
        READ TABLE lt_fills INTO ls_fill INDEX lv_index.
        IF sy-subrc = 0.
          lo_style->fill = ls_fill.
        ENDIF.
      ENDIF.

*--------------------------------------------------------------------*
* §1 number format
*--------------------------------------------------------------------*
      IF ls_xf-numfmtid IS NOT INITIAL.
        READ TABLE lt_num_formats INTO ls_num_format WITH TABLE KEY id = ls_xf-numfmtid.
        IF sy-subrc = 0.
          lo_style->number_format = ls_num_format-format.
        ENDIF.
      ENDIF.

*--------------------------------------------------------------------*
* §3 border style
*--------------------------------------------------------------------*
      IF ls_xf-applyborder = '1' AND ls_xf-borderid IS NOT INITIAL.
        lv_index = ls_xf-borderid + 1.
        READ TABLE lt_borders INTO ls_cell_border INDEX lv_index.
        IF sy-subrc = 0.
          lo_style->borders = ls_cell_border.
        ENDIF.
      ENDIF.

*--------------------------------------------------------------------*
* §4 font
*--------------------------------------------------------------------*
      IF ls_xf-applyfont = '1' AND ls_xf-fontid IS NOT INITIAL.
        lv_index = ls_xf-fontid + 1.
        READ TABLE lt_fonts INTO ls_font INDEX lv_index.
        IF sy-subrc = 0.
          lo_style->font = ls_font.
        ENDIF.
      ENDIF.

*--------------------------------------------------------------------*
* §5 - Alignment
*--------------------------------------------------------------------*
      lo_node_cellxfs_xf_alignment ?= lo_node_cellxfs_xf->find_from_name( 'alignment' ).
      IF lo_node_cellxfs_xf_alignment IS BOUND.
        fill_struct_from_attributes( EXPORTING
                                       ip_element   =  lo_node_cellxfs_xf_alignment
                                     CHANGING
                                       cp_structure = ls_alignment ).
        IF ls_alignment-horizontal IS NOT INITIAL.
          lo_style->alignment->horizontal = ls_alignment-horizontal.
        ENDIF.

        IF ls_alignment-vertical IS NOT INITIAL.
          lo_style->alignment->vertical = ls_alignment-vertical.
        ENDIF.

        IF ls_alignment-textrotation IS NOT INITIAL.
          lo_style->alignment->textrotation = ls_alignment-textrotation.
        ENDIF.

        IF ls_alignment-wraptext = '1' OR ls_alignment-wraptext = 'true'.
          lo_style->alignment->wraptext = abap_true.
        ENDIF.

        IF ls_alignment-shrinktofit = '1' OR ls_alignment-shrinktofit = 'true'.
          lo_style->alignment->shrinktofit = abap_true.
        ENDIF.

        IF ls_alignment-indent IS NOT INITIAL.
          lo_style->alignment->indent = ls_alignment-indent.
        ENDIF.
      ENDIF.

*--------------------------------------------------------------------*
* §6 - Protection
*--------------------------------------------------------------------*
      lo_node_cellxfs_xf_protection ?= lo_node_cellxfs_xf->find_from_name( 'protection' ).
      IF lo_node_cellxfs_xf_protection IS BOUND.
        fill_struct_from_attributes( EXPORTING
                                       ip_element   = lo_node_cellxfs_xf_protection
                                     CHANGING
                                       cp_structure = ls_protection ).
        IF ls_protection-locked = '1' OR ls_protection-locked = 'true'.
          lo_style->protection->locked = zcl_excel_style_protection=>c_protection_locked.
        ELSE.
          lo_style->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
        ENDIF.

        IF ls_protection-hidden = '1' OR ls_protection-hidden = 'true'.
          lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_hidden.
        ELSE.
          lo_style->protection->hidden = zcl_excel_style_protection=>c_protection_unhidden.
        ENDIF.

      ENDIF.

      INSERT lo_style INTO TABLE me->styles.

      lo_node_cellxfs_xf ?= lo_iterator_cellxfs->get_next( ).

    ENDWHILE.
  ENDIF.

ENDMETHOD.


method LOAD_STYLE_BORDERS.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
  DATA:       lo_node_border                  TYPE REF TO if_ixml_element,
              lo_node_bordertype              TYPE REF TO if_ixml_element,
              lo_node_bordercolor             TYPE REF TO if_ixml_element,
              lo_cell_border                  TYPE REF TO zcl_excel_style_borders,
              lo_border                       TYPE REF TO zcl_excel_style_border,
              ls_color                        TYPE t_color.

*--------------------------------------------------------------------*
* We need a table of used borderformats to build up our styles
* §1    A cell has 4 outer borders and 2 diagonal "borders"
*       These borders can be formatted separately but the diagonal borders
*       are always being formatted the same
*       We'll parse through the <border>-tag for each of the bordertypes
* §2    and read the corresponding formatting information

*          Following is an example how this part of a file could be set up
*          <border diagonalDown="1">
*              <left style="mediumDashDotDot">
*                  <color rgb="FFFF0000"/>
*              </left>
*              <right/>
*              <top style="thick">
*                  <color rgb="FFFF0000"/>
*              </top>
*              <bottom style="thick">
*                  <color rgb="FFFF0000"/>
*              </bottom>
*              <diagonal style="thick">
*                  <color rgb="FFFF0000"/>
*              </diagonal>
*          </border>
*--------------------------------------------------------------------*
  lo_node_border ?= ip_xml->find_from_name( 'border' ).
  WHILE lo_node_border IS BOUND.

    CREATE OBJECT lo_cell_border.

*--------------------------------------------------------------------*
* Diagonal borderlines are formatted the equally.  Determine what kind of diagonal borders are present if any
*--------------------------------------------------------------------*
* DiagonalNone = 0
* DiagonalUp   = 1
* DiagonalDown = 2
* DiagonalBoth = 3
*--------------------------------------------------------------------*
    IF lo_node_border->get_attribute( 'diagonalDown' ) IS NOT INITIAL.
      add zcl_excel_style_borders=>c_diagonal_down to lo_cell_border->diagonal_mode.
    ENDIF.

    IF lo_node_border->get_attribute( 'diagonalUp' ) IS NOT INITIAL.
      add zcl_excel_style_borders=>c_diagonal_up to lo_cell_border->diagonal_mode.
    ENDIF.

    lo_node_bordertype ?= lo_node_border->get_first_child( ).
    WHILE lo_node_bordertype IS BOUND.
*--------------------------------------------------------------------*
* §1 Determine what kind of border we are talking about
*--------------------------------------------------------------------*
* Up, down, left, right, diagonal
*--------------------------------------------------------------------*
      CREATE OBJECT lo_border.

      CASE lo_node_bordertype->get_name( ).

        WHEN 'left'.
          lo_cell_border->left = lo_border.

        WHEN 'right'.
          lo_cell_border->right = lo_border.

        WHEN 'top'.
          lo_cell_border->top = lo_border.

        WHEN 'bottom'.
          lo_cell_border->down = lo_border.

        WHEN 'diagonal'.
          lo_cell_border->diagonal = lo_border.

      ENDCASE.

*--------------------------------------------------------------------*
* §2 Read the border-formatting
*--------------------------------------------------------------------*
      lo_border->border_style = lo_node_bordertype->get_attribute( 'style' ).
      lo_node_bordercolor ?= lo_node_bordertype->find_from_name( 'color' ).
      IF lo_node_bordercolor IS BOUND.
        fill_struct_from_attributes( EXPORTING
                                       ip_element   =  lo_node_bordercolor
                                     CHANGING
                                       cp_structure = ls_color ).

        lo_border->border_color-rgb = ls_color-rgb.
        IF ls_color-indexed IS NOT INITIAL.
          lo_border->border_color-indexed = ls_color-indexed.
        ENDIF.

        IF ls_color-theme IS NOT INITIAL.
          lo_border->border_color-theme = ls_color-theme.
        ENDIF.
        lo_border->border_color-tint = ls_color-tint.
      ENDIF.

      lo_node_bordertype ?= lo_node_bordertype->get_next( ).

    ENDWHILE.

    INSERT lo_cell_border INTO TABLE ep_borders.

    lo_node_border ?= lo_node_border->get_next( ).

  ENDWHILE.


  endmethod.


method LOAD_STYLE_FILLS.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support gradientFill
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          commenting on problems/future enhancements/todos we already know of or should decide upon
*          adding comments to explain what we are trying to achieve
*          renaming variables to indicate what they are used for
*--------------------------------------------------------------------*
  DATA:    lv_value                           TYPE string,
           lo_node_fill                       TYPE REF TO if_ixml_element,
           lo_node_fill_child                 TYPE REF TO if_ixml_element,
           lo_node_bgcolor                    TYPE REF TO if_ixml_element,
           lo_node_fgcolor                    TYPE REF TO if_ixml_element,
           lo_node_stop                       TYPE REF TO if_ixml_element,
           lo_fill                            TYPE REF TO zcl_excel_style_fill,
           ls_color                           TYPE t_color.

*--------------------------------------------------------------------*
* We need a table of used fillformats to build up our styles

*          Following is an example how this part of a file could be set up
*          <fill>
*              <patternFill patternType="gray125"/>
*          </fill>
*          <fill>
*              <patternFill patternType="solid">
*                  <fgColor rgb="FFFFFF00"/>
*                  <bgColor indexed="64"/>
*              </patternFill>
*          </fill>
*--------------------------------------------------------------------*

  lo_node_fill ?= ip_xml->find_from_name( 'fill' ).
  WHILE lo_node_fill IS BOUND.

    CREATE OBJECT lo_fill.
    lo_node_fill_child ?= lo_node_fill->get_first_child( ).
    lv_value            = lo_node_fill_child->get_name( ).
    CASE lv_value.

*--------------------------------------------------------------------*
* Patternfill
*--------------------------------------------------------------------*
      WHEN 'patternFill'.
        lo_fill->filltype = lo_node_fill_child->get_attribute( 'patternType' ).
*--------------------------------------------------------------------*
* Patternfill - background color
*--------------------------------------------------------------------*
        lo_node_bgcolor = lo_node_fill_child->find_from_name( 'bgColor' ).
        IF lo_node_bgcolor IS BOUND.
          fill_struct_from_attributes( EXPORTING
                                         ip_element   = lo_node_bgcolor
                                       CHANGING
                                         cp_structure = ls_color ).

          lo_fill->bgcolor-rgb = ls_color-rgb.
          IF ls_color-indexed IS NOT INITIAL.
            lo_fill->bgcolor-indexed = ls_color-indexed.
          ENDIF.

          IF ls_color-theme IS NOT INITIAL.
            lo_fill->bgcolor-theme = ls_color-theme.
          ENDIF.
          lo_fill->bgcolor-tint = ls_color-tint.
        ENDIF.

*--------------------------------------------------------------------*
* Patternfill - foreground color
*--------------------------------------------------------------------*
        lo_node_fgcolor = lo_node_fill->find_from_name( 'fgColor' ).
        IF lo_node_fgcolor IS BOUND.
          fill_struct_from_attributes( EXPORTING
                                         ip_element   = lo_node_fgcolor
                                       CHANGING
                                         cp_structure = ls_color ).

          lo_fill->fgcolor-rgb = ls_color-rgb.
          IF ls_color-indexed IS NOT INITIAL.
            lo_fill->fgcolor-indexed = ls_color-indexed.
          ENDIF.

          IF ls_color-theme IS NOT INITIAL.
            lo_fill->fgcolor-theme = ls_color-theme.
          ENDIF.
          lo_fill->fgcolor-tint = ls_color-tint.
        ENDIF.


*--------------------------------------------------------------------*
* gradientFill
*--------------------------------------------------------------------*
      WHEN 'gradientFill'.
       lo_fill->gradtype-type   = lo_node_fill_child->get_attribute( 'type' ).
       lo_fill->gradtype-top    = lo_node_fill_child->get_attribute( 'top' ).
       lo_fill->gradtype-left   = lo_node_fill_child->get_attribute( 'left' ).
       lo_fill->gradtype-right  = lo_node_fill_child->get_attribute( 'right' ).
       lo_fill->gradtype-bottom = lo_node_fill_child->get_attribute( 'bottom' ).
       lo_fill->gradtype-degree = lo_node_fill_child->get_attribute( 'degree' ).
       free lo_node_stop.
       lo_node_stop ?= lo_node_fill_child->find_from_name( 'stop' ).
       while lo_node_stop is bound.
         if lo_fill->gradtype-position1 is initial.
            lo_fill->gradtype-position1 = lo_node_stop->get_attribute( 'position' ).
            lo_node_bgcolor = lo_node_stop->find_from_name( 'color' ).
            IF lo_node_bgcolor IS BOUND.
               fill_struct_from_attributes( EXPORTING
                                               ip_element   = lo_node_bgcolor
                                             CHANGING
                                               cp_structure = ls_color ).

                lo_fill->bgcolor-rgb = ls_color-rgb.
                IF ls_color-indexed IS NOT INITIAL.
                  lo_fill->bgcolor-indexed = ls_color-indexed.
                ENDIF.

                IF ls_color-theme IS NOT INITIAL.
                  lo_fill->bgcolor-theme = ls_color-theme.
                ENDIF.
                lo_fill->bgcolor-tint = ls_color-tint.
            ENDIF.
         elseif lo_fill->gradtype-position2 is initial.
            lo_fill->gradtype-position2 = lo_node_stop->get_attribute( 'position' ).
            lo_node_fgcolor = lo_node_stop->find_from_name( 'color' ).
            IF lo_node_fgcolor IS BOUND.
              fill_struct_from_attributes( EXPORTING
                                             ip_element   = lo_node_fgcolor
                                           CHANGING
                                             cp_structure = ls_color ).

              lo_fill->fgcolor-rgb = ls_color-rgb.
              IF ls_color-indexed IS NOT INITIAL.
                lo_fill->fgcolor-indexed = ls_color-indexed.
              ENDIF.

              IF ls_color-theme IS NOT INITIAL.
                lo_fill->fgcolor-theme = ls_color-theme.
              ENDIF.
              lo_fill->fgcolor-tint = ls_color-tint.
            ENDIF.
          elseif lo_fill->gradtype-position3 is initial.
                lo_fill->gradtype-position3 = lo_node_stop->get_attribute( 'position' ).
                "BGColor is filled already with position 1 no need to check again
          endif.

           lo_node_stop ?= lo_node_stop->get_next( ).
        ENDWHILE.

      WHEN OTHERS.

    ENDCASE.


    INSERT lo_fill INTO TABLE ep_fills.

    lo_node_fill ?= lo_node_fill->get_next( ).

  ENDWHILE.


  endmethod.


METHOD load_style_fonts.

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          aligning code
*          removing unused variables
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
  DATA:       lo_node_font TYPE REF TO if_ixml_element,
              lo_node2     TYPE REF TO if_ixml_element,
              lo_font      TYPE REF TO zcl_excel_style_font,
              ls_color     TYPE t_color.

*--------------------------------------------------------------------*
* We need a table of used fonts to build up our styles

*          Following is an example how this part of a file could be set up
*          <font>
*              <sz val="11"/>
*              <color theme="1"/>
*              <name val="Calibri"/>
*              <family val="2"/>
*              <scheme val="minor"/>
*          </font>
*--------------------------------------------------------------------*
  lo_node_font ?= ip_xml->find_from_name( 'font' ).
  WHILE lo_node_font IS BOUND.

    CREATE OBJECT lo_font.
*--------------------------------------------------------------------*
*   Bold
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name( 'b' ) IS BOUND.
      lo_font->bold = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Italic
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name( 'i' ) IS BOUND.
      lo_font->italic = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Underline
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'u' ).
    IF lo_node2 IS BOUND.
      lo_font->underline      = abap_true.
      lo_font->underline_mode = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   StrikeThrough
*--------------------------------------------------------------------*
    IF lo_node_font->find_from_name( 'strike' ) IS BOUND.
      lo_font->strikethrough = abap_true.
    ENDIF.

*--------------------------------------------------------------------*
*   Fontsize
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'sz' ).
    IF lo_node2 IS BOUND.
      lo_font->size = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   Fontname
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'name' ).
    IF lo_node2 IS BOUND.
      lo_font->name = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   Fontfamily
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'family' ).
    IF lo_node2 IS BOUND.
      lo_font->family = lo_node2->get_attribute( 'val' ).
    ENDIF.

*--------------------------------------------------------------------*
*   Fontscheme
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'scheme' ).
    IF lo_node2 IS BOUND.
      lo_font->scheme = lo_node2->get_attribute( 'val' ).
    ELSE.
      CLEAR lo_font->scheme.
    ENDIF.

*--------------------------------------------------------------------*
*   Fontcolor
*--------------------------------------------------------------------*
    lo_node2 = lo_node_font->find_from_name( 'color' ).
    IF lo_node2 IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element   =  lo_node2
                                   CHANGING
                                     cp_structure = ls_color ).
      lo_font->color-rgb = ls_color-rgb.
      IF ls_color-indexed IS NOT INITIAL.
        lo_font->color-indexed = ls_color-indexed.
      ENDIF.

      IF ls_color-theme IS NOT INITIAL.
        lo_font->color-theme = ls_color-theme.
      ENDIF.
      lo_font->color-tint = ls_color-tint.
    ENDIF.

    INSERT lo_font INTO TABLE ep_fonts.

    lo_node_font ?= lo_node_font->get_next( ).

  ENDWHILE.


ENDMETHOD.


METHOD load_style_num_formats.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Explain gaps in predefined formats
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          adding comments to explain what we are trying to achieve
*          aligning code
*--------------------------------------------------------------------*
  DATA:       lo_node_numfmt          TYPE REF TO if_ixml_element,
              ls_num_format           TYPE ZCL_EXCEL_STYLE_NUMBER_FORMAT=>T_NUM_FORMAT .

*--------------------------------------------------------------------*
* We need a table of used numberformats to build up our styles
* there are two kinds of numberformats
* §1 built-in numberformats
* §2 and those that have been explicitly added by the createor of the excel-file
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1 built-in numberformats
*--------------------------------------------------------------------*
  ep_num_formats = ZCL_EXCEL_STYLE_NUMBER_FORMAT=>mt_built_in_num_formats.

*--------------------------------------------------------------------*
* §2   Get non-internal numberformats that are found in the file explicitly

*         Following is an example how this part of a file could be set up
*         <numFmts count="1">
*             <numFmt formatCode="#,###,###,###,##0.00" numFmtId="164"/>
*         </numFmts>
*--------------------------------------------------------------------*
  lo_node_numfmt ?= ip_xml->find_from_name( 'numFmt' ).
  WHILE lo_node_numfmt IS BOUND.

    CLEAR ls_num_format.

    CREATE OBJECT ls_num_format-format.
    ls_num_format-format->format_code = lo_node_numfmt->get_attribute( 'formatCode' ).
    ls_num_format-id                  = lo_node_numfmt->get_attribute( 'numFmtId' ).
    INSERT ls_num_format INTO TABLE ep_num_formats.

    lo_node_numfmt                          ?= lo_node_numfmt->get_next( ).

  ENDWHILE.


ENDMETHOD.


method load_theme.
  data theme type ref to zcl_excel_theme.
  data: lo_theme_xml type ref to if_ixml_document.
  create object theme.
  lo_theme_xml = me->get_ixml_from_zip_archive( iv_path ).
  theme->read_theme( io_theme_xml = lo_theme_xml  ).
  ip_excel->set_theme( io_theme = theme ).
endmethod.


method LOAD_WORKBOOK.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Move macro-reading from zcl_excel_reader_xlsm to this class
*                autodetect existance of macro/vba content
*                Allow inputparameter to explicitly tell reader to ignore vba-content
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-10
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          removing unused variables
*          adding me-> where possible
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*          renaming i/o parameters:  previous input-parameter ip_path  holds a (full) filename and not a path   --> rename to iv_workbook_full_filename
*                                                             ip_excel renamed while being at it                --> rename to io_excel
*--------------------------------------------------------------------*
* issue #232   - Read worksheetstate hidden/veryHidden
*              - Stefan Schmoecker,                          2012-11-11
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-02
* changes:    correction in named ranges to correctly attach
*             sheetlocal names/ranges to the correct sheet
*--------------------------------------------------------------------*
* issue#284 - Copied formulae ignored when reading excelfile
*           - Stefan Schmoecker,                             2013-08-02
* changes:    initialize area to hold referenced formulaedata
*             after all worksheets have been read resolve formuae
*--------------------------------------------------------------------*

  CONSTANTS:  lcv_shared_strings              TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
              lcv_worksheet                   TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
              lcv_styles                      TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
              lcv_vba_project                 TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', "#EC NEEDED     for future incorporation of XLSM-reader
              lcv_theme                       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - begin data declarations
*--------------------------------------------------------------------*
              lcv_worksheet_state_hidden      TYPE string VALUE 'hidden',
              lcv_worksheet_state_veryhidden  TYPE string VALUE 'veryHidden'.
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - end data declarations
*--------------------------------------------------------------------*

  DATA:
              lv_path                         TYPE string,
              lv_filename                     TYPE chkfile,
              lv_full_filename                TYPE string,

              lo_rels_workbook                TYPE REF TO if_ixml_document,
              lt_worksheets                   TYPE STANDARD TABLE OF t_relationship WITH NON-UNIQUE DEFAULT KEY,
              lo_workbook                     TYPE REF TO if_ixml_document,
              lv_workbook_index               TYPE i,
              lv_worksheet_path               TYPE string,
              ls_sheet                        TYPE t_sheet,

              lo_node                         TYPE REF TO if_ixml_element,
              ls_relationship                 TYPE t_relationship,
              lo_worksheet                    TYPE REF TO zcl_excel_worksheet,
              lo_range                        TYPE REF TO zcl_excel_range,
              lv_worksheet_title              TYPE zexcel_sheet_title,
              lv_tabix                        TYPE sytabix,            " #235 - repeat rows/cols.  Needed to link defined name to correct worksheet

              ls_range                        TYPE t_range,
              lv_range_value                  TYPE zexcel_range_value,
*--------------------------------------------------------------------*
* #229: Set active worksheet - begin data declarations
*--------------------------------------------------------------------*
              lv_active_sheet_string          TYPE string,
              lv_zexcel_active_worksheet      TYPE zexcel_active_worksheet,
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns  - added autofilter support while changing this section
              lo_autofilter                   TYPE REF TO zcl_excel_autofilter,
              ls_area                         TYPE zexcel_s_autofilter_area,
              lv_col_start_alpha              TYPE zexcel_cell_column_alpha,
              lv_col_end_alpha                TYPE zexcel_cell_column_alpha,
              lv_row_start                    TYPE zexcel_cell_row,
              lv_row_end                      TYPE zexcel_cell_row ,
              lv_regex                        TYPE string,
              lv_range_value_1                TYPE zexcel_range_value,
              lv_range_value_2                TYPE zexcel_range_value.
*--------------------------------------------------------------------*
* #229: Set active worksheet - end data declarations
*--------------------------------------------------------------------*
  FIELD-SYMBOLS: <worksheet> TYPE t_relationship.


*--------------------------------------------------------------------*

* §1  Get the position of files related to this workbook
*         Usually this will be <root>/xl/workbook.xml
*         Thus the workbookroot will be <root>/xl/
*         The position of all related files will be given in file
*         <workbookroot>/_rels/<workbookfilename>.rels and their positions
*         be be given relative to the workbookroot

*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Id="rId6"/>
*            <Relationship Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Id="rId5"/>
*            <Relationship Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId1"/>
*            <Relationship Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId2"/>
*            <Relationship Target="worksheets/sheet3.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId3"/>
*            <Relationship Target="worksheets/sheet4.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId4"/>
*            <Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId7"/>
*        </Relationships>
*
* §2  Load data that is relevant to the complete workbook
*     Currently supported is:
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*   §2.2    Styles          - This holds all styles that are used in all worksheets
*   §2.3    Worksheets      - For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*   §2.4    [Themes]                - not supported
*   §2.5    [VBA (Macro)]           - supported in class zcl_excel_reader_xlsm but should be moved here and autodetect
*   ...
*
* §3  Some information is held in the workbookfile as well
*   §3.1    Names and order of of worksheets
*   §3.2    Active worksheet
*   §3.3    Defined names
*   ...
*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <fileVersion rupBuild="4506" lowestEdited="4" lastEdited="4" appName="xl"/>
*            <workbookPr defaultThemeVersion="124226"/>
*            <bookViews>
*                <workbookView activeTab="1" windowHeight="8445" windowWidth="19035" yWindow="120" xWindow="120"/>
*            </bookViews>
*            <sheets>
*                <sheet r:id="rId1" sheetId="1" name="Sheet1"/>
*                <sheet r:id="rId2" sheetId="2" name="Sheet2"/>
*                <sheet r:id="rId3" sheetId="3" name="Sheet3" state="hidden"/>
*                <sheet r:id="rId4" sheetId="4" name="Sheet4"/>
*            </sheets>
*            <definedNames/>
*            <calcPr calcId="125725"/>
*        </workbook>
*--------------------------------------------------------------------*

  CLEAR me->mt_ref_formulae.                                                                              " ins issue#284

*--------------------------------------------------------------------*
* §1  Get the position of files related to this workbook
*     Entry into this method is with the filename of the workbook
*--------------------------------------------------------------------*
  CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
    EXPORTING
      full_name     = iv_workbook_full_filename
    IMPORTING
      stripped_name = lv_filename
      file_path     = lv_path.

  CONCATENATE lv_path '_rels/' lv_filename '.rels'
      INTO lv_full_filename.
  lo_rels_workbook = me->get_ixml_from_zip_archive( lv_full_filename ).

  lo_node ?= lo_rels_workbook->find_from_name( 'Relationship' ). "#EC NOTEXT
  WHILE lo_node IS BOUND.

    me->fill_struct_from_attributes( EXPORTING ip_element = lo_node CHANGING cp_structure = ls_relationship ).

    CASE ls_relationship-type.

*--------------------------------------------------------------------*
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*--------------------------------------------------------------------*
      WHEN lcv_shared_strings.
        CONCATENATE lv_path ls_relationship-target
            INTO lv_full_filename.
        me->load_shared_strings( lv_full_filename ).

*--------------------------------------------------------------------*
*   §2.3    Worksheets
*           For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*           Shared strings and styles have to be present before we can start with creating the worksheets
*           thus we only store this information for use when parsing the workbookfile for sheetinformations
*--------------------------------------------------------------------*
      WHEN lcv_worksheet.
        APPEND ls_relationship TO lt_worksheets.

*--------------------------------------------------------------------*
*   §2.2    Styles           - This holds the styles that are used in all worksheets
*--------------------------------------------------------------------*
      WHEN lcv_styles.
        CONCATENATE lv_path ls_relationship-target
            INTO lv_full_filename.
        me->load_styles( ip_path  = lv_full_filename
                         ip_excel = io_excel ).
        me->load_dxf_styles( iv_path  = lv_full_filename
                             io_excel = io_excel ).
      when lcv_theme.
        CONCATENATE lv_path ls_relationship-target
            INTO lv_full_filename.
        me->load_theme(
                exporting
                  iv_path  = lv_full_filename
                  ip_excel = io_excel   " Excel creator
                  ).
      WHEN OTHERS.

    ENDCASE.

    lo_node ?= lo_node->get_next( ).

  ENDWHILE.

*--------------------------------------------------------------------*
* §3  Some information held in the workbookfile
*--------------------------------------------------------------------*
  lo_workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).

*--------------------------------------------------------------------*
*   §3.1    Names and order of of worksheets
*--------------------------------------------------------------------*
  lo_node           ?= lo_workbook->find_from_name( 'sheet' ).
  lv_workbook_index  = 1.
  WHILE lo_node IS BOUND.

    me->fill_struct_from_attributes( EXPORTING
                                       ip_element   = lo_node
                                     CHANGING
                                       cp_structure = ls_sheet ).
*--------------------------------------------------------------------*
*       Create new worksheet in workbook with correct name
*--------------------------------------------------------------------*
    lv_worksheet_title = ls_sheet-name.
    IF lv_workbook_index = 1.                                               " First sheet has been added automatically by creating io_excel
      lo_worksheet = io_excel->get_active_worksheet( ).
      lo_worksheet->set_title( lv_worksheet_title ).
    ELSE.
      lo_worksheet = io_excel->add_new_worksheet( lv_worksheet_title ).
    ENDIF.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - begin of coding
*       Set status hidden if necessary
*--------------------------------------------------------------------*
    CASE ls_sheet-state.

      WHEN lcv_worksheet_state_hidden.
        lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_hidden.

      WHEN lcv_worksheet_state_veryhidden.
        lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_veryhidden.

    ENDCASE.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - end of coding
*--------------------------------------------------------------------*
*--------------------------------------------------------------------*
*       Load worksheetdata
*--------------------------------------------------------------------*
    READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY id = ls_sheet-id.
    IF sy-subrc = 0.
      <worksheet>-sheetid = ls_sheet-sheetid.                                "ins #235 - repeat rows/cols - needed to identify correct sheet
      CONCATENATE lv_path <worksheet>-target
          INTO lv_worksheet_path.
      me->load_worksheet( ip_path      = lv_worksheet_path
                          io_worksheet = lo_worksheet ).
      <worksheet>-worksheet = lo_worksheet.
    ENDIF.

    lo_node ?= lo_node->get_next( ).
    ADD 1 TO lv_workbook_index.

  ENDWHILE.
  SORT lt_worksheets BY sheetid.                                              " needed for localSheetid -referencing

*--------------------------------------------------------------------*
*   #284: Set active worksheet - Resolve referenced formulae to
*                                explicit formulae those cells
*--------------------------------------------------------------------*
  me->resolve_referenced_formulae( ).
                                      " ins issue#284
*--------------------------------------------------------------------*
*   #229: Set active worksheet - begin coding
*   §3.2    Active worksheet
*--------------------------------------------------------------------*
  lv_zexcel_active_worksheet = 1.                                 " First sheet = active sheet if nothing else specified.
  lo_node ?=  lo_workbook->find_from_name( 'workbookView' ).
  IF lo_node IS BOUND.
    lv_active_sheet_string = lo_node->get_attribute( 'activeTab' ).
    TRY.
        lv_zexcel_active_worksheet = lv_active_sheet_string + 1.  " EXCEL numbers the sheets from 0 onwards --> index into worksheettable is increased by one
      CATCH cx_sy_conversion_error. "#EC NO_HANDLER    - error here --> just use the default 1st sheet
    ENDTRY.
  ENDIF.
  io_excel->set_active_sheet_index( lv_zexcel_active_worksheet ).
*--------------------------------------------------------------------*
* #229: Set active worksheet - end coding
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
*   §3.3    Defined names
*           So far I have encountered these
*             - named ranges      - sheetlocal
*             - named ranges      - workbookglobal
*             - autofilters       - sheetlocal  ( special range )
*             - repeat rows/cols  - sheetlocal ( special range )
*
*--------------------------------------------------------------------*
  lo_node ?=  lo_workbook->find_from_name( 'definedName' ).
  WHILE lo_node IS BOUND.

    CLEAR lo_range.                                                                                       "ins issue #235 - repeat rows/cols
    me->fill_struct_from_attributes(  EXPORTING
                                      ip_element   =  lo_node
                                     CHANGING
                                       cp_structure = ls_range ).
    lv_range_value = lo_node->get_value( ).

    IF ls_range-localsheetid IS NOT INITIAL.                                                              " issue #163+
*      READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY id = ls_range-localsheetid.                "del issue #235 - repeat rows/cols " issue #163+
*        lo_range = <worksheet>-worksheet->add_new_range( ).                                              "del issue #235 - repeat rows/cols " issue #163+
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - begin
*--------------------------------------------------------------------*
      lv_tabix = ls_range-localsheetid + 1.
      READ TABLE lt_worksheets ASSIGNING <worksheet> INDEX lv_tabix.
      IF sy-subrc = 0.
        CASE ls_range-name.

*--------------------------------------------------------------------*
* insert autofilters
*--------------------------------------------------------------------*
          WHEN zcl_excel_autofilters=>c_autofilter.
              " begin Dennis Schaaf
              TRY.
                  zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_range_value
                                                                IMPORTING e_column_start = lv_col_start_alpha
                                                                          e_column_end   = lv_col_end_alpha
                                                                          e_row_start    = ls_area-row_start
                                                                          e_row_end      = ls_area-row_end ).
                  ls_area-col_start = zcl_excel_common=>convert_column2int( lv_col_start_alpha ).
                  ls_area-col_end   = zcl_excel_common=>convert_column2int( lv_col_end_alpha ).
                  lo_autofilter = io_excel->add_new_autofilter( io_sheet = <worksheet>-worksheet ) .
                  lo_autofilter->set_filter_area( is_area = ls_area ).
                CATCH zcx_excel.
                  " we expected a range but it was not usable, so just ignore it
              ENDTRY.
              " end Dennis Schaaf

*--------------------------------------------------------------------*
* repeat print rows/columns
*--------------------------------------------------------------------*
          WHEN zif_excel_sheet_printsettings=>gcv_print_title_name.
            lo_range = <worksheet>-worksheet->add_new_range( ).
*--------------------------------------------------------------------*
* This might be a temporary solution.  Maybe ranges get be reworked
* to support areas consisting of multiple rectangles
* But for now just split the range into row and columnpart
*--------------------------------------------------------------------*
            CLEAR:lv_range_value_1,
                  lv_range_value_2.
            IF lv_range_value IS INITIAL.
* Empty --> nothing to do
            ELSE.
              IF lv_range_value(1) = `'`.  " Escaped
                lv_regex = `^('[^']*')+![^,]*,`.
              ELSE.
                lv_regex = `^[^!]*![^,]*,`.
              ENDIF.
* Split into two ranges if necessary
              FIND REGEX lv_regex IN lv_range_value MATCH LENGTH sy-fdpos.
              IF sy-subrc = 0 AND sy-fdpos > 0.
                lv_range_value_2 = lv_range_value+sy-fdpos.
                SUBTRACT 1 FROM sy-fdpos.
                lv_range_value_1 = lv_range_value(sy-fdpos).
              ELSE.
                lv_range_value_1 = lv_range_value.
              ENDIF.
            ENDIF.
* 1st range
            zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_range_value_1
                                                          IMPORTING e_column_start = lv_col_start_alpha
                                                                    e_column_end   = lv_col_end_alpha
                                                                    e_row_start    = lv_row_start
                                                                    e_row_end      = lv_row_end ).
            IF lv_col_start_alpha IS NOT INITIAL.
              <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                    iv_columns_to   = lv_col_end_alpha ).
            ENDIF.
            IF lv_row_start IS NOT INITIAL.
              <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                 iv_rows_to   = lv_row_end ).
            ENDIF.

* 2nd range
            zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_range_value_2
                                                          IMPORTING e_column_start = lv_col_start_alpha
                                                                    e_column_end   = lv_col_end_alpha
                                                                    e_row_start    = lv_row_start
                                                                    e_row_end      = lv_row_end ).
            IF lv_col_start_alpha IS NOT INITIAL.
              <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                    iv_columns_to   = lv_col_end_alpha ).
            ENDIF.
            IF lv_row_start IS NOT INITIAL.
              <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                 iv_rows_to   = lv_row_end ).
            ENDIF.

          WHEN OTHERS.

        ENDCASE.
      ENDIF.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - end
*--------------------------------------------------------------------*
    ELSE.                                                                                                 " issue #163+
      lo_range = io_excel->add_new_range( ).                                                              " issue #163+
    ENDIF.                                                                                                " issue #163+
*    lo_range = ip_excel->add_new_range( ).                                                               " issue #163-
    IF lo_range IS BOUND.                                                                                 "ins issue #235 - repeat rows/cols
      lo_range->name = ls_range-name.
      lo_range->set_range_value( lv_range_value ).
    ENDIF.                                                                                                "ins issue #235 - repeat rows/cols
    lo_node ?= lo_node->get_next( ).

  ENDWHILE.

  endmethod.


METHOD load_worksheet.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Header/footer
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,
*              - ...
* changes: renaming variables to naming conventions
*          aligning code                                            (started)
*          add a list of open ToDos here
*          adding comments to explain what we are trying to achieve (started)
*--------------------------------------------------------------------*
* issue #345 - Dump on small pagemargins
*              Took the chance to modularize this very long method
*              by extracting the code that needed correction into
*              own method ( load_worksheet_pagemargins )
*--------------------------------------------------------------------*
  TYPES: BEGIN OF lty_cell,
           r TYPE string,
           t TYPE string,
           s TYPE string,
         END OF lty_cell.

  TYPES: BEGIN OF lty_column,
           min          TYPE string,
           max          TYPE string,
           width        TYPE float,
           customwidth  TYPE string,
           style        TYPE string,
           bestfit      TYPE string,
           collapsed    TYPE string,
           hidden       TYPE string,
           outlinelevel TYPE string,
         END OF lty_column.

  TYPES: BEGIN OF lty_sheetview,
           showgridlines     TYPE zexcel_show_gridlines,
           tabselected       TYPE string,
           zoomscalenormal   TYPE string,
           workbookviewid    TYPE string,
           showrowcolheaders TYPE string,
         END OF lty_sheetview.

  TYPES: BEGIN OF lty_mergecell,
           ref TYPE string,
         END OF lty_mergecell.

  TYPES: BEGIN OF lty_row,
           r            TYPE string,
           customheight TYPE string,
           ht           TYPE float,
           spans        TYPE string,
           thickbot     TYPE string,
           customformat TYPE string,
           thicktop     TYPE string,
           collapsed    TYPE string,
           hidden       TYPE string,
           outlinelevel TYPE string,
         END OF lty_row.

  TYPES: BEGIN OF lty_page_setup,
           id          TYPE string,
           orientation TYPE string,
           scale       TYPE string,
           fittoheight TYPE string,
           fittowidth  TYPE string,
           papersize   TYPE string,
         END OF lty_page_setup.

  TYPES: BEGIN OF lty_sheetformatpr,
           customheight     TYPE string,
           defaultrowheight TYPE string,
           customwidth      TYPE string,
           defaultcolwidth  TYPE string,
         END OF lty_sheetformatpr.

  TYPES: BEGIN OF lty_headerfooter,
           alignwithmargins TYPE string,
           differentoddeven TYPE string,
         END OF lty_headerfooter.

  TYPES: BEGIN OF lty_tabcolor,
           rgb   TYPE string,
           theme TYPE string,
         END OF lty_tabcolor.

  TYPES: BEGIN OF lty_datavalidation,
           type             TYPE zexcel_data_val_type,
           allowblank       TYPE flag,
           showinputmessage TYPE flag,
           showerrormessage TYPE flag,
           showdropdown     TYPE flag,
           operator         TYPE zexcel_data_val_operator,
           formula1         TYPE zexcel_validation_formula1,
           formula2         TYPE zexcel_validation_formula1,
           sqref            TYPE string,
           cell_column      TYPE zexcel_cell_column_alpha,
           cell_column_to   TYPE zexcel_cell_column_alpha,
           cell_row         TYPE zexcel_cell_row,
           cell_row_to      TYPE zexcel_cell_row,
           error            TYPE string,
           errortitle       TYPE string,
           prompt           TYPE string,
           prompttitle      TYPE string,
           errorstyle       TYPE zexcel_data_val_error_style,
         END OF lty_datavalidation.



  CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
             lc_xml_attr_true_int TYPE string VALUE '1',
             lc_rel_drawing       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
             lc_rel_hyperlink     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
             lc_rel_printer       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'.

  DATA: lo_ixml_worksheet           TYPE REF TO if_ixml_document,
        lo_ixml_cells               TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator            TYPE REF TO if_ixml_node_iterator,
        lo_ixml_iterator2           TYPE REF TO if_ixml_node_iterator,
        lo_ixml_row_elem            TYPE REF TO if_ixml_element,
        lo_ixml_cell_elem           TYPE REF TO if_ixml_element,
        ls_cell                     TYPE lty_cell,
        lv_index                    TYPE i,
        lo_ixml_value_elem          TYPE REF TO if_ixml_element,
        lo_ixml_formula_elem        TYPE REF TO if_ixml_element,
        lv_cell_value               TYPE zexcel_cell_value,
        lv_cell_formula             TYPE zexcel_cell_formula,
        lv_cell_column              TYPE zexcel_cell_column_alpha,
        lv_cell_row                 TYPE zexcel_cell_row,
        lo_excel_style              TYPE REF TO zcl_excel_style,
        lv_style_guid               TYPE zexcel_cell_style,

        lo_ixml_imension_elem       TYPE REF TO if_ixml_element, "#+234
        lv_dimension_range          TYPE string,  "#+234

        lo_ixml_sheetview_elem      TYPE REF TO if_ixml_element,
        ls_sheetview                TYPE lty_sheetview,
        lo_ixml_pane_elem           TYPE REF TO if_ixml_element,
        ls_excel_pane               TYPE zexcel_pane,
        lv_pane_cell_row            TYPE zexcel_cell_row,
        lv_pane_cell_col_a          TYPE zexcel_cell_column_alpha,
        lv_pane_cell_col            TYPE zexcel_cell_column,

        lo_ixml_mergecells          TYPE REF TO if_ixml_node_collection,
        lo_ixml_mergecell_elem      TYPE REF TO if_ixml_element,
        ls_mergecell                TYPE lty_mergecell,
        lv_merge_column_start       TYPE zexcel_cell_column_alpha,
        lv_merge_column_end         TYPE zexcel_cell_column_alpha,
        lv_merge_row_start          TYPE zexcel_cell_row,
        lv_merge_row_end            TYPE zexcel_cell_row,

        lo_ixml_sheetformatpr_elem  TYPE REF TO if_ixml_element,
        ls_sheetformatpr            TYPE lty_sheetformatpr,
        lv_height                   TYPE float,

        lo_ixml_headerfooter_elem   TYPE REF TO if_ixml_element,
        ls_headerfooter             TYPE lty_headerfooter,
        ls_odd_header               TYPE zexcel_s_worksheet_head_foot,
        ls_odd_footer               TYPE zexcel_s_worksheet_head_foot,
        ls_even_header              TYPE zexcel_s_worksheet_head_foot,
        ls_even_footer              TYPE zexcel_s_worksheet_head_foot,
        lo_ixml_hf_value_elem       TYPE REF TO if_ixml_element,

        lo_ixml_pagesetup_elem      TYPE REF TO if_ixml_element,
        lo_ixml_sheetpr             TYPE REF TO if_ixml_element,
        lv_fit_to_page              TYPE string,
        ls_pagesetup                TYPE lty_page_setup,

        lo_ixml_columns             TYPE REF TO if_ixml_node_collection,
        lo_ixml_column_elem         TYPE REF TO if_ixml_element,
        ls_column                   TYPE lty_column,
        lv_column_alpha             TYPE zexcel_cell_column_alpha,
        lo_column                   TYPE REF TO zcl_excel_column,
        lv_outline_level            TYPE int4,

        lo_ixml_tabcolor            TYPE REF TO if_ixml_element,
        ls_tabcolor                 TYPE lty_tabcolor,
        ls_excel_s_tabcolor         TYPE zexcel_s_tabcolor,

        lo_ixml_rows                TYPE REF TO if_ixml_node_collection,
        ls_row                      TYPE lty_row,
        lv_max_col                  TYPE i,     "for use with SPANS element
*              lv_min_col                     TYPE i,     "for use with SPANS element                    " not in use currently
        lv_max_col_s                TYPE char10,     "for use with SPANS element
        lv_min_col_s                TYPE char10,     "for use with SPANS element
        lo_row                      TYPE REF TO zcl_excel_row,
*---    End of current code aligning -------------------------------------------------------------------

        lv_path                     TYPE string,
        lo_ixml_node                TYPE REF TO if_ixml_element,
        ls_relationship             TYPE t_relationship,
        lo_ixml_rels_worksheet      TYPE REF TO if_ixml_document,
        lv_rels_worksheet_path      TYPE string,
        lv_stripped_name            TYPE chkfile,
        lv_dirname                  TYPE string,

        lt_external_hyperlinks      TYPE gtt_external_hyperlinks,
        ls_external_hyperlink       LIKE LINE OF lt_external_hyperlinks,

        lo_ixml_datavalidations     TYPE REF TO if_ixml_node_collection,
        lo_ixml_datavalidation_elem TYPE REF TO if_ixml_element,
        ls_datavalidation           TYPE lty_datavalidation,
        lo_data_validation          TYPE REF TO zcl_excel_data_validation,
        lv_datavalidation_range     TYPE string,
        lt_datavalidation_range     TYPE TABLE OF string,
        ex                          TYPE REF TO cx_root.

*--------------------------------------------------------------------*
* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets
*--------------------------------------------------------------------*

  " Read Workbook Relationships
  CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
    EXPORTING
      full_name     = ip_path
    IMPORTING
      stripped_name = lv_stripped_name
      file_path     = lv_dirname.
  CONCATENATE lv_dirname '_rels/' lv_stripped_name '.rels'
    INTO lv_rels_worksheet_path.
  TRY.                                                                          " +#222  _rels/xxx.rels might not be present.  If not found there can be no drawings --> just ignore this section
      lo_ixml_rels_worksheet = me->get_ixml_from_zip_archive( lv_rels_worksheet_path ).
      lo_ixml_node ?= lo_ixml_rels_worksheet->find_from_name( 'Relationship' ).
    CATCH zcx_excel.                              "#EC NO_HANDLER +#222
      " +#222   No errorhandling necessary - node will be unbound if error occurs
  ENDTRY.                                                   " +#222
  WHILE lo_ixml_node IS BOUND.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_node
                                 CHANGING
                                   cp_structure = ls_relationship ).
    CONCATENATE lv_dirname ls_relationship-target INTO lv_path.
    lv_path = resolve_path( lv_path ).

    CASE ls_relationship-type.
      WHEN lc_rel_drawing.
        " Read Drawings
* Issue # 339       Not all drawings are in the path mentioned below.
*                   Some Excel elements like textfields (which we don't support ) have a drawing-part in the relationsships
*                   but no "xl/drawings/_rels/drawing____.xml.rels" part.
*                   Since we don't support these there is no need to read them.  Catching exceptions thrown
*                   in the "load_worksheet_drawing" shouldn't lead to an abortion of the reading
        TRY.
            me->load_worksheet_drawing( ip_path      = lv_path
                                      io_worksheet = io_worksheet ).
          CATCH zcx_excel. "--> then ignore it
        ENDTRY.

      WHEN lc_rel_printer.
        " Read Printer settings

      WHEN lc_rel_hyperlink.
        MOVE-CORRESPONDING ls_relationship TO ls_external_hyperlink.
        INSERT ls_external_hyperlink INTO TABLE lt_external_hyperlinks.

      WHEN OTHERS.
    ENDCASE.

    lo_ixml_node ?= lo_ixml_node->get_next( ).
  ENDWHILE.


  lo_ixml_worksheet = me->get_ixml_from_zip_archive( ip_path ).


  lo_ixml_tabcolor ?= lo_ixml_worksheet->find_from_name( 'tabColor' ).
  IF lo_ixml_tabcolor IS BOUND.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_tabcolor
                                CHANGING
                                  cp_structure = ls_tabcolor ).
* Theme not supported yet
    IF ls_tabcolor-rgb IS NOT INITIAL.
      ls_excel_s_tabcolor-rgb = ls_tabcolor-rgb.
      io_worksheet->set_tabcolor( ls_excel_s_tabcolor ).
    ENDIF.
  ENDIF.

  lo_ixml_rows = lo_ixml_worksheet->get_elements_by_tag_name( name = 'row' ).
  lo_ixml_iterator = lo_ixml_rows->create_iterator( ).
  lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_row_elem IS BOUND.

    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_row_elem
                                 CHANGING
                                   cp_structure = ls_row ).
    SPLIT ls_row-spans AT ':' INTO lv_min_col_s lv_max_col_s.
    lv_index = lv_max_col_s.
    IF lv_index > lv_max_col.
      lv_max_col = lv_index.
    ENDIF.
    lv_cell_row = ls_row-r.
      lo_row = io_worksheet->get_row( lv_cell_row ).
      IF ls_row-customheight = '1'.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_true ).
      ELSE.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_false ).
      ENDIF.

      IF   ls_row-collapsed = lc_xml_attr_true
        OR ls_row-collapsed = lc_xml_attr_true_int.
        lo_row->set_collapsed( abap_true ).
      ENDIF.

      IF   ls_row-hidden = lc_xml_attr_true
        OR ls_row-hidden = lc_xml_attr_true_int.
        lo_row->set_visible( abap_false ).
      ENDIF.

      IF ls_row-outlinelevel > ''.
*        outline_level = condense( row-outlineLevel ).  "For basis 7.02 and higher
        CONDENSE  ls_row-outlinelevel.
        lv_outline_level = ls_row-outlinelevel.
        IF lv_outline_level > 0.
          lo_row->set_outline_level( lv_outline_level ).
        ENDIF.
      ENDIF.

    lo_ixml_cells = lo_ixml_row_elem->get_elements_by_tag_name( name = 'c' ).
    lo_ixml_iterator2 = lo_ixml_cells->create_iterator( ).
    lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
    WHILE lo_ixml_cell_elem IS BOUND.
      CLEAR: lv_cell_value,
             lv_cell_formula,
             lv_style_guid.

      fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_cell_elem CHANGING cp_structure = ls_cell ).

      lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name( name = 'v' ).

      CASE ls_cell-t.
        WHEN 's'. " String values are stored as index in shared string table
          IF lo_ixml_value_elem IS BOUND.
            lv_index = lo_ixml_value_elem->get_value( ) + 1.
            READ TABLE shared_strings INTO lv_cell_value INDEX lv_index.
          ENDIF.
        WHEN 'inlineStr'. " inlineStr values are kept in special node
          lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name( name = 'is' ).
          IF lo_ixml_value_elem IS BOUND.
            lv_cell_value = lo_ixml_value_elem->get_value( ).
          ENDIF.
        WHEN OTHERS. "other types are stored directly
          IF lo_ixml_value_elem IS BOUND.
            lv_cell_value = lo_ixml_value_elem->get_value( ).
          ENDIF.
      ENDCASE.

      CLEAR lv_style_guid.
      "read style based on index
      IF ls_cell-s IS NOT INITIAL.
        lv_index = ls_cell-s + 1.
        READ TABLE styles INTO lo_excel_style INDEX lv_index.
        IF sy-subrc = 0.
          lv_style_guid = lo_excel_style->get_guid( ).
        ENDIF.
      ENDIF.

      lo_ixml_formula_elem = lo_ixml_cell_elem->find_from_name( name = 'f' ).
      IF lo_ixml_formula_elem IS BOUND.
        lv_cell_formula = lo_ixml_formula_elem->get_value( ).
*--------------------------------------------------------------------*
* Begin of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
        DATA: BEGIN OF ls_formula_attributes,
                ref TYPE string,
                si  TYPE i,
                t   TYPE string,
              END OF ls_formula_attributes,
              ls_ref_formula TYPE ty_ref_formulae.

        fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_formula_elem CHANGING cp_structure = ls_formula_attributes ).
        IF ls_formula_attributes-t = 'shared'.
          zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                              i_columnrow = ls_cell-r
                                                            IMPORTING
                                                              e_column    = lv_cell_column
                                                              e_row       = lv_cell_row ).

          TRY.
              CLEAR ls_ref_formula.
              ls_ref_formula-sheet     = io_worksheet.
              ls_ref_formula-row       = lv_cell_row.
              ls_ref_formula-column    = zcl_excel_common=>convert_column2int( lv_cell_column ).
              ls_ref_formula-si        = ls_formula_attributes-si.
              ls_ref_formula-ref       = ls_formula_attributes-ref.
              ls_ref_formula-formula   = lv_cell_formula.
              INSERT ls_ref_formula INTO TABLE me->mt_ref_formulae.
            CATCH cx_root INTO ex.
              RAISE EXCEPTION TYPE zcx_excel
                EXPORTING
                  previous = ex.
          ENDTRY.
        ENDIF.
*--------------------------------------------------------------------*
* End of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
      ENDIF.

      IF   lv_cell_value    IS NOT INITIAL
        OR lv_cell_formula  IS NOT INITIAL
        OR lv_style_guid    IS NOT INITIAL.
        zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                            i_columnrow = ls_cell-r
                                                          IMPORTING
                                                            e_column    = lv_cell_column
                                                            e_row       = lv_cell_row ).
        io_worksheet->set_cell( ip_column     = lv_cell_column  " cell_elem Column
                                ip_row        = lv_cell_row     " cell_elem row_elem
                                ip_value      = lv_cell_value   " cell_elem Value
                                ip_formula    = lv_cell_formula
                                ip_data_type  = ls_cell-t
                                ip_style      = lv_style_guid ).
      ENDIF.
      lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
    ENDWHILE.
    lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.

*--------------------------------------------------------------------*
*#234 - column width not read correctly - begin of coding
*       reason - libre office doesn't use SPAN in row - definitions
*--------------------------------------------------------------------*
  IF lv_max_col = 0.
    lo_ixml_imension_elem = lo_ixml_worksheet->find_from_name( name = 'dimension' ).
    IF lo_ixml_imension_elem IS BOUND.
      lv_dimension_range = lo_ixml_imension_elem->get_attribute( 'ref' ).
      IF lv_dimension_range CS ':'.
        REPLACE REGEX '\D+\d+:(\D+)\d+' IN lv_dimension_range WITH '$1'.  " Get max column
      ELSE.
        REPLACE REGEX '(\D+)\d+' IN lv_dimension_range WITH '$1'.  " Get max column
      ENDIF.
      lv_max_col = zcl_excel_common=>convert_column2int( lv_dimension_range ).
    ENDIF.
  ENDIF.
*--------------------------------------------------------------------*
*#234 - column width not read correctly - end of coding
*--------------------------------------------------------------------*

  "Get the customized column width
  lo_ixml_columns = lo_ixml_worksheet->get_elements_by_tag_name( name = 'col' ).
  lo_ixml_iterator = lo_ixml_columns->create_iterator( ).
  lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_column_elem IS BOUND.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_column_elem
                                 CHANGING
                                   cp_structure = ls_column ).
    lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
    IF   ls_column-customwidth   = lc_xml_attr_true
      OR ls_column-customwidth   = lc_xml_attr_true_int
      OR ls_column-bestfit       = lc_xml_attr_true
      OR ls_column-bestfit       = lc_xml_attr_true_int
      OR ls_column-collapsed     = lc_xml_attr_true
      OR ls_column-collapsed     = lc_xml_attr_true_int
      OR ls_column-hidden        = lc_xml_attr_true
      OR ls_column-hidden        = lc_xml_attr_true_int
      OR ls_column-outlinelevel  > ''
      OR ls_column-style         > ''.
      lv_index = ls_column-min.
      WHILE lv_index <= ls_column-max AND lv_index <= lv_max_col.

        lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_index ).
        lo_column =  io_worksheet->get_column( lv_column_alpha ).

        IF   ls_column-customwidth = lc_xml_attr_true
          OR ls_column-customwidth = lc_xml_attr_true_int
          OR ls_column-width       IS NOT INITIAL.          "+#234
          lo_column->set_width( ls_column-width ).
        ENDIF.

        IF   ls_column-bestfit = lc_xml_attr_true
          OR ls_column-bestfit = lc_xml_attr_true_int.
          lo_column->set_auto_size( abap_true ).
        ENDIF.

        IF   ls_column-collapsed = lc_xml_attr_true
          OR ls_column-collapsed = lc_xml_attr_true_int.
          lo_column->set_collapsed( abap_true ).
        ENDIF.

        IF   ls_column-hidden = lc_xml_attr_true
          OR ls_column-hidden = lc_xml_attr_true_int.
          lo_column->set_visible( abap_false ).
        ENDIF.

        IF ls_column-outlinelevel > ''.
*          outline_level = condense( column-outlineLevel ).
          CONDENSE ls_column-outlinelevel.
          lv_outline_level = ls_column-outlinelevel.
          IF lv_outline_level > 0.
            lo_column->set_outline_level( lv_outline_level ).
          ENDIF.
        ENDIF.

        IF ls_column-style > ''.
          sy-index = ls_column-style + 1.
          READ TABLE styles INTO lo_excel_style INDEX sy-index.
          DATA: dummy_zexcel_cell_style TYPE zexcel_cell_style.
          dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
          lo_column->set_column_style_by_guid( dummy_zexcel_cell_style ).
        ENDIF.

        ADD 1 TO lv_index.
      ENDWHILE.
    ENDIF.

* issue #367 - hide columns from
    IF ls_column-max = zcl_excel_common=>c_excel_sheet_max_col.     " Max = very right column
      IF ls_column-hidden = 1     " all hidden
        AND ls_column-min > 0.
        io_worksheet->zif_excel_sheet_properties~hide_columns_from = zcl_excel_common=>convert_column2alpha( ls_column-min ).
      ELSEIF ls_column-style > ''.
        sy-index = ls_column-style + 1.
        READ TABLE styles INTO lo_excel_style INDEX sy-index.
        dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
* Set style for remaining columns
        io_worksheet->zif_excel_sheet_properties~set_style( dummy_zexcel_cell_style ).
      ENDIF.
    ENDIF.


  ENDWHILE.

  "Now we need to get information from the sheetView node
  lo_ixml_sheetview_elem = lo_ixml_worksheet->find_from_name( name = 'sheetView' ).
  fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_sheetview_elem CHANGING cp_structure = ls_sheetview ).
  IF ls_sheetview-showgridlines IS INITIAL OR
     ls_sheetview-showgridlines = lc_xml_attr_true OR
     ls_sheetview-showgridlines = lc_xml_attr_true_int.
    "If the attribute is not specified or set to true, we will show grid lines
    ls_sheetview-showgridlines = abap_true.
  ELSE.
    ls_sheetview-showgridlines = abap_false.
  ENDIF.
  io_worksheet->set_show_gridlines( ls_sheetview-showgridlines ).


  "Add merge cell information
  lo_ixml_mergecells = lo_ixml_worksheet->get_elements_by_tag_name( name = 'mergeCell' ).
  lo_ixml_iterator = lo_ixml_mergecells->create_iterator( ).
  lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_mergecell_elem IS BOUND.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_mergecell_elem
                                 CHANGING
                                   cp_structure = ls_mergecell ).
    zcl_excel_common=>convert_range2column_a_row( EXPORTING
                                                    i_range = ls_mergecell-ref
                                                  IMPORTING
                                                    e_column_start = lv_merge_column_start
                                                    e_column_end   = lv_merge_column_end
                                                    e_row_start    = lv_merge_row_start
                                                    e_row_end      = lv_merge_row_end ).
    lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
    io_worksheet->set_merge( EXPORTING
                               ip_column_start = lv_merge_column_start
                               ip_column_end   = lv_merge_column_end
                               ip_row          = lv_merge_row_start
                               ip_row_to       = lv_merge_row_end ).
  ENDWHILE.

  " read sheet format properties
  lo_ixml_sheetformatpr_elem = lo_ixml_worksheet->find_from_name( 'sheetFormatPr' ).
  IF lo_ixml_sheetformatpr_elem IS NOT INITIAL.
    fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_sheetformatpr_elem CHANGING cp_structure = ls_sheetformatpr ).
    IF ls_sheetformatpr-customheight = '1'.
      lv_height = ls_sheetformatpr-defaultrowheight.
      lo_row = io_worksheet->get_default_row( ).
      lo_row->set_row_height( lv_height ).
    ENDIF.

    " TODO...  column
  ENDIF.

  " Read in page margins
  me->load_worksheet_pagemargins( EXPORTING
                                    io_ixml_worksheet = lo_ixml_worksheet
                                    io_worksheet      = io_worksheet ).

* FitToPage
  lo_ixml_sheetpr ?=  lo_ixml_worksheet->find_from_name( 'pageSetUpPr' ).
  IF lo_ixml_sheetpr IS BOUND.

    lv_fit_to_page = lo_ixml_sheetpr->get_attribute_ns( 'fitToPage' ).
    IF lv_fit_to_page IS NOT INITIAL.
      io_worksheet->sheet_setup->fit_to_page = 'X'.
    ENDIF.
  ENDIF.
  " Read in page setup
  lo_ixml_pagesetup_elem = lo_ixml_worksheet->find_from_name( 'pageSetup' ).
  IF lo_ixml_pagesetup_elem IS NOT INITIAL.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_pagesetup_elem
                                 CHANGING
                                   cp_structure = ls_pagesetup ).
    io_worksheet->sheet_setup->orientation = ls_pagesetup-orientation.
    io_worksheet->sheet_setup->scale = ls_pagesetup-scale.
    io_worksheet->sheet_setup->paper_size = ls_pagesetup-papersize.
    IF io_worksheet->sheet_setup->fit_to_page = 'X'.
      IF ls_pagesetup-fittowidth IS NOT INITIAL.
        io_worksheet->sheet_setup->fit_to_width = ls_pagesetup-fittowidth.
      ELSE.
        io_worksheet->sheet_setup->fit_to_width = 1.  " Default if not given - Excel doesn't write this to xml
      ENDIF.
      IF ls_pagesetup-fittoheight IS NOT INITIAL.
        io_worksheet->sheet_setup->fit_to_height = ls_pagesetup-fittoheight.
      ELSE.
        io_worksheet->sheet_setup->fit_to_height = 1. " Default if not given - Excel doesn't write this to xml
      ENDIF.
    ENDIF.
  ENDIF.



  " Read header footer
  lo_ixml_headerfooter_elem = lo_ixml_worksheet->find_from_name( 'headerFooter' ).
  IF lo_ixml_headerfooter_elem IS NOT INITIAL.
    fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_headerfooter_elem CHANGING cp_structure = ls_headerfooter ).
    io_worksheet->sheet_setup->diff_oddeven_headerfooter = ls_headerfooter-differentoddeven.

    lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name( 'oddFooter' ).
    IF lo_ixml_hf_value_elem IS NOT INITIAL.
      ls_odd_footer-left_value = lo_ixml_hf_value_elem->get_value( ).
    ENDIF.

*        2do§1   Header/footer
    " TODO.. get the rest.

    io_worksheet->sheet_setup->set_header_footer( ip_odd_header   = ls_odd_header
                                                  ip_odd_footer   = ls_odd_footer
                                                  ip_even_header  = ls_even_header
                                                  ip_even_footer  = ls_even_footer ).

  ENDIF.

  " Start fix 194 Read attributes HIDDEN, OUTLINELEVEL, COLLAPSED in ZCL_EXCEL_READER_2007
  " Read pane
  lo_ixml_pane_elem = lo_ixml_sheetview_elem->find_from_name( name = 'pane' ).
  IF lo_ixml_pane_elem IS BOUND.
    fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_pane_elem CHANGING cp_structure = ls_excel_pane ).
    " Issue #194
    " Replace REGEX with method from the common class
    zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                        i_columnrow = ls_excel_pane-topleftcell
                                                      IMPORTING
                                                        e_column    = lv_pane_cell_col_a    " Cell Column
                                                        e_row       = lv_pane_cell_row ).   " Natural number
    lv_pane_cell_col = zcl_excel_common=>convert_column2int( lv_pane_cell_col_a ).
    SUBTRACT 1 FROM: lv_pane_cell_col,
                     lv_pane_cell_row.
    IF    lv_pane_cell_col > 0
      AND lv_pane_cell_row > 0.
      io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row
                                  ip_num_columns = lv_pane_cell_col ).
    ELSEIF lv_pane_cell_row > 0.
      io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row ).
    ELSE.
      io_worksheet->freeze_panes( ip_num_columns = lv_pane_cell_col ).
    ENDIF.
  ENDIF.
  " End fix 194 Read attributes HIDDEN, OUTLINELEVEL, COLLAPSED in ZCL_EXCEL_READER_2007

  " Start fix 276 Read data validations
  lo_ixml_datavalidations = lo_ixml_worksheet->get_elements_by_tag_name( name = 'dataValidation' ).
  lo_ixml_iterator = lo_ixml_datavalidations->create_iterator( ).
  lo_ixml_datavalidation_elem  ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_datavalidation_elem  IS BOUND.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_datavalidation_elem
                                 CHANGING
                                   cp_structure = ls_datavalidation ).
    CLEAR lo_ixml_formula_elem.
    lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name( name = 'formula1' ).
    IF lo_ixml_formula_elem IS BOUND.
      ls_datavalidation-formula1 = lo_ixml_formula_elem->get_value( ).
    ENDIF.
    CLEAR lo_ixml_formula_elem.
    lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name( name = 'formula2' ).
    IF lo_ixml_formula_elem IS BOUND.
      ls_datavalidation-formula2 = lo_ixml_formula_elem->get_value( ).
    ENDIF.
    SPLIT ls_datavalidation-sqref AT space INTO TABLE lt_datavalidation_range.
    LOOP AT lt_datavalidation_range INTO lv_datavalidation_range.
      zcl_excel_common=>convert_range2column_a_row( EXPORTING
                                                      i_range = lv_datavalidation_range
                                                    IMPORTING
                                                      e_column_start = ls_datavalidation-cell_column
                                                      e_column_end   = ls_datavalidation-cell_column_to
                                                      e_row_start    = ls_datavalidation-cell_row
                                                      e_row_end      = ls_datavalidation-cell_row_to ).
      lo_data_validation                   = io_worksheet->add_new_data_validation( ).
      lo_data_validation->type             = ls_datavalidation-type.
      lo_data_validation->allowblank       = ls_datavalidation-allowblank.
      IF ls_datavalidation-showinputmessage IS INITIAL.
        lo_data_validation->showinputmessage = abap_false.
      ELSE.
        lo_data_validation->showinputmessage = abap_true.
      ENDIF.
      IF ls_datavalidation-showerrormessage IS INITIAL.
        lo_data_validation->showerrormessage = abap_false.
      ELSE.
        lo_data_validation->showerrormessage = abap_true.
      ENDIF.
      IF ls_datavalidation-showdropdown IS INITIAL.
        lo_data_validation->showdropdown = abap_false.
      ELSE.
        lo_data_validation->showdropdown = abap_true.
      ENDIF.
      lo_data_validation->operator         = ls_datavalidation-operator.
      lo_data_validation->formula1         = ls_datavalidation-formula1.
      lo_data_validation->formula2         = ls_datavalidation-formula2.
      lo_data_validation->prompttitle      = ls_datavalidation-prompttitle.
      lo_data_validation->prompt           = ls_datavalidation-prompt.
      lo_data_validation->errortitle       = ls_datavalidation-errortitle.
      lo_data_validation->error            = ls_datavalidation-error.
      lo_data_validation->errorstyle       = ls_datavalidation-errorstyle.
      lo_data_validation->cell_row         = ls_datavalidation-cell_row.
      lo_data_validation->cell_row_to      = ls_datavalidation-cell_row_to.
      lo_data_validation->cell_column      = ls_datavalidation-cell_column.
      lo_data_validation->cell_column_to   = ls_datavalidation-cell_column_to.
    ENDLOOP.
    lo_ixml_datavalidation_elem ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.
  " End fix 276 Read data validations

  " Read hyperlinks
  TRY.
      me->load_worksheet_hyperlinks( io_ixml_worksheet      = lo_ixml_worksheet
                                     io_worksheet           = io_worksheet
                                     it_external_hyperlinks = lt_external_hyperlinks ).
    CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
  ENDTRY.

  TRY.
      me->fill_row_outlines( io_worksheet           = io_worksheet ).
    CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
  ENDTRY.

  " Issue #366 - conditional formatting
  TRY.
      me->load_worksheet_cond_format( io_ixml_worksheet      = lo_ixml_worksheet
                                      io_worksheet           = io_worksheet ).
    CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
  ENDTRY.

  " Issue #377 - pagebreaks
  TRY.
      me->load_worksheet_pagebreaks( io_ixml_worksheet      = lo_ixml_worksheet
                                     io_worksheet           = io_worksheet ).
    CATCH zcx_excel. " Ignore pagebreak reading errors - pass everything we were able to identify
  ENDTRY.


ENDMETHOD.


METHOD load_worksheet_cond_format.

  DATA: lo_ixml_cond_formats  TYPE REF TO if_ixml_node_collection,
        lo_ixml_cond_format   TYPE REF TO if_ixml_element,
        lo_ixml_iterator      TYPE REF TO if_ixml_node_iterator,
        lo_ixml_rules         TYPE REF TO if_ixml_node_collection,
        lo_ixml_rule          TYPE REF TO if_ixml_element,
        lo_ixml_iterator2     TYPE REF TO if_ixml_node_iterator,
        lo_style_cond  TYPE REF TO zcl_excel_style_cond,
        lo_style_cond2 TYPE REF TO zcl_excel_style_cond.


  DATA: lv_area           TYPE string,
        lt_areas          TYPE STANDARD TABLE OF string WITH NON-UNIQUE DEFAULT KEY,
        lv_area_start_row TYPE zexcel_cell_row,
        lv_area_end_row   TYPE zexcel_cell_row,
        lv_area_start_col TYPE zexcel_cell_column_alpha,
        lv_area_end_col   TYPE zexcel_cell_column_alpha,
        lv_rule           TYPE zexcel_condition_rule.


*  FIELD-SYMBOLS: <ls_external_hyperlink> LIKE LINE OF it_external_hyperlinks.

  lo_ixml_cond_formats =  io_ixml_worksheet->get_elements_by_tag_name( name = 'conditionalFormatting' ).
  lo_ixml_iterator     =  lo_ixml_cond_formats->create_iterator( ).
  lo_ixml_cond_format  ?= lo_ixml_iterator->get_next( ).

  WHILE lo_ixml_cond_format IS BOUND.

    CLEAR: lv_area,
           lo_ixml_rule,
           lo_style_cond.

*--------------------------------------------------------------------*
* Get type of rule
*--------------------------------------------------------------------*
    lo_ixml_rules       =  lo_ixml_cond_format->get_elements_by_tag_name( name = 'cfRule' ).
    lo_ixml_iterator2   =  lo_ixml_rules->create_iterator( ).
    lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).

    WHILE lo_ixml_rule IS BOUND.
      lv_rule = lo_ixml_rule->get_attribute_ns( 'type' ).
      CLEAR lo_style_cond.

*--------------------------------------------------------------------*
* Depending on ruletype get additional information
*--------------------------------------------------------------------*
      CASE lv_rule.

        WHEN zcl_excel_style_cond=>c_rule_cellis.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_ci( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_databar.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_db( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_expression.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_ex( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_iconset.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_is( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_colorscale.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_cs( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_top10.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_t10( io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).

        WHEN zcl_excel_style_cond=>c_rule_above_average.
          lo_style_cond = io_worksheet->add_new_style_cond( '' ).
          load_worksheet_cond_format_aa(  io_ixml_rule  = lo_ixml_rule
                                         io_style_cond = lo_style_cond ).
        WHEN OTHERS.
      ENDCASE.

      IF lo_style_cond IS BOUND.
        lo_style_cond->rule      = lv_rule.
        lo_style_cond->priority  = lo_ixml_rule->get_attribute_ns( 'priority' ).
*--------------------------------------------------------------------*
* Set area to which conditional formatting belongs
*--------------------------------------------------------------------*
        lv_area =  lo_ixml_cond_format->get_attribute_ns( 'sqref' ).
        SPLIT lv_area AT space INTO TABLE lt_areas.
        DELETE lt_areas WHERE table_line IS INITIAL.
        LOOP AT lt_areas INTO lv_area.

          zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_area
                                                        IMPORTING e_column_start = lv_area_start_col
                                                                  e_column_end   = lv_area_end_col
                                                                  e_row_start    = lv_area_start_row
                                                                  e_row_end      = lv_area_end_row   ).
          lo_style_cond->add_range( ip_start_column = lv_area_start_col
                                    ip_stop_column  = lv_area_end_col
                                    ip_start_row    = lv_area_start_row
                                    ip_stop_row     = lv_area_end_row   ).
        ENDLOOP.

      ENDIF.
      lo_ixml_rule        ?= lo_ixml_iterator2->get_next( ).
    ENDWHILE.


    lo_ixml_cond_format ?= lo_ixml_iterator->get_next( ).

  ENDWHILE.

ENDMETHOD.


METHOD load_worksheet_cond_format_aa.
  DATA: lv_dxf_style_index TYPE i,
        val                TYPE string.

  FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

*--------------------------------------------------------------------*
* above or below average
*--------------------------------------------------------------------*
  val  = io_ixml_rule->get_attribute_ns( 'aboveAverage' ).
  IF val = '0'.  " 0 = below average
    io_style_cond->mode_above_average-above_average = space.
  ELSE.
    io_style_cond->mode_above_average-above_average = 'X'. " Not present or <> 0 --> we use above average
  ENDIF.

*--------------------------------------------------------------------*
* Equal average also?
*--------------------------------------------------------------------*
  CLEAR val.
  val  = io_ixml_rule->get_attribute_ns( 'equalAverage' ).
  IF val = '1'.  " 0 = below average
    io_style_cond->mode_above_average-equal_average = 'X'.
  ELSE.
    io_style_cond->mode_above_average-equal_average = ' '. " Not present or <> 1 --> we use not equal average
  ENDIF.

*--------------------------------------------------------------------*
* Standard deviation instead of value ( 2nd stddev, 3rd stdev )
*--------------------------------------------------------------------*
  CLEAR val.
  val  = io_ixml_rule->get_attribute_ns( 'stdDev' ).
  CASE val.
    WHEN 1
      OR 2
      OR 3.  " These seem to be supported by excel - don't try anything more
      io_style_cond->mode_above_average-standard_deviation = val.
  ENDCASE.

*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
  lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
  READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
  IF sy-subrc = 0.
    io_style_cond->mode_above_average-cell_style = <ls_dxf_style>-guid.
  ENDIF.

ENDMETHOD.


METHOD load_worksheet_cond_format_ci.
  DATA: lo_ixml_nodes      TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator   TYPE REF TO if_ixml_node_iterator,
        lo_ixml            TYPE REF TO if_ixml_element,
        lv_dxf_style_index TYPE i,
        lo_excel_style     LIKE LINE OF me->styles.

  FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

  io_style_cond->mode_cellis-operator  = io_ixml_rule->get_attribute_ns( 'operator' ).
  lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
  READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
  IF sy-subrc = 0.
    io_style_cond->mode_cellis-cell_style = <ls_dxf_style>-guid.
  ENDIF.

  lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name( 'formula' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_cellis-formula  = lo_ixml->get_value( ).

      WHEN 2.
        io_style_cond->mode_cellis-formula2 = lo_ixml->get_value( ).

      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.


ENDMETHOD.


METHOD load_worksheet_cond_format_cs.
  DATA: lo_ixml_nodes    TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator TYPE REF TO if_ixml_node_iterator,
        lo_ixml          TYPE REF TO if_ixml_element.


  lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name( 'cfvo' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_colorscale-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_colorscale-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 2.
        io_style_cond->mode_colorscale-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_colorscale-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 3.
        io_style_cond->mode_colorscale-cfvo3_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_colorscale-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.

  lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name( 'color' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_colorscale-colorrgb1  = lo_ixml->get_attribute_ns( 'rgb' ).

      WHEN 2.
        io_style_cond->mode_colorscale-colorrgb2  = lo_ixml->get_attribute_ns( 'rgb' ).

      WHEN 3.
        io_style_cond->mode_colorscale-colorrgb3  = lo_ixml->get_attribute_ns( 'rgb' ).

      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.

ENDMETHOD.


METHOD load_worksheet_cond_format_db.
  DATA: lo_ixml_nodes        TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator     TYPE REF TO if_ixml_node_iterator,
        lo_ixml              TYPE REF TO if_ixml_element.

  lo_ixml ?= io_ixml_rule->find_from_name( 'color' ).
  IF lo_ixml IS BOUND.
    io_style_cond->mode_databar-colorrgb = lo_ixml->get_attribute_ns( 'rgb' ).
  ENDIF.

  lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name( 'cfvo' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_databar-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_databar-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 2.
        io_style_cond->mode_databar-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_databar-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.


ENDMETHOD.


METHOD load_worksheet_cond_format_ex.
  DATA: lo_ixml_nodes      TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator   TYPE REF TO if_ixml_node_iterator,
        lo_ixml            TYPE REF TO if_ixml_element,
        lv_dxf_style_index TYPE i,
        lo_excel_style     LIKE LINE OF me->styles.

  FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

  lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
  READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
  IF sy-subrc = 0.
    io_style_cond->mode_expression-cell_style = <ls_dxf_style>-guid.
  ENDIF.

  lo_ixml_nodes ?= io_ixml_rule->get_elements_by_tag_name( 'formula' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_expression-formula  = lo_ixml->get_value( ).


      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.


ENDMETHOD.


METHOD load_worksheet_cond_format_is.
  DATA: lo_ixml_nodes        TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator     TYPE REF TO if_ixml_node_iterator,
        lo_ixml              TYPE REF TO if_ixml_element,
        lo_ixml_rule_iconset TYPE REF TO if_ixml_element.

  lo_ixml_rule_iconset ?= io_ixml_rule->get_first_child( ).
  io_style_cond->mode_iconset-iconset   = lo_ixml_rule_iconset->get_attribute_ns( 'iconSet' ).
  io_style_cond->mode_iconset-showvalue = lo_ixml_rule_iconset->get_attribute_ns( 'showValue' ).
  lo_ixml_nodes ?= lo_ixml_rule_iconset->get_elements_by_tag_name( 'cfvo' ).
  lo_ixml_iterator = lo_ixml_nodes->create_iterator( ).
  lo_ixml ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml IS BOUND.

    CASE sy-index.
      WHEN 1.
        io_style_cond->mode_iconset-cfvo1_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_iconset-cfvo1_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 2.
        io_style_cond->mode_iconset-cfvo2_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_iconset-cfvo2_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 3.
        io_style_cond->mode_iconset-cfvo3_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_iconset-cfvo3_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 4.
        io_style_cond->mode_iconset-cfvo4_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_iconset-cfvo4_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN 5.
        io_style_cond->mode_iconset-cfvo5_type  = lo_ixml->get_attribute_ns( 'type' ).
        io_style_cond->mode_iconset-cfvo5_value = lo_ixml->get_attribute_ns( 'val' ).

      WHEN OTHERS.
        EXIT.
    ENDCASE.

    lo_ixml ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.

ENDMETHOD.


METHOD load_worksheet_cond_format_t10.
  DATA: lv_dxf_style_index TYPE i.

  FIELD-SYMBOLS: <ls_dxf_style> LIKE LINE OF me->mt_dxf_styles.

  io_style_cond->mode_top10-topxx_count  = io_ixml_rule->get_attribute_ns( 'rank' ).        " Top10, Top20, Top 50...

  io_style_cond->mode_top10-percent      = io_ixml_rule->get_attribute_ns( 'percent' ).     " Top10 percent instead of Top10 values
  IF io_style_cond->mode_top10-percent = '1'.
    io_style_cond->mode_top10-percent = 'X'.
  ELSE.
    io_style_cond->mode_top10-percent = ' '.
  ENDIF.

  io_style_cond->mode_top10-bottom       = io_ixml_rule->get_attribute_ns( 'bottom' ).      " Bottom10 instead of Top10
  IF io_style_cond->mode_top10-bottom = '1'.
    io_style_cond->mode_top10-bottom = 'X'.
  ELSE.
    io_style_cond->mode_top10-bottom = ' '.
  ENDIF.
*--------------------------------------------------------------------*
* Cell formatting for top10
*--------------------------------------------------------------------*
  lv_dxf_style_index  = io_ixml_rule->get_attribute_ns( 'dxfId' ).
  READ TABLE me->mt_dxf_styles ASSIGNING <ls_dxf_style> WITH KEY dxf = lv_dxf_style_index.
  IF sy-subrc = 0.
    io_style_cond->mode_top10-cell_style = <ls_dxf_style>-guid.
  ENDIF.

ENDMETHOD.


method LOAD_WORKSHEET_DRAWING.

  TYPES: BEGIN OF t_c_nv_pr,
    name TYPE string,
    id TYPE string,
    END OF t_c_nv_pr.

  TYPES: BEGIN OF t_blip,
    cstate TYPE string,
    embed TYPE string,
    END OF t_blip.

  TYPES: BEGIN OF t_chart,
    id TYPE string,
    END OF t_chart.

  CONSTANTS: lc_xml_attr_true                TYPE string VALUE 'true',
             lc_xml_attr_true_int            TYPE string VALUE '1'.
  CONSTANTS: lc_rel_chart       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
             lc_rel_image       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'.

  DATA: drawing             TYPE REF TO if_ixml_document,
        anchors             TYPE REF TO if_ixml_node_collection,
        node                TYPE REF TO if_ixml_element,
        coll_length         TYPE i,
        iterator            TYPE REF TO if_ixml_node_iterator,
        anchor_elem         TYPE REF TO if_ixml_element,

        relationship        TYPE t_relationship,
        rel_drawings        TYPE t_rel_drawings,
        rel_drawing         TYPE t_rel_drawing,
        rels_drawing        TYPE REF TO if_ixml_document,
        rels_drawing_path   TYPE string,
        stripped_name       TYPE chkfile,
        dirname             TYPE string,

        path                TYPE string,
        path2               TYPE text255,
        file_ext2           TYPE char10.

  " Read Workbook Relationships
  CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
    EXPORTING
      full_name     = ip_path
    IMPORTING
      stripped_name = stripped_name
      file_path     = dirname.
  CONCATENATE dirname '_rels/' stripped_name '.rels'
    INTO rels_drawing_path.
  rels_drawing_path = resolve_path( rels_drawing_path ).
  rels_drawing = me->get_ixml_from_zip_archive( rels_drawing_path ).
  node ?= rels_drawing->find_from_name( 'Relationship' ).
  WHILE node IS BOUND.
    fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = relationship ).

    rel_drawing-id = relationship-id.

    CONCATENATE dirname relationship-target INTO path.
    path = resolve_path( path ).
    rel_drawing-content = me->get_from_zip_archive( path ). "------------> This is for template usage

    path2 = path.
    zcl_excel_common=>split_file( EXPORTING ip_file = path2
                                  IMPORTING ep_extension = file_ext2 ).
    rel_drawing-file_ext = file_ext2.

    "-------------Added by Alessandro Iannacci - Should load graph xml
    CASE relationship-type.
      WHEN lc_rel_chart.
        "Read chart xml
        rel_drawing-content_xml = me->get_ixml_from_zip_archive( path ).
      WHEN OTHERS.
    ENDCASE.
    "----------------------------


    APPEND rel_drawing TO rel_drawings.

    node ?= node->get_next( ).
  ENDWHILE.

  drawing = me->get_ixml_from_zip_archive( ip_path ).

* one-cell anchor **************
  anchors = drawing->get_elements_by_tag_name( name = 'oneCellAnchor' namespace = 'xdr' ).
  coll_length = anchors->get_length( ).
  iterator = anchors->create_iterator( ).
  DO coll_length TIMES.
    anchor_elem ?= iterator->get_next( ).

    CALL METHOD me->load_drawing_anchor
      EXPORTING
        io_anchor_element   = anchor_elem
        io_worksheet        = io_worksheet
        it_related_drawings = rel_drawings.

  ENDDO.

* two-cell anchor ******************
  anchors = drawing->get_elements_by_tag_name( name = 'twoCellAnchor' namespace = 'xdr' ).
  coll_length = anchors->get_length( ).
  iterator = anchors->create_iterator( ).
  DO coll_length TIMES.
    anchor_elem ?= iterator->get_next( ).

    CALL METHOD me->load_drawing_anchor
      EXPORTING
        io_anchor_element   = anchor_elem
        io_worksheet        = io_worksheet
        it_related_drawings = rel_drawings.

  ENDDO.

  endmethod.


METHOD load_worksheet_hyperlinks.

  DATA: lo_ixml_hyperlinks TYPE REF TO if_ixml_node_collection,
        lo_ixml_hyperlink  TYPE REF TO if_ixml_element,
        lo_ixml_iterator   TYPE REF TO if_ixml_node_iterator,
        lv_row             TYPE zexcel_cell_row,
        lv_column          TYPE zexcel_cell_column_alpha,
        lo_hyperlink       TYPE REF TO zcl_excel_hyperlink,
        lv_value           TYPE zexcel_cell_value.

  DATA: BEGIN OF ls_hyperlink,
          ref      TYPE string,
          display  TYPE string,
          location TYPE string,
          tooltip  TYPE string,
          r_id     TYPE string,
        END OF ls_hyperlink.

  FIELD-SYMBOLS: <ls_external_hyperlink> LIKE LINE OF it_external_hyperlinks.

  lo_ixml_hyperlinks =  io_ixml_worksheet->get_elements_by_tag_name( name = 'hyperlink' ).
  lo_ixml_iterator   =  lo_ixml_hyperlinks->create_iterator( ).
  lo_ixml_hyperlink  ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_hyperlink IS BOUND.

    CLEAR ls_hyperlink.
    CLEAR lo_hyperlink.

    ls_hyperlink-ref      = lo_ixml_hyperlink->get_attribute_ns( 'ref' ).
    ls_hyperlink-display  = lo_ixml_hyperlink->get_attribute_ns( 'display' ).
    ls_hyperlink-location = lo_ixml_hyperlink->get_attribute_ns( 'location' ).
    ls_hyperlink-tooltip  = lo_ixml_hyperlink->get_attribute_ns( 'tooltip' ).
    ls_hyperlink-r_id     = lo_ixml_hyperlink->get_attribute( name      = 'id'
                                                              namespace = 'r' ).
    IF ls_hyperlink-r_id IS INITIAL.  " Internal link
      lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = ls_hyperlink-location ).
    ELSE.                             " External link
      READ TABLE it_external_hyperlinks ASSIGNING <ls_external_hyperlink> WITH TABLE KEY id = ls_hyperlink-r_id.
      IF sy-subrc = 0.
        lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = <ls_external_hyperlink>-target ).
      ENDIF.
    ENDIF.
    IF lo_hyperlink IS BOUND.  " because of unsupported external links

      zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                          i_columnrow = ls_hyperlink-ref
                                                        IMPORTING
                                                          e_row       = lv_row
                                                          e_column    = lv_column ).
* Currently it is not allowed to pass a hyperlink w/o text, but text has already been read.
* So just reread it and be done with it
      io_worksheet->get_cell( EXPORTING
                                ip_column     = lv_column
                                 ip_row       = lv_row
                               IMPORTING
                                 ep_value     = lv_value ).
      io_worksheet->set_cell( ip_column     = lv_column
                              ip_row        = lv_row
                              ip_value      = lv_value
                              ip_hyperlink  = lo_hyperlink ).
    ENDIF.

    lo_ixml_hyperlink ?= lo_ixml_iterator->get_next( ).

  ENDWHILE.


ENDMETHOD.


METHOD load_worksheet_pagebreaks.

  DATA: lo_node              TYPE REF TO if_ixml_element,
        lo_ixml_rowbreaks    TYPE REF TO if_ixml_node_collection,
        lo_ixml_colbreaks    TYPE REF TO if_ixml_node_collection,
        lo_ixml_iterator     TYPE REF TO if_ixml_node_iterator,
        lo_ixml_rowbreak     TYPE REF TO if_ixml_element,
        lo_ixml_colbreak     TYPE REF TO if_ixml_element,
        lo_style_cond TYPE REF TO zcl_excel_style_cond,
        lv_count             TYPE i.


  DATA: lt_pagebreaks TYPE STANDARD TABLE OF zcl_excel_worksheet_pagebreaks=>ts_pagebreak_at,
        lo_pagebreaks TYPE REF TO zcl_excel_worksheet_pagebreaks.

  FIELD-SYMBOLS: <ls_pagebreak_row> LIKE LINE OF lt_pagebreaks.
  FIELD-SYMBOLS: <ls_pagebreak_col> LIKE LINE OF lt_pagebreaks.

*--------------------------------------------------------------------*
* Get minimal number of cells where to add pagebreaks
* Since rows and columns are handled in separate nodes
* Build table to identify these cells
*--------------------------------------------------------------------*
  lo_node ?= io_ixml_worksheet->find_from_name( 'rowBreaks' ).
  check lo_node is bound.
  lo_ixml_rowbreaks =  lo_node->get_elements_by_tag_name( name = 'brk' ).
  lo_ixml_iterator  =  lo_ixml_rowbreaks->create_iterator( ).
  lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
  WHILE lo_ixml_rowbreak IS BOUND.
    APPEND INITIAL LINE TO lt_pagebreaks ASSIGNING <ls_pagebreak_row>.
    <ls_pagebreak_row>-cell_row = lo_ixml_rowbreak->get_attribute_ns( 'id' ).

    lo_ixml_rowbreak  ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.
  CHECK <ls_pagebreak_row> IS ASSIGNED.

  lo_node ?= io_ixml_worksheet->find_from_name( 'colBreaks' ).
  check lo_node is bound.
  lo_ixml_colbreaks =  lo_node->get_elements_by_tag_name( name = 'brk' ).
  lo_ixml_iterator  =  lo_ixml_colbreaks->create_iterator( ).
  lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
  CLEAR lv_count.
  WHILE lo_ixml_colbreak IS BOUND.
    ADD 1 TO lv_count.
    READ TABLE lt_pagebreaks INDEX lv_count ASSIGNING <ls_pagebreak_col>.
    IF sy-subrc <> 0.
      APPEND INITIAL LINE TO lt_pagebreaks ASSIGNING <ls_pagebreak_col>.
      <ls_pagebreak_col>-cell_row = <ls_pagebreak_row>-cell_row.
    ENDIF.
    <ls_pagebreak_col>-cell_column = lo_ixml_colbreak->get_attribute_ns( 'id' ).

    lo_ixml_colbreak  ?= lo_ixml_iterator->get_next( ).
  ENDWHILE.
*--------------------------------------------------------------------*
* Finally add each pagebreak
*--------------------------------------------------------------------*
  lo_pagebreaks = io_worksheet->get_pagebreaks( ).
  LOOP AT lt_pagebreaks ASSIGNING <ls_pagebreak_row>.
    lo_pagebreaks->add_pagebreak( ip_column = <ls_pagebreak_row>-cell_column
                                  ip_row    = <ls_pagebreak_row>-cell_row ).
  ENDLOOP.


ENDMETHOD.


METHOD load_worksheet_pagemargins.

  TYPES:  BEGIN OF lty_page_margins,
            footer TYPE string,
            header TYPE string,
            bottom TYPE string,
            top    TYPE string,
            right  TYPE string,
            left   TYPE string,
          END OF lty_page_margins.

  DATA:lo_ixml_pagemargins_elem TYPE REF TO if_ixml_element,
       ls_pagemargins           TYPE lty_page_margins.


  lo_ixml_pagemargins_elem = io_ixml_worksheet->find_from_name( 'pageMargins' ).
  IF lo_ixml_pagemargins_elem IS NOT INITIAL.
    fill_struct_from_attributes( EXPORTING
                                   ip_element = lo_ixml_pagemargins_elem
                                 CHANGING
                                   cp_structure = ls_pagemargins ).
    io_worksheet->sheet_setup->margin_bottom = zcl_excel_common=>excel_string_to_number( ls_pagemargins-bottom ).
    io_worksheet->sheet_setup->margin_footer = zcl_excel_common=>excel_string_to_number( ls_pagemargins-footer ).
    io_worksheet->sheet_setup->margin_header = zcl_excel_common=>excel_string_to_number( ls_pagemargins-header ).
    io_worksheet->sheet_setup->margin_left   = zcl_excel_common=>excel_string_to_number( ls_pagemargins-left   ).
    io_worksheet->sheet_setup->margin_right  = zcl_excel_common=>excel_string_to_number( ls_pagemargins-right  ).
    io_worksheet->sheet_setup->margin_top    = zcl_excel_common=>excel_string_to_number( ls_pagemargins-top    ).
  ENDIF.

ENDMETHOD.


METHOD read_from_applserver.

    DATA: lv_filelength                   TYPE i,
          lt_binary_data                  TYPE STANDARD TABLE OF x255 WITH NON-UNIQUE DEFAULT KEY,
          ls_binary_data                  LIKE LINE OF lt_binary_data,
          lv_filename                     TYPE string,
          lv_max_length_line              TYPE i,
          lv_actual_length_line           TYPE i,
          lv_errormessage                 TYPE string.

    MOVE i_filename TO lv_filename.

    DESCRIBE FIELD ls_binary_data LENGTH lv_max_length_line IN BYTE MODE.
    OPEN DATASET lv_filename FOR INPUT IN BINARY MODE.
    IF sy-subrc <> 0.
      lv_errormessage = 'A problem occured when reading the file'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.
    WHILE sy-subrc = 0.

      READ DATASET lv_filename INTO ls_binary_data MAXIMUM LENGTH lv_max_length_line ACTUAL LENGTH lv_actual_length_line.
      APPEND ls_binary_data TO lt_binary_data.
      lv_filelength = lv_filelength + lv_actual_length_line.

    ENDWHILE.
    CLOSE DATASET lv_filename.

*--------------------------------------------------------------------*
* Binary data needs to be provided as XSTRING for further processing
*--------------------------------------------------------------------*
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_filelength
      IMPORTING
        buffer       = r_excel_data
      TABLES
        binary_tab   = lt_binary_data.
  ENDMETHOD.


METHOD read_from_local_file.
    DATA:  lv_filelength                   TYPE i,
           lt_binary_data                  TYPE STANDARD TABLE OF x255 WITH NON-UNIQUE DEFAULT KEY,
           ls_binary_data                  LIKE LINE OF lt_binary_data,
           lv_filename                     TYPE string,
           lv_errormessage                 TYPE string.

    MOVE i_filename TO lv_filename.

    cl_gui_frontend_services=>gui_upload( EXPORTING
                                            filename                = lv_filename
                                            filetype                = 'BIN'         " We are basically working with zipped directories --> force binary read
                                          IMPORTING
                                            filelength              = lv_filelength
                                          CHANGING
                                            data_tab                = lt_binary_data
                                          EXCEPTIONS
                                            file_open_error         = 1
                                            file_read_error         = 2
                                            no_batch                = 3
                                            gui_refuse_filetransfer = 4
                                            invalid_type            = 5
                                            no_authority            = 6
                                            unknown_error           = 7
                                            bad_data_format         = 8
                                            header_not_allowed      = 9
                                            separator_not_allowed   = 10
                                            header_too_long         = 11
                                            unknown_dp_error        = 12
                                            access_denied           = 13
                                            dp_out_of_memory        = 14
                                            disk_full               = 15
                                            dp_timeout              = 16
                                            not_supported_by_gui    = 17
                                            error_no_gui            = 18
                                            OTHERS                  = 19 ).
    IF sy-subrc <> 0.
      lv_errormessage = 'A problem occured when reading the file'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* Binary data needs to be provided as XSTRING for further processing
*--------------------------------------------------------------------*
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_filelength
      IMPORTING
        buffer       = r_excel_data
      TABLES
        binary_tab   = lt_binary_data.

  ENDMETHOD.


method RESOLVE_PATH.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Determine whether the replacement should be done
*                iterative to allow /../../..   or something alike
*        2do§2   Determine whether /./ has to be supported as well
*        2do§3   Create unit-test for this method
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: replaced previous coding by regular expression
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1  This routine will receive a path, that may have a relative pathname (/../) included somewhere
*     The output should be a resolved path without relative references
*     Example:  Input     xl/worksheets/../drawings/drawing1.xml
*               Output    xl/drawings/drawing1.xml
*--------------------------------------------------------------------*

  rp_result = ip_path.
*--------------------------------------------------------------------*
* §1  Remove relative pathnames
*--------------------------------------------------------------------*
*  Regular expression   [^/]*/\.\./
*                       [^/]*            --> any number of characters other than /
*   followed by              /\.\./      --> the sequence /../
*   ==> worksheets/../ will be found in the example
*--------------------------------------------------------------------*
  REPLACE REGEX '[^/]*/\.\./' IN rp_result WITH ``.


  endmethod.


method RESOLVE_REFERENCED_FORMULAE.
  TYPES:      BEGIN OF ty_referenced_cells,
                sheet                         TYPE REF TO zcl_excel_worksheet,
                si                            TYPE i,
                row_from                      TYPE i,
                row_to                        TYPE i,
                col_from                      TYPE i,
                col_to                        TYPE i,
                formula                       TYPE string,
                ref_cell                      TYPE char10,
              END OF ty_referenced_cells.

  DATA:       ls_ref_formula                  LIKE LINE OF me->mt_ref_formulae,
              lts_referenced_cells            TYPE SORTED TABLE OF ty_referenced_cells WITH NON-UNIQUE KEY sheet si row_from row_to col_from col_to,
              ls_referenced_cell              LIKE LINE OF lts_referenced_cells,
              lv_col_from                     TYPE zexcel_cell_column_alpha,
              lv_col_to                       TYPE zexcel_cell_column_alpha,
              lv_resulting_formula            TYPE string,
              lv_current_cell                 TYPE char10.


  me->mt_ref_formulae = me->mt_ref_formulae.

*--------------------------------------------------------------------*
* Get referenced Cells,  Build ranges for easy lookup
*--------------------------------------------------------------------*
  LOOP AT me->mt_ref_formulae INTO ls_ref_formula WHERE ref <> space.

    CLEAR ls_referenced_cell.
    ls_referenced_cell-sheet      = ls_ref_formula-sheet.
    ls_referenced_cell-si         = ls_ref_formula-si.
    ls_referenced_cell-formula    = ls_ref_formula-formula.

    TRY.
        zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = ls_ref_formula-ref
                                                      IMPORTING e_column_start = lv_col_from
                                                                e_column_end   = lv_col_to
                                                                e_row_start    = ls_referenced_cell-row_from
                                                                e_row_end      = ls_referenced_cell-row_to  ).
        ls_referenced_cell-col_from = zcl_excel_common=>convert_column2int( lv_col_from ).
        ls_referenced_cell-col_to   = zcl_excel_common=>convert_column2int( lv_col_to ).


        CLEAR ls_referenced_cell-ref_cell.
        TRY.
            ls_referenced_cell-ref_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
            ls_referenced_cell-ref_cell+3  = ls_ref_formula-row.
            CONDENSE ls_referenced_cell-ref_cell NO-GAPS.
          CATCH zcx_excel.
        ENDTRY.

        INSERT ls_referenced_cell INTO TABLE lts_referenced_cells.
      CATCH zcx_excel.
    ENDTRY.

  ENDLOOP.

*  break x0009004.
*--------------------------------------------------------------------*
* For each referencing cell determine the referenced cell
* and resolve the formula
*--------------------------------------------------------------------*
  LOOP AT me->mt_ref_formulae INTO ls_ref_formula WHERE ref = space.


    CLEAR lv_current_cell.
    TRY.
        lv_current_cell(3) = zcl_excel_common=>convert_column2alpha( ls_ref_formula-column ).
        lv_current_cell+3  = ls_ref_formula-row.
        CONDENSE lv_current_cell NO-GAPS.
      CATCH zcx_excel.
    ENDTRY.

    LOOP AT lts_referenced_cells INTO ls_referenced_cell WHERE sheet     = ls_ref_formula-sheet
                                                           AND si        = ls_ref_formula-si
                                                           AND row_from <= ls_ref_formula-row
                                                           AND row_to   >= ls_ref_formula-row
                                                           AND col_from <= ls_ref_formula-column
                                                           AND col_to   >= ls_ref_formula-column.

      TRY.

          lv_resulting_formula = zcl_excel_common=>determine_resulting_formula( iv_reference_cell     = ls_referenced_cell-ref_cell
                                                                                iv_reference_formula  = ls_referenced_cell-formula
                                                                                iv_current_cell       = lv_current_cell ).

          ls_referenced_cell-sheet->set_cell_formula( ip_column   = ls_ref_formula-column
                                                      ip_row      = ls_ref_formula-row
                                                      ip_formula  = lv_resulting_formula ).
        CATCH zcx_excel.
      ENDTRY.
      EXIT.

    ENDLOOP.

  ENDLOOP.
  endmethod.


METHOD zif_excel_reader~load.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Map Document Properties to ZCL_EXCEL
*--------------------------------------------------------------------*

  CONSTANTS:  lcv_core_properties             TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
              lcv_office_document             TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'.

  DATA:       lo_rels                         TYPE REF TO if_ixml_document,
              lo_node                         TYPE REF TO if_ixml_element,
              ls_relationship                 TYPE t_relationship.

*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller

* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets

* §3  Extracting from this the path&file where the workbook is located
*     Following is an example how this file could be set up
*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*            <Relationship Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Id="rId2"/>
*            <Relationship Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/>
*        </Relationships>
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
* §1  Create EXCEL-Object we want to return to caller
*--------------------------------------------------------------------*
  IF iv_zcl_excel_classname IS INITIAL.
    CREATE OBJECT r_excel.
  ELSE.
    CREATE OBJECT r_excel TYPE (iv_zcl_excel_classname).
  ENDIF.

  zip = create_zip_archive( i_xlsx_binary = i_excel2007
                            i_use_alternate_zip = i_use_alternate_zip ).

*--------------------------------------------------------------------*
* §2  Get file in folderstructure
*--------------------------------------------------------------------*
  lo_rels = get_ixml_from_zip_archive( '_rels/.rels' ).

*--------------------------------------------------------------------*
* §3  Cycle through the Relationship Tags and use the ones we need
*--------------------------------------------------------------------*
  lo_node ?= lo_rels->find_from_name( 'Relationship' ).     "#EC NOTEXT
  WHILE lo_node IS BOUND.

    fill_struct_from_attributes( EXPORTING
                                   ip_element   = lo_node
                                 CHANGING
                                   cp_structure = ls_relationship ).
    CASE ls_relationship-type.

      WHEN lcv_office_document.
*--------------------------------------------------------------------*
* Parse workbook - main part here
*--------------------------------------------------------------------*
        load_workbook( iv_workbook_full_filename  = ls_relationship-target
                       io_excel                   = r_excel ).

      WHEN lcv_core_properties.
        " 2do§1   Map Document Properties to ZCL_EXCEL

      WHEN OTHERS.

    ENDCASE.
    lo_node ?= lo_node->get_next( ).

  ENDWHILE.


ENDMETHOD.


METHOD zif_excel_reader~load_file.

  DATA: lv_excel_data TYPE xstring.

*--------------------------------------------------------------------*
* Read file into binary string
*--------------------------------------------------------------------*
  IF i_from_applserver = abap_true.
    lv_excel_data = read_from_applserver( i_filename ).
  ELSE.
    lv_excel_data = read_from_local_file( i_filename ).
  ENDIF.

*--------------------------------------------------------------------*
* Parse Excel data into ZCL_EXCEL object from binary string
*--------------------------------------------------------------------*
  r_excel = zif_excel_reader~load( i_excel2007            = lv_excel_data
                                   i_use_alternate_zip    = i_use_alternate_zip
                                   iv_zcl_excel_classname = iv_zcl_excel_classname ).

ENDMETHOD.
ENDCLASS.
