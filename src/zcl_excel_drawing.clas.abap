class ZCL_EXCEL_DRAWING definition
  public
  final
  create public .

public section.
*"* public components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
  type-pools ABAP .

  constants C_GRAPH_PIE type ZEXCEL_GRAPH_TYPE value 1. "#EC NOTEXT
  constants C_GRAPH_LINE type ZEXCEL_GRAPH_TYPE value 2. "#EC NOTEXT
  constants C_GRAPH_BARS type ZEXCEL_GRAPH_TYPE value 0. "#EC NOTEXT
  data GRAPH_TYPE type ZEXCEL_GRAPH_TYPE .
  data TITLE type STRING value 'image1.jpg'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  . " .
  data X_REFERENCES type CHAR1 .
  data Y_REFERENCES type CHAR1 .
  constants TYPE_IMAGE type ZEXCEL_DRAWING_TYPE value 'image'. "#EC NOTEXT
  constants TYPE_CHART type ZEXCEL_DRAWING_TYPE value 'chart'. "#EC NOTEXT
  constants ANCHOR_ABSOLUTE type ZEXCEL_DRAWING_ANCHOR value 'ABS'. "#EC NOTEXT
  constants ANCHOR_ONE_CELL type ZEXCEL_DRAWING_ANCHOR value 'ONE'. "#EC NOTEXT
  constants ANCHOR_TWO_CELL type ZEXCEL_DRAWING_ANCHOR value 'TWO'. "#EC NOTEXT
  data GRAPH type ref to ZCL_EXCEL_GRAPH .
  constants C_MEDIA_TYPE_BMP type STRING value 'bmp'. "#EC NOTEXT
  constants C_MEDIA_TYPE_XML type STRING value 'xml'. "#EC NOTEXT
  constants C_MEDIA_TYPE_JPG type STRING value 'jpg'. "#EC NOTEXT

  methods CONSTRUCTOR
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE default ZCL_EXCEL_DRAWING=>TYPE_IMAGE
      !IP_TITLE type CLIKE optional .
  methods CREATE_MEDIA_NAME
    importing
      !IP_INDEX type I .
  methods GET_FROM_COL
    returning
      value(R_FROM_COL) type ZEXCEL_CELL_COLUMN .
  methods GET_FROM_ROW
    returning
      value(R_FROM_ROW) type ZEXCEL_CELL_ROW .
  methods GET_GUID
    returning
      value(EP_GUID) type GUID_16 .
  methods GET_HEIGHT_EMU_STR
    returning
      value(R_HEIGHT) type STRING .
  methods GET_MEDIA
    returning
      value(R_MEDIA) type XSTRING .
  methods GET_MEDIA_NAME
    returning
      value(R_NAME) type STRING .
  methods GET_MEDIA_TYPE
    returning
      value(R_TYPE) type STRING .
  methods GET_NAME
    returning
      value(R_NAME) type STRING .
  methods GET_TO_COL
    returning
      value(R_TO_COL) type ZEXCEL_CELL_COLUMN .
  methods GET_TO_ROW
    returning
      value(R_TO_ROW) type ZEXCEL_CELL_ROW .
  methods GET_WIDTH_EMU_STR
    returning
      value(R_WIDTH) type STRING .
  class-methods PIXEL2EMU
    importing
      !IP_PIXEL type INT4
      !IP_DPI type INT2 optional
    returning
      value(R_EMU) type STRING .
  class-methods EMU2PIXEL
    importing
      !IP_EMU type INT4
      !IP_DPI type INT2 optional
    returning
      value(R_PIXEL) type INT4 .
  methods SET_MEDIA
    importing
      !IP_MEDIA type XSTRING optional
      !IP_MEDIA_TYPE type STRING
      !IP_WIDTH type INT4 default 0
      !IP_HEIGHT type INT4 default 0 .
  methods SET_MEDIA_MIME
    importing
      !IP_IO type SKWF_IO
      !IP_WIDTH type INT4
      !IP_HEIGHT type INT4 .
  methods SET_MEDIA_WWW
    importing
      !IP_KEY type WWWDATATAB
      !IP_WIDTH type INT4
      !IP_HEIGHT type INT4 .
  methods SET_POSITION
    importing
      !IP_FROM_ROW type ZEXCEL_CELL_ROW
      !IP_FROM_COL type ZEXCEL_CELL_COLUMN_ALPHA
      !IP_ROWOFF type INT4 optional
      !IP_COLOFF type INT4 optional .
  methods SET_POSITION2
    importing
      !IP_FROM type ZEXCEL_DRAWING_LOCATION
      !IP_TO type ZEXCEL_DRAWING_LOCATION
      !IP_ANCHOR type ZEXCEL_DRAWING_ANCHOR optional .
  methods GET_POSITION
    returning
      value(RP_POSITION) type ZEXCEL_DRAWING_POSITION .
  methods GET_TYPE
    returning
      value(RP_TYPE) type ZEXCEL_DRAWING_TYPE .
  methods GET_INDEX
    returning
      value(RP_INDEX) type STRING .
  methods LOAD_CHART_ATTRIBUTES
    importing
      value(IP_CHART) type ref to IF_IXML_DOCUMENT .
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
protected section.
private section.

*"* private components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
  data TYPE type ZEXCEL_DRAWING_TYPE value TYPE_IMAGE. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data INDEX type STRING .
  data ANCHOR type ZEXCEL_DRAWING_ANCHOR value ANCHOR_ONE_CELL. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  constants C_MEDIA_SOURCE_WWW type CHAR1 value 1. "#EC NOTEXT
  constants C_MEDIA_SOURCE_XSTRING type CHAR1 value 0. "#EC NOTEXT
  constants C_MEDIA_SOURCE_MIME type CHAR1 value 2. "#EC NOTEXT
  data GUID type GUID_16 .
  data MEDIA type XSTRING .
  data MEDIA_KEY_WWW type WWWDATATAB .
  data MEDIA_NAME type STRING .
  data MEDIA_SOURCE type CHAR1 .
  data MEDIA_TYPE type STRING .
  data IO type SKWF_IO .
  data FROM_LOC type ZEXCEL_DRAWING_LOCATION .
  data TO_LOC type ZEXCEL_DRAWING_LOCATION .
  data SIZE type ZEXCEL_DRAWING_SIZE .
ENDCLASS.



CLASS ZCL_EXCEL_DRAWING IMPLEMENTATION.


METHOD constructor.

*  CALL FUNCTION 'GUID_CREATE'                                  " del issue #379 - function is outdated in newer releases
*    IMPORTING
*      ev_guid_16 = me->guid.
  me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).      " ins issue #379 - replacement for outdated function call

  IF ip_title IS NOT INITIAL.
    title = ip_title.
  ELSE.
    title = me->guid.
  ENDIF.

  me->type = ip_type.

* inizialize dimension range
  anchor = anchor_one_cell.
  from_loc-col = 1.
  from_loc-row = 1.
ENDMETHOD.


method CREATE_MEDIA_NAME.

* if media name is initial, create unique name
  CHECK media_name IS INITIAL.

  index = ip_index.
  CONCATENATE me->type index INTO media_name.
  CONDENSE media_name NO-GAPS.
  endmethod.


METHOD emu2pixel.
* suppose 96 DPI
  IF ip_dpi IS SUPPLIED.
*    r_emu = ip_pixel  * 914400 / ip_dpi.
    r_pixel = ip_emu * ip_dpi / 914400.
  ELSE.
* suppose 96 DPI
*    r_emu = ip_pixel  * 914400 / 96.
    r_pixel = ip_emu * 96 / 914400.
  ENDIF.
ENDMETHOD.


method GET_FROM_COL.
  r_from_col = me->from_loc-col.
  endmethod.


method GET_FROM_ROW.
  r_from_row = me->from_loc-row.
  endmethod.


method GET_GUID.

  ep_guid = me->guid.

  endmethod.


method GET_HEIGHT_EMU_STR.
  r_height = pixel2emu( size-height ).
  CONDENSE r_height NO-GAPS.
  endmethod.


method GET_INDEX.
  rp_index = me->index.
  endmethod.


METHOD get_media.

  DATA: lv_language TYPE sylangu.
  DATA: lt_bin_mime TYPE sdokcntbins.
  DATA: lt_mime     TYPE tsfmime,
        lv_filesize TYPE i,
        lv_filesizec(10).

  CASE media_source.
    WHEN c_media_source_xstring.
      r_media = media.
    WHEN c_media_source_www.
      CALL FUNCTION 'WWWDATA_IMPORT'
        EXPORTING
          key    = media_key_www
        TABLES
          mime   = lt_mime
        EXCEPTIONS
          OTHERS = 1.

      CALL FUNCTION 'WWWPARAMS_READ'
        EXPORTING
          relid = media_key_www-relid
          objid = media_key_www-objid
          name  = 'filesize'
        IMPORTING
          value = lv_filesizec.

      lv_filesize = lv_filesizec.
      CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
        EXPORTING
          input_length = lv_filesize
        IMPORTING
          buffer       = r_media
        TABLES
          binary_tab   = lt_mime
        EXCEPTIONS
          failed       = 1
          OTHERS       = 2.
    WHEN c_media_source_mime.
      lv_language = sy-langu.
      cl_wb_mime_repository=>load_mime( EXPORTING
                                          io        = me->io
                                        IMPORTING
                                          filesize  = lv_filesize
                                          bin_data  = lt_bin_mime
                                        CHANGING
                                          language  = lv_language ).

      CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
        EXPORTING
          input_length = lv_filesize
        IMPORTING
          buffer       = r_media
        TABLES
          binary_tab   = lt_bin_mime
        EXCEPTIONS
          failed       = 1
          OTHERS       = 2.
  ENDCASE.
ENDMETHOD.


method GET_MEDIA_NAME.
  CONCATENATE media_name  `.` media_type INTO r_name.
  endmethod.


method GET_MEDIA_TYPE.
  r_type = media_type.
  endmethod.


method GET_NAME.
  r_name = title.
  endmethod.


method GET_POSITION.
  rp_position-anchor = anchor.
  rp_position-from = from_loc.
  rp_position-to = to_loc.
  rp_position-size = size.
  endmethod.


method GET_TO_COL.
  r_to_col = me->to_loc-col.
  endmethod.


method GET_TO_ROW.
  r_to_row = me->to_loc-row.
  endmethod.


method GET_TYPE.
  rp_type = me->type.
  endmethod.


method GET_WIDTH_EMU_STR.
  r_width = pixel2emu( size-width ).
  CONDENSE r_width NO-GAPS.
  endmethod.


METHOD load_chart_attributes.
  DATA: node                TYPE REF TO if_ixml_element.
  DATA: node2               TYPE REF TO if_ixml_element.
  DATA: node3               TYPE REF TO if_ixml_element.
  DATA: node4               TYPE REF TO if_ixml_element.

  DATA lo_barchart TYPE REF TO zcl_excel_graph_bars.
  DATA lo_piechart TYPE REF TO zcl_excel_graph_pie.
  DATA lo_linechart TYPE REF TO zcl_excel_graph_line.

  TYPES: BEGIN OF t_prop,
        val           TYPE string,
        rtl           TYPE string,
        lang          TYPE string,
        formatcode    TYPE string,
        sourcelinked  TYPE string,
  END OF t_prop.

  TYPES: BEGIN OF t_pagemargins,
        b TYPE string,
        l TYPE string,
        r TYPE string,
        t TYPE string,
        header TYPE string,
        footer TYPE string,
  END OF t_pagemargins.

  DATA ls_prop TYPE t_prop.
  DATA ls_pagemargins TYPE t_pagemargins.

  DATA lo_collection TYPE REF TO if_ixml_node_collection.
  DATA lo_node       TYPE REF TO if_ixml_node.
  DATA lo_iterator   TYPE REF TO if_ixml_node_iterator.
  DATA lv_idx        TYPE i.
  DATA lv_order      TYPE i.
  DATA lv_invertifnegative      TYPE string.
  DATA lv_symbol      TYPE string.
  DATA lv_smooth      TYPE c.
  DATA lv_sername    TYPE string.
  DATA lv_label      TYPE string.
  DATA lv_value      TYPE string.
  DATA lv_axid       TYPE string.
  DATA lv_orientation TYPE string.
  DATA lv_delete TYPE string.
  DATA lv_axpos TYPE string.
  DATA lv_formatcode TYPE string.
  DATA lv_sourcelinked TYPE string.
  DATA lv_majortickmark TYPE string.
  DATA lv_minortickmark TYPE string.
  DATA lv_ticklblpos TYPE string.
  DATA lv_crossax TYPE string.
  DATA lv_crosses TYPE string.
  DATA lv_auto TYPE string.
  DATA lv_lblalgn TYPE string.
  DATA lv_lbloffset TYPE string.
  DATA lv_nomultilvllbl TYPE string.
  DATA lv_crossbetween TYPE string.

  node ?= ip_chart->if_ixml_node~get_first_child( ).
  CHECK node IS NOT INITIAL.

  CASE me->graph_type.
    WHEN c_graph_bars.
      CREATE OBJECT lo_barchart.
      me->graph = lo_barchart.
    WHEN c_graph_pie.
      CREATE OBJECT lo_piechart.
      me->graph = lo_piechart.
    WHEN c_graph_line.
      CREATE OBJECT lo_linechart.
      me->graph = lo_linechart.
    WHEN OTHERS.
  ENDCASE.

  "Fill properties
  node2 ?= node->find_from_name( name = 'date1904' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_1904val = ls_prop-val.
  node2 ?= node->find_from_name( name = 'lang' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_langval = ls_prop-val.
  node2 ?= node->find_from_name( name = 'roundedCorners' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_roundedcornersval = ls_prop-val.

  "style
  node2 ?= node->find_from_name( name = 'style' namespace = 'c14' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_c14styleval = ls_prop-val.
  node2 ?= node->find_from_name( name = 'style' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_styleval = ls_prop-val.
  "---------------------------Read graph properties
  "ADDED
  CLEAR node2.
  node2 ?= node->find_from_name( name = 'title' namespace = 'c' ).
  IF node2 IS BOUND AND node2 IS NOT INITIAL.
    node3 ?= node2->find_from_name( name = 't' namespace = 'a' ).
    me->graph->title = node3->get_value( ).
  ENDIF.
  "END

  node2 ?= node->find_from_name( name = 'autoTitleDeleted' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_autotitledeletedval = ls_prop-val.

  "plotArea
  CASE me->graph_type.
    WHEN c_graph_bars.
      node2 ?= node->find_from_name( name = 'barDir' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_bardirval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'grouping' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_groupingval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'varyColors' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_varycolorsval = ls_prop-val.

      "Load series
      CALL METHOD node->get_elements_by_tag_name
        EXPORTING
*         depth     = 0
          name      = 'ser'
*         namespace = ''
        RECEIVING
          rval      = lo_collection.
      CALL METHOD lo_collection->create_iterator
        RECEIVING
          rval = lo_iterator.
      lo_node = lo_iterator->get_next( ).
      IF lo_node IS BOUND.
        node2 ?= lo_node->query_interface( ixml_iid_element ).
      ENDIF.
      WHILE lo_node IS BOUND.
        node3 ?= node2->find_from_name( name = 'idx' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_idx = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'order' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_order = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'invertIfNegative' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_invertifnegative = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'v' namespace = 'c' ).
        IF node3 IS BOUND.
          lv_sername = node3->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'strRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_label = node4->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'numRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_value = node4->get_value( ).
        ENDIF.
        CALL METHOD lo_barchart->create_serie
          EXPORTING
            ip_idx              = lv_idx
            ip_order            = lv_order
            ip_invertifnegative = lv_invertifnegative
            ip_lbl              = lv_label
            ip_ref              = lv_value
            ip_sername          = lv_sername.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
      ENDWHILE.
      "note: numCache avoided
      node2 ?= node->find_from_name( name = 'showLegendKey' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showlegendkeyval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showVal' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showvalval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showCatName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showcatnameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showSerName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showsernameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showPercent' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showpercentval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showBubbleSize' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_showbubblesizeval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'gapWidth' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_gapwidthval = ls_prop-val.

      "Load axes
      node2 ?= node->find_from_name( name = 'barChart' namespace = 'c' ).
      CALL METHOD node2->get_elements_by_tag_name
        EXPORTING
*         depth     = 0
          name      = 'axId'
*         namespace = ''
        RECEIVING
          rval      = lo_collection.
      CALL METHOD lo_collection->create_iterator
        RECEIVING
          rval = lo_iterator.
      lo_node = lo_iterator->get_next( ).
      IF lo_node IS BOUND.
        node2 ?= lo_node->query_interface( ixml_iid_element ).
      ENDIF.
      WHILE lo_node IS BOUND.
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lv_axid = ls_prop-val.
        IF sy-index EQ 1. "catAx
          node2 ?= node->find_from_name( name = 'catAx' namespace = 'c' ).
          node3 ?= node2->find_from_name( name = 'orientation' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_orientation = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'delete' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_delete = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'axPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_axpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'numFmt' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_formatcode = ls_prop-formatcode.
          lv_sourcelinked = ls_prop-sourcelinked.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_majortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_minortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'tickLblPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_ticklblpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossAx' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossax = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crosses' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crosses = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'auto' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_auto = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'lblAlgn' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_lblalgn = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'lblOffset' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_lbloffset = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'noMultiLvlLbl' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_nomultilvllbl = ls_prop-val.
          CALL METHOD lo_barchart->create_ax
            EXPORTING
              ip_axid          = lv_axid
              ip_type          = zcl_excel_graph_bars=>c_catax
              ip_orientation   = lv_orientation
              ip_delete        = lv_delete
              ip_axpos         = lv_axpos
              ip_formatcode    = lv_formatcode
              ip_sourcelinked  = lv_sourcelinked
              ip_majortickmark = lv_majortickmark
              ip_minortickmark = lv_minortickmark
              ip_ticklblpos    = lv_ticklblpos
              ip_crossax       = lv_crossax
              ip_crosses       = lv_crosses
              ip_auto          = lv_auto
              ip_lblalgn       = lv_lblalgn
              ip_lbloffset     = lv_lbloffset
              ip_nomultilvllbl = lv_nomultilvllbl.
        ELSEIF sy-index EQ 2. "valAx
          node2 ?= node->find_from_name( name = 'valAx' namespace = 'c' ).
          node3 ?= node2->find_from_name( name = 'orientation' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_orientation = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'delete' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_delete = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'axPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_axpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'numFmt' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_formatcode = ls_prop-formatcode.
          lv_sourcelinked = ls_prop-sourcelinked.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_majortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_minortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'tickLblPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_ticklblpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossAx' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossax = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crosses' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crosses = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossBetween' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossbetween = ls_prop-val.
          CALL METHOD lo_barchart->create_ax
            EXPORTING
              ip_axid          = lv_axid
              ip_type          = zcl_excel_graph_bars=>c_valax
              ip_orientation   = lv_orientation
              ip_delete        = lv_delete
              ip_axpos         = lv_axpos
              ip_formatcode    = lv_formatcode
              ip_sourcelinked  = lv_sourcelinked
              ip_majortickmark = lv_majortickmark
              ip_minortickmark = lv_minortickmark
              ip_ticklblpos    = lv_ticklblpos
              ip_crossax       = lv_crossax
              ip_crosses       = lv_crosses
              ip_crossbetween  = lv_crossbetween.
        ENDIF.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
      ENDWHILE.

    WHEN c_graph_pie.
      node2 ?= node->find_from_name( name = 'varyColors' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_varycolorsval = ls_prop-val.

      "Load series
      CALL METHOD node->get_elements_by_tag_name
        EXPORTING
*         depth     = 0
          name      = 'ser'
*         namespace = ''
        RECEIVING
          rval      = lo_collection.
      CALL METHOD lo_collection->create_iterator
        RECEIVING
          rval = lo_iterator.
      lo_node = lo_iterator->get_next( ).
      IF lo_node IS BOUND.
        node2 ?= lo_node->query_interface( ixml_iid_element ).
      ENDIF.
      WHILE lo_node IS BOUND.
        node3 ?= node2->find_from_name( name = 'idx' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_idx = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'order' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_order = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'v' namespace = 'c' ).
        IF node3 IS BOUND.
          lv_sername = node3->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'strRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_label = node4->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'numRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_value = node4->get_value( ).
        ENDIF.
        CALL METHOD lo_piechart->create_serie
          EXPORTING
            ip_idx     = lv_idx
            ip_order   = lv_order
            ip_lbl     = lv_label
            ip_ref     = lv_value
            ip_sername = lv_sername.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
      ENDWHILE.

      "note: numCache avoided
      node2 ?= node->find_from_name( name = 'showLegendKey' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showlegendkeyval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showVal' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showvalval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showCatName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showcatnameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showSerName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showsernameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showPercent' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showpercentval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showBubbleSize' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showbubblesizeval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showLeaderLines' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_showleaderlinesval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'firstSliceAng' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_firstsliceangval = ls_prop-val.
    WHEN c_graph_line.
      node2 ?= node->find_from_name( name = 'grouping' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_groupingval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'varyColors' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_varycolorsval = ls_prop-val.

      "Load series
      CALL METHOD node->get_elements_by_tag_name
        EXPORTING
*         depth     = 0
          name      = 'ser'
*         namespace = ''
        RECEIVING
          rval      = lo_collection.
      CALL METHOD lo_collection->create_iterator
        RECEIVING
          rval = lo_iterator.
      lo_node = lo_iterator->get_next( ).
      IF lo_node IS BOUND.
        node2 ?= lo_node->query_interface( ixml_iid_element ).
      ENDIF.
      WHILE lo_node IS BOUND.
        node3 ?= node2->find_from_name( name = 'idx' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_idx = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'order' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_order = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'symbol' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_symbol = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'smooth' namespace = 'c' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
        lv_smooth = ls_prop-val.
        node3 ?= node2->find_from_name( name = 'v' namespace = 'c' ).
        IF node3 IS BOUND.
          lv_sername = node3->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'strRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_label = node4->get_value( ).
        ENDIF.
        node3 ?= node2->find_from_name( name = 'numRef' namespace = 'c' ).
        IF node3 IS BOUND.
          node4 ?= node3->find_from_name( name = 'f' namespace = 'c' ).
          lv_value = node4->get_value( ).
        ENDIF.
        CALL METHOD lo_linechart->create_serie
          EXPORTING
            ip_idx     = lv_idx
            ip_order   = lv_order
            ip_symbol  = lv_symbol
            ip_smooth  = lv_smooth
            ip_lbl     = lv_label
            ip_ref     = lv_value
            ip_sername = lv_sername.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
      ENDWHILE.
      "note: numCache avoided
      node2 ?= node->find_from_name( name = 'showLegendKey' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showlegendkeyval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showVal' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showvalval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showCatName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showcatnameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showSerName' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showsernameval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showPercent' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showpercentval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'showBubbleSize' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_showbubblesizeval = ls_prop-val.

      node ?= node->find_from_name( name = 'lineChart' namespace = 'c' ).
      node2 ?= node->find_from_name( name = 'marker' namespace = 'c' depth = '1' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_markerval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'smooth' namespace = 'c' depth = '1' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_smoothval = ls_prop-val.
      node ?= ip_chart->if_ixml_node~get_first_child( ).
      CHECK node IS NOT INITIAL.

      "Load axes
      node2 ?= node->find_from_name( name = 'lineChart' namespace = 'c' ).
      CALL METHOD node2->get_elements_by_tag_name
        EXPORTING
*         depth     = 0
          name      = 'axId'
*         namespace = ''
        RECEIVING
          rval      = lo_collection.
      CALL METHOD lo_collection->create_iterator
        RECEIVING
          rval = lo_iterator.
      lo_node = lo_iterator->get_next( ).
      IF lo_node IS BOUND.
        node2 ?= lo_node->query_interface( ixml_iid_element ).
      ENDIF.
      WHILE lo_node IS BOUND.
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lv_axid = ls_prop-val.
        IF sy-index EQ 1. "catAx
          node2 ?= node->find_from_name( name = 'catAx' namespace = 'c' ).
          node3 ?= node2->find_from_name( name = 'orientation' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_orientation = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'delete' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_delete = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'axPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_axpos = ls_prop-val.
*          node3 ?= node2->find_from_name( name = 'numFmt' namespace = 'c' ).
*          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
*          lv_formatcode = ls_prop-formatcode.
*          lv_sourcelinked = ls_prop-sourcelinked.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_majortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_minortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'tickLblPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_ticklblpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossAx' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossax = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crosses' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crosses = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'auto' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_auto = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'lblAlgn' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_lblalgn = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'lblOffset' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_lbloffset = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'noMultiLvlLbl' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_nomultilvllbl = ls_prop-val.
          CALL METHOD lo_linechart->create_ax
            EXPORTING
              ip_axid          = lv_axid
              ip_type          = zcl_excel_graph_line=>c_catax
              ip_orientation   = lv_orientation
              ip_delete        = lv_delete
              ip_axpos         = lv_axpos
              ip_formatcode    = lv_formatcode
              ip_sourcelinked  = lv_sourcelinked
              ip_majortickmark = lv_majortickmark
              ip_minortickmark = lv_minortickmark
              ip_ticklblpos    = lv_ticklblpos
              ip_crossax       = lv_crossax
              ip_crosses       = lv_crosses
              ip_auto          = lv_auto
              ip_lblalgn       = lv_lblalgn
              ip_lbloffset     = lv_lbloffset
              ip_nomultilvllbl = lv_nomultilvllbl.
        ELSEIF sy-index EQ 2. "valAx
          node2 ?= node->find_from_name( name = 'valAx' namespace = 'c' ).
          node3 ?= node2->find_from_name( name = 'orientation' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_orientation = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'delete' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_delete = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'axPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_axpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'numFmt' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_formatcode = ls_prop-formatcode.
          lv_sourcelinked = ls_prop-sourcelinked.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_majortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'majorTickMark' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_minortickmark = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'tickLblPos' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_ticklblpos = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossAx' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossax = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crosses' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crosses = ls_prop-val.
          node3 ?= node2->find_from_name( name = 'crossBetween' namespace = 'c' ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_crossbetween = ls_prop-val.
          CALL METHOD lo_linechart->create_ax
            EXPORTING
              ip_axid          = lv_axid
              ip_type          = zcl_excel_graph_line=>c_valax
              ip_orientation   = lv_orientation
              ip_delete        = lv_delete
              ip_axpos         = lv_axpos
              ip_formatcode    = lv_formatcode
              ip_sourcelinked  = lv_sourcelinked
              ip_majortickmark = lv_majortickmark
              ip_minortickmark = lv_minortickmark
              ip_ticklblpos    = lv_ticklblpos
              ip_crossax       = lv_crossax
              ip_crosses       = lv_crosses
              ip_crossbetween  = lv_crossbetween.
        ENDIF.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
      ENDWHILE.
    WHEN OTHERS.
  ENDCASE.

  "legend
  CASE me->graph_type.
    WHEN c_graph_bars.
      node2 ?= node->find_from_name( name = 'legendPos' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_legendposval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'overlay' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_barchart->ns_overlayval = ls_prop-val.
    WHEN c_graph_line.
      node2 ?= node->find_from_name( name = 'legendPos' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_legendposval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'overlay' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_linechart->ns_overlayval = ls_prop-val.
    WHEN c_graph_pie.
      node2 ?= node->find_from_name( name = 'legendPos' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_legendposval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'overlay' namespace = 'c' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_overlayval = ls_prop-val.
      node2 ?= node->find_from_name( name = 'pPr' namespace = 'a' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_pprrtl = ls_prop-rtl.
      node2 ?= node->find_from_name( name = 'endParaRPr' namespace = 'a' ).
      zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
      lo_piechart->ns_endpararprlang = ls_prop-lang.

    WHEN OTHERS.
  ENDCASE.

  node2 ?= node->find_from_name( name = 'plotVisOnly' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_plotvisonlyval = ls_prop-val.
  node2 ?= node->find_from_name( name = 'dispBlanksAs' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_dispblanksasval = ls_prop-val.
  node2 ?= node->find_from_name( name = 'showDLblsOverMax' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
  me->graph->ns_showdlblsovermaxval = ls_prop-val.
  "---------------------

  node2 ?= node->find_from_name( name = 'pageMargins' namespace = 'c' ).
  zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_pagemargins ).
  me->graph->pagemargins = ls_pagemargins.


ENDMETHOD.


method PIXEL2EMU.
* suppose 96 DPI
  IF ip_dpi IS SUPPLIED.
    r_emu = ip_pixel  * 914400 / ip_dpi.
  ELSE.
* suppose 96 DPI
    r_emu = ip_pixel  * 914400 / 96.
  ENDIF.
  endmethod.


method SET_MEDIA.
  IF ip_media IS SUPPLIED.
    media = ip_media.
  ENDIF.
  media_type = ip_media_type.
  media_source = c_media_source_xstring.
  IF ip_width IS SUPPLIED.
    size-width  = ip_width.
  ENDIF.
  IF ip_height IS SUPPLIED.
    size-height = ip_height.
  ENDIF.
  endmethod.


METHOD set_media_mime.

  DATA: lv_language TYPE sylangu.

  io = ip_io.
  media_source = c_media_source_mime.
  size-width  = ip_width.
  size-height = ip_height.

  lv_language = sy-langu.
  cl_wb_mime_repository=>load_mime( EXPORTING
                                      io        = ip_io
                                    IMPORTING
                                      filename  = media_name
                                      "mimetype = media_type
                                    CHANGING
                                      language  = lv_language  ).

  SPLIT media_name AT '.' INTO media_name media_type.

ENDMETHOD.


method SET_MEDIA_WWW.
  DATA: lv_value(20).

  media_key_www = ip_key.
  media_source = c_media_source_www.

  CALL FUNCTION 'WWWPARAMS_READ'
    EXPORTING
      relid = media_key_www-relid
      objid = media_key_www-objid
      name  = 'fileextension'
    IMPORTING
      value = lv_value.
  media_type = lv_value.
  SHIFT media_type LEFT DELETING LEADING '.'.

  size-width  = ip_width.
  size-height = ip_height.
  endmethod.


method SET_POSITION.
  from_loc-col = zcl_excel_common=>convert_column2int( ip_from_col ) - 1.
  IF ip_coloff IS SUPPLIED.
    from_loc-col_offset = ip_coloff.
  ENDIF.
  from_loc-row = ip_from_row - 1.
  IF ip_rowoff IS SUPPLIED.
    from_loc-row_offset = ip_rowoff.
  ENDIF.
  anchor = anchor_one_cell.
  endmethod.


method SET_POSITION2.

  data: lv_anchor                     type zexcel_drawing_anchor.
  lv_anchor = ip_anchor.

  IF lv_anchor IS INITIAL.
    IF ip_to IS NOT INITIAL.
      lv_anchor = anchor_two_cell.
    ELSE.
      lv_anchor = anchor_one_cell.
    ENDIF.
  ENDIF.

  CASE lv_anchor.
    WHEN anchor_absolute OR anchor_one_cell.
      CLEAR: me->to_loc.
    WHEN anchor_two_cell.
      CLEAR: me->size.
  ENDCASE.

  me->from_loc = ip_from.
  me->to_loc = ip_to.
  me->anchor = lv_anchor.

  endmethod.
ENDCLASS.
