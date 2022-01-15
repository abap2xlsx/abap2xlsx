CLASS zcl_excel_drawing DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!

    CONSTANTS c_graph_pie TYPE zexcel_graph_type VALUE 1.   "#EC NOTEXT
    CONSTANTS c_graph_line TYPE zexcel_graph_type VALUE 2.  "#EC NOTEXT
    CONSTANTS c_graph_bars TYPE zexcel_graph_type VALUE 0.  "#EC NOTEXT
    DATA graph_type TYPE zexcel_graph_type .
    DATA title TYPE string VALUE 'image1.jpg'.              "#EC NOTEXT
    CONSTANTS type_image TYPE zexcel_drawing_type VALUE 'image'. "#EC NOTEXT
    CONSTANTS type_chart TYPE zexcel_drawing_type VALUE 'chart'. "#EC NOTEXT
    CONSTANTS anchor_absolute TYPE zexcel_drawing_anchor VALUE 'ABS'. "#EC NOTEXT
    CONSTANTS anchor_one_cell TYPE zexcel_drawing_anchor VALUE 'ONE'. "#EC NOTEXT
    CONSTANTS anchor_two_cell TYPE zexcel_drawing_anchor VALUE 'TWO'. "#EC NOTEXT
    DATA graph TYPE REF TO zcl_excel_graph .
    CONSTANTS c_media_type_bmp TYPE string VALUE 'bmp'.     "#EC NOTEXT
    CONSTANTS c_media_type_xml TYPE string VALUE 'xml'.     "#EC NOTEXT
    CONSTANTS c_media_type_jpg TYPE string VALUE 'jpg'.     "#EC NOTEXT
    CONSTANTS type_image_header_footer TYPE zexcel_drawing_type VALUE 'hd_ft'. "#EC NOTEXT
    CONSTANTS: BEGIN OF namespace,
                 c   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                 a   TYPE string VALUE 'http://schemas.openxmlformats.org/drawingml/2006/main',
                 r   TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                 mc  TYPE string VALUE 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                 c14 TYPE string VALUE 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
               END OF namespace.

    METHODS constructor
      IMPORTING
        !ip_type  TYPE zexcel_drawing_type DEFAULT zcl_excel_drawing=>type_image
        !ip_title TYPE clike OPTIONAL .
    METHODS create_media_name
      IMPORTING
        !ip_index TYPE i .
    METHODS get_from_col
      RETURNING
        VALUE(r_from_col) TYPE zexcel_cell_column .
    METHODS get_from_row
      RETURNING
        VALUE(r_from_row) TYPE zexcel_cell_row .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zexcel_guid .
    METHODS get_height_emu_str
      RETURNING
        VALUE(r_height) TYPE string .
    METHODS get_media
      RETURNING
        VALUE(r_media) TYPE xstring .
    METHODS get_media_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_media_type
      RETURNING
        VALUE(r_type) TYPE string .
    METHODS get_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_to_col
      RETURNING
        VALUE(r_to_col) TYPE zexcel_cell_column .
    METHODS get_to_row
      RETURNING
        VALUE(r_to_row) TYPE zexcel_cell_row .
    METHODS get_width_emu_str
      RETURNING
        VALUE(r_width) TYPE string .
    CLASS-METHODS pixel2emu
      IMPORTING
        !ip_pixel    TYPE int4
        !ip_dpi      TYPE int2 OPTIONAL
      RETURNING
        VALUE(r_emu) TYPE string .
    CLASS-METHODS emu2pixel
      IMPORTING
        !ip_emu        TYPE int4
        !ip_dpi        TYPE int2 OPTIONAL
      RETURNING
        VALUE(r_pixel) TYPE int4 .
    METHODS set_media
      IMPORTING
        !ip_media      TYPE xstring OPTIONAL
        !ip_media_type TYPE string
        !ip_width      TYPE int4 DEFAULT 0
        !ip_height     TYPE int4 DEFAULT 0 .
    METHODS set_media_mime
      IMPORTING
        !ip_io     TYPE skwf_io
        !ip_width  TYPE int4
        !ip_height TYPE int4 .
    METHODS set_media_www
      IMPORTING
        !ip_key    TYPE wwwdatatab
        !ip_width  TYPE int4
        !ip_height TYPE int4 .
    METHODS set_position
      IMPORTING
        !ip_from_row TYPE zexcel_cell_row
        !ip_from_col TYPE zexcel_cell_column_alpha
        !ip_rowoff   TYPE int4 OPTIONAL
        !ip_coloff   TYPE int4 OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_position2
      IMPORTING
        !ip_from   TYPE zexcel_drawing_location
        !ip_to     TYPE zexcel_drawing_location
        !ip_anchor TYPE zexcel_drawing_anchor OPTIONAL .
    METHODS get_position
      RETURNING
        VALUE(rp_position) TYPE zexcel_drawing_position .
    METHODS get_type
      RETURNING
        VALUE(rp_type) TYPE zexcel_drawing_type .
    METHODS get_index
      RETURNING
        VALUE(rp_index) TYPE string .
    METHODS load_chart_attributes
      IMPORTING
        VALUE(ip_chart) TYPE REF TO if_ixml_document .
  PROTECTED SECTION.
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
    DATA type TYPE zexcel_drawing_type VALUE type_image. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA index TYPE string .
    DATA anchor TYPE zexcel_drawing_anchor VALUE anchor_one_cell. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_media_source_www TYPE c VALUE 1.            "#EC NOTEXT
    CONSTANTS c_media_source_xstring TYPE c VALUE 0.        "#EC NOTEXT
    CONSTANTS c_media_source_mime TYPE c VALUE 2.           "#EC NOTEXT
    DATA guid TYPE zexcel_guid .
    DATA media TYPE xstring .
    DATA media_key_www TYPE wwwdatatab .
    DATA media_name TYPE string .
    DATA media_source TYPE c .
    DATA media_type TYPE string .
    DATA io TYPE skwf_io .
    DATA from_loc TYPE zexcel_drawing_location .
    DATA to_loc TYPE zexcel_drawing_location .
    DATA size TYPE zexcel_drawing_size .
ENDCLASS.



CLASS ZCL_EXCEL_DRAWING IMPLEMENTATION.


  METHOD constructor.

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


  METHOD create_media_name.

* if media name is initial, create unique name
    CHECK media_name IS INITIAL.

    index = ip_index.
    CONCATENATE me->type index INTO media_name.
    CONDENSE media_name NO-GAPS.
  ENDMETHOD.


  METHOD emu2pixel.
* suppose 96 DPI
    IF ip_dpi IS SUPPLIED.
      r_pixel = ip_emu * ip_dpi / 914400.
    ELSE.
* suppose 96 DPI
      r_pixel = ip_emu * 96 / 914400.
    ENDIF.
  ENDMETHOD.


  METHOD get_from_col.
    r_from_col = me->from_loc-col.
  ENDMETHOD.


  METHOD get_from_row.
    r_from_row = me->from_loc-row.
  ENDMETHOD.


  METHOD get_guid.

    ep_guid = me->guid.

  ENDMETHOD.


  METHOD get_height_emu_str.
    r_height = pixel2emu( size-height ).
    CONDENSE r_height NO-GAPS.
  ENDMETHOD.


  METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


  METHOD get_media.

    DATA: lv_language LIKE sy-langu.
    DATA: lt_bin_mime TYPE sdokcntbins.
    DATA: lt_mime          TYPE tsfmime,
          lv_filesize      TYPE i,
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


  METHOD get_media_name.
    CONCATENATE media_name  `.` media_type INTO r_name.
  ENDMETHOD.


  METHOD get_media_type.
    r_type = media_type.
  ENDMETHOD.


  METHOD get_name.
    r_name = title.
  ENDMETHOD.


  METHOD get_position.
    rp_position-anchor = anchor.
    rp_position-from = from_loc.
    rp_position-to = to_loc.
    rp_position-size = size.
  ENDMETHOD.


  METHOD get_to_col.
    r_to_col = me->to_loc-col.
  ENDMETHOD.


  METHOD get_to_row.
    r_to_row = me->to_loc-row.
  ENDMETHOD.


  METHOD get_type.
    rp_type = me->type.
  ENDMETHOD.


  METHOD get_width_emu_str.
    r_width = pixel2emu( size-width ).
    CONDENSE r_width NO-GAPS.
  ENDMETHOD.


  METHOD load_chart_attributes.
    DATA: node                TYPE REF TO if_ixml_element.
    DATA: node2               TYPE REF TO if_ixml_element.
    DATA: node3               TYPE REF TO if_ixml_element.
    DATA: node4               TYPE REF TO if_ixml_element.

    DATA lo_barchart TYPE REF TO zcl_excel_graph_bars.
    DATA lo_piechart TYPE REF TO zcl_excel_graph_pie.
    DATA lo_linechart TYPE REF TO zcl_excel_graph_line.

    TYPES: BEGIN OF t_prop,
             val          TYPE string,
             rtl          TYPE string,
             lang         TYPE string,
             formatcode   TYPE string,
             sourcelinked TYPE string,
           END OF t_prop.

    TYPES: BEGIN OF t_pagemargins,
             b      TYPE string,
             l      TYPE string,
             r      TYPE string,
             t      TYPE string,
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
    node2 ?= node->find_from_name_ns( name = 'date1904' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_1904val = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'lang' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_langval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'roundedCorners' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_roundedcornersval = ls_prop-val.

    "style
    node2 ?= node->find_from_name_ns( name = 'style' uri = namespace-c14 ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_c14styleval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'style' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_styleval = ls_prop-val.
    "---------------------------Read graph properties
    "ADDED
    CLEAR node2.
    node2 ?= node->find_from_name_ns( name = 'title' uri = namespace-c ).
    IF node2 IS BOUND AND node2 IS NOT INITIAL.
      node3 ?= node2->find_from_name_ns( name = 't' uri = namespace-a ).
      me->graph->title = node3->get_value( ).
    ENDIF.
    "END

    node2 ?= node->find_from_name_ns( name = 'autoTitleDeleted' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_autotitledeletedval = ls_prop-val.

    "plotArea
    CASE me->graph_type.
      WHEN c_graph_bars.
        node2 ?= node->find_from_name_ns( name = 'barDir' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_bardirval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'grouping' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_groupingval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'invertIfNegative' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_invertifnegative = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
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
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_showbubblesizeval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'gapWidth' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_barchart->ns_gapwidthval = ls_prop-val.

        "Load axes
        node2 ?= node->find_from_name_ns( name = 'barChart' uri = namespace-c ).
        CALL METHOD node2->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'axId'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
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
            node2 ?= node->find_from_name_ns( name = 'catAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'auto' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_auto = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblAlgn' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lblalgn = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblOffset' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lbloffset = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'noMultiLvlLbl' uri = namespace-c ).
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
            node2 ?= node->find_from_name_ns( name = 'valAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossBetween' uri = namespace-c ).
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
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
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
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showbubblesizeval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showLeaderLines' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_showleaderlinesval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'firstSliceAng' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_firstsliceangval = ls_prop-val.
      WHEN c_graph_line.
        node2 ?= node->find_from_name_ns( name = 'grouping' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_groupingval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'varyColors' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_varycolorsval = ls_prop-val.

        "Load series
        CALL METHOD node->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'ser'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
        CALL METHOD lo_collection->create_iterator
          RECEIVING
            rval = lo_iterator.
        lo_node = lo_iterator->get_next( ).
        IF lo_node IS BOUND.
          node2 ?= lo_node->query_interface( ixml_iid_element ).
        ENDIF.
        WHILE lo_node IS BOUND.
          node3 ?= node2->find_from_name_ns( name = 'idx' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_idx = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'order' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_order = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'symbol' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_symbol = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'smooth' uri = namespace-c ).
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
          lv_smooth = ls_prop-val.
          node3 ?= node2->find_from_name_ns( name = 'v' uri = namespace-c ).
          IF node3 IS BOUND.
            lv_sername = node3->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'strRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
            lv_label = node4->get_value( ).
          ENDIF.
          node3 ?= node2->find_from_name_ns( name = 'numRef' uri = namespace-c ).
          IF node3 IS BOUND.
            node4 ?= node3->find_from_name_ns( name = 'f' uri = namespace-c ).
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
        node2 ?= node->find_from_name_ns( name = 'showLegendKey' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showlegendkeyval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showVal' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showvalval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showCatName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showcatnameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showSerName' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showsernameval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showPercent' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showpercentval = ls_prop-val.
        node2 ?= node->find_from_name_ns( name = 'showBubbleSize' uri = namespace-c ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_showbubblesizeval = ls_prop-val.

        node ?= node->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
        node2 ?= node->find_from_name_ns( name = 'marker' uri = namespace-c depth = '1' ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_markerval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'smooth' uri = namespace-c depth = '1' ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_linechart->ns_smoothval = ls_prop-val.
        node ?= ip_chart->if_ixml_node~get_first_child( ).
        CHECK node IS NOT INITIAL.

        "Load axes
        node2 ?= node->find_from_name_ns( name = 'lineChart' uri = namespace-c ).
        CALL METHOD node2->get_elements_by_tag_name_ns
          EXPORTING
*           depth     = 0
            name = 'axId'
            uri  = namespace-c
          RECEIVING
            rval = lo_collection.
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
            node2 ?= node->find_from_name_ns( name = 'catAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'auto' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_auto = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblAlgn' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lblalgn = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'lblOffset' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_lbloffset = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'noMultiLvlLbl' uri = namespace-c ).
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
            node2 ?= node->find_from_name_ns( name = 'valAx' uri = namespace-c ).
            node3 ?= node2->find_from_name_ns( name = 'orientation' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_orientation = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'delete' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_delete = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'axPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_axpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'numFmt' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_formatcode = ls_prop-formatcode.
            lv_sourcelinked = ls_prop-sourcelinked.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_majortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'majorTickMark' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_minortickmark = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'tickLblPos' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_ticklblpos = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossAx' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crossax = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crosses' uri = namespace-c ).
            zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node3 CHANGING cp_structure = ls_prop ).
            lv_crosses = ls_prop-val.
            node3 ?= node2->find_from_name_ns( name = 'crossBetween' uri = namespace-c ).
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
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_barchart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_barchart->ns_overlayval = ls_prop-val.
        ENDIF.
      WHEN c_graph_line.
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_linechart->ns_overlayval = ls_prop-val.
        ENDIF.
      WHEN c_graph_pie.
        node2 ?= node->find_from_name_ns( name = 'legendPos' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_piechart->ns_legendposval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'overlay' uri = namespace-c ).
        IF node2 IS BOUND.
          zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
          lo_piechart->ns_overlayval = ls_prop-val.
        ENDIF.
        node2 ?= node->find_from_name_ns( name = 'pPr' uri = namespace-a ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_pprrtl = ls_prop-rtl.
        node2 ?= node->find_from_name_ns( name = 'endParaRPr' uri = namespace-a ).
        zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
        lo_piechart->ns_endpararprlang = ls_prop-lang.

      WHEN OTHERS.
    ENDCASE.

    node2 ?= node->find_from_name_ns( name = 'plotVisOnly' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_plotvisonlyval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'dispBlanksAs' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_dispblanksasval = ls_prop-val.
    node2 ?= node->find_from_name_ns( name = 'showDLblsOverMax' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_prop ).
    me->graph->ns_showdlblsovermaxval = ls_prop-val.
    "---------------------

    node2 ?= node->find_from_name_ns( name = 'pageMargins' uri = namespace-c ).
    zcl_excel_reader_2007=>fill_struct_from_attributes( EXPORTING ip_element = node2 CHANGING cp_structure = ls_pagemargins ).
    me->graph->pagemargins = ls_pagemargins.


  ENDMETHOD.


  METHOD pixel2emu.
* suppose 96 DPI
    IF ip_dpi IS SUPPLIED.
      r_emu = ip_pixel  * 914400 / ip_dpi.
    ELSE.
* suppose 96 DPI
      r_emu = ip_pixel  * 914400 / 96.
    ENDIF.
  ENDMETHOD.


  METHOD set_media.
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
  ENDMETHOD.


  METHOD set_media_mime.

    DATA: lv_language LIKE sy-langu.

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


  METHOD set_media_www.
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
  ENDMETHOD.


  METHOD set_position.
    from_loc-col = zcl_excel_common=>convert_column2int( ip_from_col ) - 1.
    IF ip_coloff IS SUPPLIED.
      from_loc-col_offset = ip_coloff.
    ENDIF.
    from_loc-row = ip_from_row - 1.
    IF ip_rowoff IS SUPPLIED.
      from_loc-row_offset = ip_rowoff.
    ENDIF.
    anchor = anchor_one_cell.
  ENDMETHOD.


  METHOD set_position2.

    DATA: lv_anchor                     TYPE zexcel_drawing_anchor.
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

  ENDMETHOD.
ENDCLASS.
