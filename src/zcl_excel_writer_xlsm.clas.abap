CLASS zcl_excel_writer_xlsm DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_writer_2007
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
    CONSTANTS c_xl_vbaproject TYPE string VALUE 'xl/vbaProject.bin'. "#EC NOTEXT

    METHODS add_further_data_to_zip
        REDEFINITION .
    METHODS create
        REDEFINITION .
    METHODS create_content_types
        REDEFINITION .
    METHODS create_xl_relationships
        REDEFINITION .
    METHODS create_xl_sheet
        REDEFINITION .
    METHODS create_xl_workbook
        REDEFINITION .
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_WRITER_XLSM
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_writer_xlsm IMPLEMENTATION.


  METHOD add_further_data_to_zip.

    super->add_further_data_to_zip( io_zip = io_zip ).

* Add vbaProject.bin to zip
    io_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

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
          lo_drawings         TYPE REF TO zcl_excel_drawings.

    DATA: lv_content         TYPE xstring,
          lv_active          TYPE flag,
          lv_xl_sheet        TYPE string,
          lv_xl_sheet_rels   TYPE string,
          lv_xl_drawing      TYPE string,
          lv_xl_drawing_rels TYPE string,
          lv_syindex         TYPE string,
          lv_value           TYPE string,
          lv_drawing_index   TYPE i,
          lv_comment_index   TYPE i. " (+) Issue 588

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
    lv_drawing_index = 1.

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      IF lo_active_worksheet->get_guid( ) EQ lo_worksheet->get_guid( ).
        lv_active = abap_true.
      ELSE.
        lv_active = abap_false.
      ENDIF.

      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.
      lv_syindex = sy-index.
      lv_comment_index = sy-index. " (+) Issue 588
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet WITH lv_syindex.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index ). " (+) Issue 588
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet_rels WITH lv_syindex.
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

* Add drawings **********************************
      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_syindex = lv_drawing_index.
        SHIFT lv_syindex RIGHT DELETING TRAILING space.
        SHIFT lv_syindex LEFT DELETING LEADING space.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing WITH lv_syindex.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_syindex.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
        ADD 1 TO lv_drawing_index.
      ENDIF.
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
      lv_value = lo_drawing->get_media_name( ).
      CONCATENATE 'xl/charts/' lv_value INTO lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    ENDWHILE.

**********************************************************************
* STEP 9: Add vbaProject.bin to zip
    lo_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

**********************************************************************
* STEP 12: Create the final zip
    ep_excel = lo_zip->save( ).

  ENDMETHOD.


  METHOD create_content_types.
** Constant node name
    DATA: lc_xml_node_workb_ct    TYPE string VALUE 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
          lc_xml_node_default     TYPE string VALUE 'Default',
          " Node attributes
          lc_xml_attr_partname    TYPE string VALUE 'PartName',
          lc_xml_attr_extension   TYPE string VALUE 'Extension',
          lc_xml_attr_contenttype TYPE string VALUE 'ContentType',
          lc_xml_node_workb_pn    TYPE string VALUE '/xl/workbook.xml',
          lc_xml_node_bin_ext     TYPE string VALUE 'bin',
          lc_xml_node_bin_ct      TYPE string VALUE 'application/vnd.ms-office.vbaProject'.


    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    DATA: lv_subrc       TYPE sysubrc,
          lv_contenttype TYPE string.

**********************************************************************
* STEP 3: Create standard contentType
    ep_content = super->create_content_types( ).

**********************************************************************
* STEP 2: modify XML adding the extension bin definition

    CREATE OBJECT lo_document_xml.
    lv_subrc = lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    " extension node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_default
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_extension
                                  value = lc_xml_node_bin_ext ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                  value = lc_xml_node_bin_ct ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: modify XML changing the contentType of node Override /xl/workbook.xml

    lo_collection = lo_document->get_elements_by_tag_name( 'Override' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lv_contenttype = lo_element->get_attribute_ns( lc_xml_attr_partname ).
      IF lv_contenttype EQ lc_xml_node_workb_pn.
        lo_element->remove_attribute_ns( lc_xml_attr_contenttype ).
        lo_element->set_attribute_ns( name  = lc_xml_attr_contenttype
                                      value = lc_xml_node_workb_ct ).
        EXIT.
      ENDIF.
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.


  METHOD create_xl_relationships.

** Constant node name
    DATA: lc_xml_node_relationship TYPE string VALUE 'Relationship',
          " Node attributes
          lc_xml_attr_id           TYPE string VALUE 'Id',
          lc_xml_attr_type         TYPE string VALUE 'Type',
          lc_xml_attr_target       TYPE string VALUE 'Target',
          " Node id
          lc_xml_node_ridx_id      TYPE string VALUE 'rId#',
          " Node type
          lc_xml_node_rid_vba_tp   TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
          " Node target
          lc_xml_node_rid_vba_tg   TYPE string VALUE 'vbaProject.bin'.

    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    DATA: lv_xml_node_ridx_id TYPE string,
          lv_size             TYPE i,
          lv_subrc            TYPE sysubrc,
          lv_syindex(2)       TYPE c.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_relationships( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    CREATE OBJECT lo_document_xml.
    lv_subrc = lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).


    lv_size = excel->get_worksheets_size( ).

    " Relationship node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship
                                                     parent = lo_document ).
    ADD 4 TO lv_size.
    lv_syindex = lv_size.
    SHIFT lv_syindex RIGHT DELETING TRAILING space.
    SHIFT lv_syindex LEFT DELETING LEADING space.
    lv_xml_node_ridx_id = lc_xml_node_ridx_id.
    REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
    lo_element->set_attribute_ns( name  = lc_xml_attr_id
                                  value = lv_xml_node_ridx_id ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_type
                                  value = lc_xml_node_rid_vba_tp ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_target
                                  value = lc_xml_node_rid_vba_tg ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).

  ENDMETHOD.


  METHOD create_xl_sheet.

** Constant node name
    DATA: lc_xml_attr_codename      TYPE string VALUE 'codeName'.

    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    DATA: lv_subrc      TYPE sysubrc.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_sheet( io_worksheet = io_worksheet
                                         iv_active    = iv_active ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    CREATE OBJECT lo_document_xml.
    lv_subrc = lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'sheetPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = io_worksheet->zif_excel_sheet_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  ENDMETHOD.


  METHOD create_xl_workbook.

** Constant node name
    DATA: lc_xml_attr_codename      TYPE string VALUE 'codeName'.

    DATA: lo_ixml          TYPE REF TO if_ixml,
          lo_document      TYPE REF TO if_ixml_document,
          lo_document_xml  TYPE REF TO cl_xml_document,
          lo_element_root  TYPE REF TO if_ixml_node,
          lo_element       TYPE REF TO if_ixml_element,
          lo_collection    TYPE REF TO if_ixml_node_collection,
          lo_iterator      TYPE REF TO if_ixml_node_iterator,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_ostream       TYPE REF TO if_ixml_ostream,
          lo_renderer      TYPE REF TO if_ixml_renderer.

    DATA: lv_subrc      TYPE sysubrc.

**********************************************************************
* STEP 3: Create standard relationship
    ep_content = super->create_xl_workbook( ).

**********************************************************************
* STEP 2: modify XML adding the vbaProject relation

    CREATE OBJECT lo_document_xml.
    lv_subrc = lo_document_xml->parse_xstring( ep_content ).

    lo_document ?= lo_document_xml->m_document.
    lo_element_root = lo_document->if_ixml_node~get_first_child( ).

    lo_collection = lo_document->get_elements_by_tag_name( 'fileVersion' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

    lo_collection = lo_document->get_elements_by_tag_name( 'workbookPr' ).
    lo_iterator = lo_collection->create_iterator( ).
    lo_element ?= lo_iterator->get_next( ).
    WHILE lo_element IS BOUND.
      lo_element->set_attribute_ns( name  = lc_xml_attr_codename
                                    value = me->excel->zif_excel_book_vba_project~codename_pr ).
      lo_element ?= lo_iterator->get_next( ).
    ENDWHILE.

**********************************************************************
* STEP 3: Create xstring stream
    CLEAR ep_content.
    lo_ixml = cl_ixml=>create( ).
    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = ep_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lo_document ).
    lo_renderer->render( ).
  ENDMETHOD.
ENDCLASS.
