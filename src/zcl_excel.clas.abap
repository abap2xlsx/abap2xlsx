class ZCL_EXCEL definition
  public
  create public .

public section.

*"* public components of class ZCL_EXCEL
*"* do not include other source files here!!!
  interfaces ZIF_EXCEL_BOOK_PROPERTIES .
  interfaces ZIF_EXCEL_BOOK_PROTECTION .
  interfaces ZIF_EXCEL_BOOK_VBA_PROJECT .

  data LEGACY_PALETTE type ref to ZCL_EXCEL_LEGACY_PALETTE read-only .
  data SECURITY type ref to ZCL_EXCEL_SECURITY .
  data USE_TEMPLATE type XFELD .

  methods ADD_NEW_AUTOFILTER
    importing
      !IO_SHEET type ref to ZCL_EXCEL_WORKSHEET
    returning
      value(RO_AUTOFILTER) type ref to ZCL_EXCEL_AUTOFILTER
    raising
      ZCX_EXCEL .
  methods ADD_NEW_DRAWING
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE default ZCL_EXCEL_DRAWING=>TYPE_IMAGE
      !IP_TITLE type CLIKE optional
    returning
      value(EO_DRAWING) type ref to ZCL_EXCEL_DRAWING .
  methods ADD_NEW_RANGE
    returning
      value(EO_RANGE) type ref to ZCL_EXCEL_RANGE .
  methods ADD_NEW_STYLE
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE optional
    returning
      value(EO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods ADD_NEW_WORKSHEET
    importing
      !IP_TITLE type ZEXCEL_SHEET_TITLE optional
    returning
      value(EO_WORKSHEET) type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods ADD_STATIC_STYLES .
  methods CONSTRUCTOR .
  methods DELETE_WORKSHEET
    importing
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET
    raising
      ZCX_EXCEL .
  methods DELETE_WORKSHEET_BY_INDEX
    importing
      !IV_INDEX type NUMERIC .
  methods DELETE_WORKSHEET_BY_NAME
    importing
      !IV_TITLE type CLIKE .
  methods GET_ACTIVE_SHEET_INDEX
    returning
      value(R_ACTIVE_WORKSHEET) type ZEXCEL_ACTIVE_WORKSHEET .
  methods GET_ACTIVE_WORKSHEET
    returning
      value(EO_WORKSHEET) type ref to ZCL_EXCEL_WORKSHEET .
  methods GET_AUTOFILTERS_REFERENCE
    returning
      value(RO_AUTOFILTERS) type ref to ZCL_EXCEL_AUTOFILTERS .
  methods GET_DEFAULT_STYLE
    returning
      value(EP_STYLE) type ZEXCEL_CELL_STYLE .
  methods GET_DRAWINGS_ITERATOR
    importing
      !IP_TYPE type ZEXCEL_DRAWING_TYPE
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_NEXT_TABLE_ID
    returning
      value(EP_ID) type I .
  methods GET_RANGES_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_STATIC_CELLSTYLE_GUID
    importing
      !IP_CSTYLE_COMPLETE type ZEXCEL_S_CSTYLE_COMPLETE
      !IP_CSTYLEX_COMPLETE type ZEXCEL_S_CSTYLEX_COMPLETE
    returning
      value(EP_GUID) type ZEXCEL_CELL_STYLE .
  methods GET_STYLES_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_STYLE_INDEX_IN_STYLES
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE
    returning
      value(EP_INDEX) type SYTABIX
    raising
      ZCX_EXCEL .
  methods GET_STYLE_TO_GUID
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE
    returning
      value(EP_STYLEMAPPING) type ZEXCEL_S_STYLEMAPPING
    raising
      ZCX_EXCEL .
  methods GET_THEME
    exporting
      !EO_THEME type ref to ZCL_EXCEL_THEME .
  methods GET_WORKSHEETS_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods GET_WORKSHEETS_NAME
    returning
      value(EP_NAME) type ZEXCEL_WORKSHEETS_NAME .
  methods GET_WORKSHEETS_SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_WORKSHEET_BY_INDEX
    importing
      !IV_INDEX type NUMERIC
    returning
      value(EO_WORKSHEET) type ref to ZCL_EXCEL_WORKSHEET .
  methods GET_WORKSHEET_BY_NAME
    importing
      !IP_SHEET_NAME type ZEXCEL_SHEET_TITLE
    returning
      value(EO_WORKSHEET) type ref to ZCL_EXCEL_WORKSHEET .
  methods SET_ACTIVE_SHEET_INDEX
    importing
      !I_ACTIVE_WORKSHEET type ZEXCEL_ACTIVE_WORKSHEET
    raising
      ZCX_EXCEL .
  methods SET_ACTIVE_SHEET_INDEX_BY_NAME
    importing
      !I_WORKSHEET_NAME type ZEXCEL_WORKSHEETS_NAME .
  methods SET_DEFAULT_STYLE
    importing
      !IP_STYLE type ZEXCEL_CELL_STYLE
    raising
      ZCX_EXCEL .
  methods SET_THEME
    importing
      !IO_THEME type ref to ZCL_EXCEL_THEME .
protected section.

  data WORKSHEETS type ref to ZCL_EXCEL_WORKSHEETS .
private section.

  constants VERSION type CHAR10 value '7.0.6'. "#EC NOTEXT
  data AUTOFILTERS type ref to ZCL_EXCEL_AUTOFILTERS .
  data CHARTS type ref to ZCL_EXCEL_DRAWINGS .
  data DEFAULT_STYLE type ZEXCEL_CELL_STYLE .
*"* private components of class ZCL_EXCEL
*"* do not include other source files here!!!
  data DRAWINGS type ref to ZCL_EXCEL_DRAWINGS .
  data RANGES type ref to ZCL_EXCEL_RANGES .
  data STYLES type ref to ZCL_EXCEL_STYLES .
  data T_STYLEMAPPING1 type ZEXCEL_T_STYLEMAPPING1 .
  data T_STYLEMAPPING2 type ZEXCEL_T_STYLEMAPPING2 .
  data THEME type ref to ZCL_EXCEL_THEME .

  methods GET_STYLE_FROM_GUID
    importing
      !IP_GUID type ZEXCEL_CELL_STYLE
    returning
      value(EO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods STYLEMAPPING_DYNAMIC_STYLE
    importing
      !IP_STYLE type ref to ZCL_EXCEL_STYLE
    returning
      value(EO_STYLE2) type ZEXCEL_S_STYLEMAPPING .
ENDCLASS.



CLASS ZCL_EXCEL IMPLEMENTATION.


METHOD add_new_autofilter.
* Check for autofilter reference: new or overwrite; only one per sheet
  ro_autofilter = autofilters->add( io_sheet ) .
ENDMETHOD.


method ADD_NEW_DRAWING.
* Create default blank worksheet
  CREATE OBJECT eo_drawing
    EXPORTING
      ip_type  = ip_type
      ip_title = ip_title.

  CASE ip_type.
    WHEN 'image'.
      drawings->add( eo_drawing ).
    WHEN 'chart'.
      charts->add( eo_drawing ).
  ENDCASE.
  endmethod.


method ADD_NEW_RANGE.
* Create default blank range
  CREATE OBJECT eo_range.
  ranges->add( eo_range ).
  endmethod.


method ADD_NEW_STYLE.
* Start of deletion # issue 139 - Dateretention of cellstyles
*  CREATE OBJECT eo_style.
*  styles->add( eo_style ).
* End of deletion # issue 139 - Dateretention of cellstyles
* Start of insertion # issue 139 - Dateretention of cellstyles
* Create default style
  CREATE OBJECT eo_style
    EXPORTING
      ip_guid = ip_guid.
  styles->add( eo_style ).

  DATA: style2 TYPE zexcel_s_stylemapping.
* Copy to new representations
  style2 = stylemapping_dynamic_style( eo_style ).
  INSERT style2 INTO TABLE t_stylemapping1.
  INSERT style2 INTO TABLE t_stylemapping2.
* End of insertion # issue 139 - Dateretention of cellstyles

  endmethod.


method ADD_NEW_WORKSHEET.

* Create default blank worksheet
  CREATE OBJECT eo_worksheet
    EXPORTING
      ip_excel = me
      ip_title = ip_title.

  worksheets->add( eo_worksheet ).
  worksheets->active_worksheet = worksheets->size( ).
  endmethod.


METHOD add_static_styles.
  " # issue 139
  FIELD-SYMBOLS: <style1> LIKE LINE OF t_stylemapping1,
                 <style2> LIKE LINE OF t_stylemapping2.
  DATA: style TYPE REF TO zcl_excel_style.

  LOOP AT me->t_stylemapping1 ASSIGNING <style1> WHERE added_to_iterator IS INITIAL.
    READ TABLE me->t_stylemapping2 ASSIGNING <style2> WITH TABLE KEY guid = <style1>-guid.
    CHECK sy-subrc = 0.  " Should always be true since these tables are being filled parallel

    style = me->add_new_style( <style1>-guid ).

    zcl_excel_common=>recursive_struct_to_class( EXPORTING i_source  = <style1>-complete_style
                                                           i_sourcex = <style1>-complete_stylex
                                                 CHANGING  e_target  = style ).

  ENDLOOP.
ENDMETHOD.


method CONSTRUCTOR.
  DATA: lo_style      TYPE REF TO zcl_excel_style.

* Inizialize instance objects
  CREATE OBJECT security.
  CREATE OBJECT worksheets.
  CREATE OBJECT ranges.
  CREATE OBJECT styles.
  CREATE OBJECT drawings
    EXPORTING
      ip_type = zcl_excel_drawing=>type_image.
  CREATE OBJECT charts
    EXPORTING
      ip_type = zcl_excel_drawing=>type_chart.
  CREATE OBJECT legacy_palette.
  CREATE OBJECT autofilters.

  me->zif_excel_book_protection~initialize( ).
  me->zif_excel_book_properties~initialize( ).

  me->add_new_worksheet( ).
  me->add_new_style( ). " Standard style
  lo_style = me->add_new_style( ). " Standard style with fill gray125
  lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_pattern_gray125.

  endmethod.


METHOD delete_worksheet.

  DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
        l_size          TYPE i,
        lv_errormessage TYPE string.

  l_size = get_worksheets_size( ).
  IF l_size = 1.  " Only 1 worksheet left --> check whether this is the worksheet to be deleted
    lo_worksheet = me->get_worksheet_by_index( 1 ).
    IF lo_worksheet = io_worksheet.
      lv_errormessage = 'Deleting last remaining worksheet is not allowed'(002).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = lv_errormessage.
    ENDIF.
  ENDIF.

  me->worksheets->remove( io_worksheet ).

ENDMETHOD.


METHOD delete_worksheet_by_index.

  DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
        lv_errormessage TYPE string.

  lo_worksheet = me->get_worksheet_by_index( iv_index ).
  IF lo_worksheet IS NOT BOUND.
    lv_errormessage = 'Worksheet not existing'(001).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.
  me->delete_worksheet( lo_worksheet ).

ENDMETHOD.


METHOD delete_worksheet_by_name.

  DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
        lv_errormessage TYPE string.

  lo_worksheet = me->get_worksheet_by_name( iv_title ).
  IF lo_worksheet IS NOT BOUND.
    lv_errormessage = 'Worksheet not existing'(001).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.
  me->delete_worksheet( lo_worksheet ).

ENDMETHOD.


method GET_ACTIVE_SHEET_INDEX.
  r_active_worksheet = me->worksheets->active_worksheet.
  endmethod.


method GET_ACTIVE_WORKSHEET.

  eo_worksheet = me->worksheets->get( me->worksheets->active_worksheet ).

  endmethod.


method GET_AUTOFILTERS_REFERENCE.

  ro_autofilters = autofilters.

  endmethod.


method GET_DEFAULT_STYLE.
  ep_style = me->default_style.
  endmethod.


method GET_DRAWINGS_ITERATOR.

  CASE ip_type.
    WHEN zcl_excel_drawing=>type_image.
      eo_iterator = me->drawings->get_iterator( ).
    WHEN zcl_excel_drawing=>type_chart.
      eo_iterator = me->charts->get_iterator( ).
    WHEN OTHERS.
  ENDCASE.

  endmethod.


method GET_NEXT_TABLE_ID.
  DATA: lo_worksheet      TYPE REF TO zcl_excel_worksheet,
        lo_iterator       TYPE REF TO cl_object_collection_iterator,
        lv_tables_count   TYPE i.

  lo_iterator = me->get_worksheets_iterator( ).
  WHILE lo_iterator->if_object_collection_iterator~has_next( ) EQ abap_true.
    lo_worksheet ?= lo_iterator->if_object_collection_iterator~get_next( ).

    lv_tables_count = lo_worksheet->get_tables_size( ).
    ADD lv_tables_count TO ep_id.

  ENDWHILE.

  ADD 1 TO ep_id.

  endmethod.


method GET_RANGES_ITERATOR.

  eo_iterator = me->ranges->get_iterator( ).

  endmethod.


METHOD get_static_cellstyle_guid.
  " # issue 139
  DATA: style LIKE LINE OF me->t_stylemapping1.

  READ TABLE me->t_stylemapping1 INTO style
    WITH TABLE KEY dynamic_style_guid = style-guid  " no dynamic style  --> look for initial guid here
                   complete_style     = ip_cstyle_complete
                   complete_stylex    = ip_cstylex_complete.
  IF sy-subrc <> 0.
    style-complete_style  = ip_cstyle_complete.
    style-complete_stylex = ip_cstylex_complete.
*    CALL FUNCTION 'GUID_CREATE'                               " del issue #379 - function is outdated in newer releases
*      IMPORTING
*        ev_guid_16 = style-guid.
    style-guid = zcl_excel_obsolete_func_wrap=>guid_create( ). " ins issue #379 - replacement for outdated function call
    INSERT style INTO TABLE me->t_stylemapping1.
    INSERT style INTO TABLE me->t_stylemapping2.

  ENDIF.

  ep_guid = style-guid.
ENDMETHOD.


method GET_STYLES_ITERATOR.

  eo_iterator = me->styles->get_iterator( ).

  endmethod.


METHOD get_style_from_guid.

  DATA: lo_style    TYPE REF TO zcl_excel_style,
        lo_iterator TYPE REF TO cl_object_collection_iterator.

  lo_iterator = styles->get_iterator( ).
  WHILE lo_iterator->has_next( ) = abap_true.
    lo_style ?= lo_iterator->get_next( ).
    IF lo_style->get_guid( ) = ip_guid.
      eo_style = lo_style.
      RETURN.
    ENDIF.
  ENDWHILE.

ENDMETHOD.


method GET_STYLE_INDEX_IN_STYLES.
  DATA: index TYPE syindex.
  DATA: lo_iterator TYPE REF TO cl_object_collection_iterator,
        lo_style    TYPE REF TO zcl_excel_style.

  CHECK ip_guid IS NOT INITIAL.


  lo_iterator = me->get_styles_iterator( ).
  WHILE lo_iterator->has_next( ) = 'X'.
    ADD 1 TO index.
    lo_style ?= lo_iterator->get_next( ).
    IF lo_style->get_guid( ) = ip_guid.
      ep_index = index.
      EXIT.
    ENDIF.
  ENDWHILE.

  IF ep_index IS INITIAL.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'Index not found'.
  else.
    SUBTRACT 1 from ep_index.  " In excel list starts with "0"
  ENDIF.
  endmethod.


METHOD get_style_to_guid.
  DATA: lo_style TYPE REF TO zcl_excel_style.
  " # issue 139
  READ TABLE me->t_stylemapping2 INTO ep_stylemapping WITH TABLE KEY guid = ip_guid.
  IF sy-subrc <> 0.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = 'GUID not found'.
  ENDIF.

  IF ep_stylemapping-dynamic_style_guid IS NOT INITIAL.
    lo_style = me->get_style_from_guid( ip_guid ).
    zcl_excel_common=>recursive_class_to_struct( EXPORTING i_source = lo_style
                                                 CHANGING  e_target =  ep_stylemapping-complete_style
                                                           e_targetx = ep_stylemapping-complete_stylex ).
  ENDIF.
ENDMETHOD.


method GET_THEME.
  eo_theme = theme.
endmethod.


method GET_WORKSHEETS_ITERATOR.

  eo_iterator = me->worksheets->get_iterator( ).

  endmethod.


method GET_WORKSHEETS_NAME.

  ep_name = me->worksheets->name.

  endmethod.


METHOD get_worksheets_size.

  ep_size = me->worksheets->size( ).

ENDMETHOD.


METHOD get_worksheet_by_index.


  DATA: lv_index TYPE zexcel_active_worksheet.

  lv_index = iv_index.
  eo_worksheet = me->worksheets->get( lv_index ).

ENDMETHOD.


METHOD get_worksheet_by_name.

  DATA: lv_index TYPE zexcel_active_worksheet,
        l_size   TYPE i.

  l_size = get_worksheets_size( ).

  DO l_size TIMES.
    lv_index = sy-index.
    eo_worksheet = me->worksheets->get( lv_index ).
    IF eo_worksheet->get_title( ) = ip_sheet_name.
      RETURN.
    ENDIF.
  ENDDO.

  CLEAR eo_worksheet.

ENDMETHOD.


METHOD set_active_sheet_index.
  DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
        lv_errormessage TYPE string.

*--------------------------------------------------------------------*
* Check whether worksheet exists
*--------------------------------------------------------------------*
  lo_worksheet = me->get_worksheet_by_index( i_active_worksheet ).
  IF lo_worksheet IS NOT BOUND.
    lv_errormessage = 'Worksheet not existing'(001).
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = lv_errormessage.
  ENDIF.

  me->worksheets->active_worksheet = i_active_worksheet.

ENDMETHOD.


METHOD set_active_sheet_index_by_name.

  DATA: ws_it    TYPE REF TO cl_object_collection_iterator,
        ws       TYPE REF TO zcl_excel_worksheet,
        lv_title TYPE zexcel_sheet_title,
        count    TYPE i VALUE 1.

  ws_it = me->worksheets->get_iterator( ).

  WHILE ws_it->if_object_collection_iterator~has_next( ) = abap_true.
    ws ?= ws_it->if_object_collection_iterator~get_next( ).
    lv_title = ws->get_title( ).
    IF lv_title = i_worksheet_name.
      me->worksheets->active_worksheet = count.
      EXIT.
    ENDIF.
    count = count + 1.
  ENDWHILE.

ENDMETHOD.


method SET_DEFAULT_STYLE.
  me->default_style = ip_style.
  endmethod.


method SET_THEME.
  theme = io_theme.
endmethod.


method STYLEMAPPING_DYNAMIC_STYLE.
" # issue 139
  eo_style2-dynamic_style_guid  = ip_style->get_guid( ).
  eo_style2-guid                = eo_style2-dynamic_style_guid.
  eo_style2-added_to_iterator   = abap_true.

* don't care about attributes here, since this data may change
* dynamically

  endmethod.


METHOD zif_excel_book_properties~initialize.
  DATA: lv_timestamp TYPE timestampl.

  me->zif_excel_book_properties~application     = 'Microsoft Excel'.
  me->zif_excel_book_properties~appversion      = '12.0000'.

  GET TIME STAMP FIELD lv_timestamp.
  me->zif_excel_book_properties~created         = lv_timestamp.
  me->zif_excel_book_properties~creator         = sy-uname.
  me->zif_excel_book_properties~description     = zcl_excel=>version.
  me->zif_excel_book_properties~modified        = lv_timestamp.
  me->zif_excel_book_properties~lastmodifiedby  = sy-uname.
ENDMETHOD.


method ZIF_EXCEL_BOOK_PROTECTION~INITIALIZE.
  me->zif_excel_book_protection~protected      = zif_excel_book_protection=>c_unprotected.
  me->zif_excel_book_protection~lockrevision   = zif_excel_book_protection=>c_unlocked.
  me->zif_excel_book_protection~lockstructure  = zif_excel_book_protection=>c_unlocked.
  me->zif_excel_book_protection~lockwindows    = zif_excel_book_protection=>c_unlocked.
  CLEAR me->zif_excel_book_protection~workbookpassword.
  CLEAR me->zif_excel_book_protection~revisionspassword.
  endmethod.


method ZIF_EXCEL_BOOK_VBA_PROJECT~SET_CODENAME.
  me->zif_excel_book_vba_project~codename = ip_codename.
  endmethod.


method ZIF_EXCEL_BOOK_VBA_PROJECT~SET_CODENAME_PR.
  me->zif_excel_book_vba_project~codename_pr = ip_codename_pr.
  endmethod.


method ZIF_EXCEL_BOOK_VBA_PROJECT~SET_VBAPROJECT.
  me->zif_excel_book_vba_project~vbaproject = ip_vbaproject.
  endmethod.
ENDCLASS.
