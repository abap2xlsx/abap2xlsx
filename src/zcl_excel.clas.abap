CLASS zcl_excel DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL
*"* do not include other source files here!!!
    INTERFACES zif_excel_book_properties .
    INTERFACES zif_excel_book_protection .
    INTERFACES zif_excel_book_vba_project .

    DATA legacy_palette TYPE REF TO zcl_excel_legacy_palette READ-ONLY .
    DATA security TYPE REF TO zcl_excel_security .
    DATA use_template TYPE abap_bool .
    CONSTANTS version TYPE c LENGTH 10 VALUE '7.15.0'.      "#EC NOTEXT

    METHODS add_new_autofilter
      IMPORTING
        !io_sheet            TYPE REF TO zcl_excel_worksheet
      RETURNING
        VALUE(ro_autofilter) TYPE REF TO zcl_excel_autofilter
      RAISING
        zcx_excel .
    METHODS add_new_comment
      RETURNING
        VALUE(eo_comment) TYPE REF TO zcl_excel_comment .
    METHODS add_new_drawing
      IMPORTING
        !ip_type          TYPE zexcel_drawing_type DEFAULT zcl_excel_drawing=>type_image
        !ip_title         TYPE clike OPTIONAL
      RETURNING
        VALUE(eo_drawing) TYPE REF TO zcl_excel_drawing .
    METHODS add_new_range
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS add_new_style
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style OPTIONAL
        !io_clone_of    TYPE REF TO zcl_excel_style OPTIONAL
          PREFERRED PARAMETER !ip_guid
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS add_new_worksheet
      IMPORTING
        !ip_title           TYPE zexcel_sheet_title OPTIONAL
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS add_static_styles .
    METHODS constructor .
    METHODS delete_worksheet
      IMPORTING
        !io_worksheet TYPE REF TO zcl_excel_worksheet
      RAISING
        zcx_excel .
    METHODS delete_worksheet_by_index
      IMPORTING
        !iv_index TYPE numeric
      RAISING
        zcx_excel .
    METHODS delete_worksheet_by_name
      IMPORTING
        !iv_title TYPE clike
      RAISING
        zcx_excel .
    METHODS get_active_sheet_index
      RETURNING
        VALUE(r_active_worksheet) TYPE zexcel_active_worksheet .
    METHODS get_active_worksheet
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_autofilters_reference
      RETURNING
        VALUE(ro_autofilters) TYPE REF TO zcl_excel_autofilters .
    METHODS get_default_style
      RETURNING
        VALUE(ep_style) TYPE zexcel_cell_style .
    METHODS get_drawings_iterator
      IMPORTING
        !ip_type           TYPE zexcel_drawing_type
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_next_table_id
      RETURNING
        VALUE(ep_id) TYPE i .
    METHODS get_ranges_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_static_cellstyle_guid
      IMPORTING
        !ip_cstyle_complete  TYPE zexcel_s_cstyle_complete
        !ip_cstylex_complete TYPE zexcel_s_cstylex_complete
      RETURNING
        VALUE(ep_guid)       TYPE zexcel_cell_style .
    METHODS get_styles_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_style_index_in_styles
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style
      RETURNING
        VALUE(ep_index) TYPE i
      RAISING
        zcx_excel .
    METHODS get_style_to_guid
      IMPORTING
        !ip_guid               TYPE zexcel_cell_style
      RETURNING
        VALUE(ep_stylemapping) TYPE zexcel_s_stylemapping
      RAISING
        zcx_excel .
    METHODS get_theme
      EXPORTING
        !eo_theme TYPE REF TO zcl_excel_theme .
    METHODS get_worksheets_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_worksheets_name
      RETURNING
        VALUE(ep_name) TYPE zexcel_worksheets_name .
    METHODS get_worksheets_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_worksheet_by_index
      IMPORTING
        !iv_index           TYPE numeric
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS get_worksheet_by_name
      IMPORTING
        !ip_sheet_name      TYPE zexcel_sheet_title
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO zcl_excel_worksheet .
    METHODS set_active_sheet_index
      IMPORTING
        !i_active_worksheet TYPE zexcel_active_worksheet
      RAISING
        zcx_excel .
    METHODS set_active_sheet_index_by_name
      IMPORTING
        !i_worksheet_name TYPE zexcel_worksheets_name .
    METHODS set_default_style
      IMPORTING
        !ip_style TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS set_theme
      IMPORTING
        !io_theme TYPE REF TO zcl_excel_theme .
    METHODS fill_template
      IMPORTING
        !iv_data TYPE REF TO zcl_excel_template_data
      RAISING
        zcx_excel .
  PROTECTED SECTION.

    DATA worksheets TYPE REF TO zcl_excel_worksheets .
  PRIVATE SECTION.

    DATA autofilters TYPE REF TO zcl_excel_autofilters .
    DATA charts TYPE REF TO zcl_excel_drawings .
    DATA default_style TYPE zexcel_cell_style .
*"* private components of class ZCL_EXCEL
*"* do not include other source files here!!!
    DATA drawings TYPE REF TO zcl_excel_drawings .
    DATA ranges TYPE REF TO zcl_excel_ranges .
    DATA styles TYPE REF TO zcl_excel_styles .
    DATA t_stylemapping1 TYPE zexcel_t_stylemapping1 .
    DATA t_stylemapping2 TYPE zexcel_t_stylemapping2 .
    DATA theme TYPE REF TO zcl_excel_theme .
    DATA comments TYPE REF TO zcl_excel_comments .

    METHODS get_style_from_guid
      IMPORTING
        !ip_guid        TYPE zexcel_cell_style
      RETURNING
        VALUE(eo_style) TYPE REF TO zcl_excel_style .
    METHODS stylemapping_dynamic_style
      IMPORTING
        !ip_style        TYPE REF TO zcl_excel_style
      RETURNING
        VALUE(eo_style2) TYPE zexcel_s_stylemapping .
ENDCLASS.



CLASS zcl_excel IMPLEMENTATION.


  METHOD add_new_autofilter.
* Check for autofilter reference: new or overwrite; only one per sheet
    ro_autofilter = autofilters->add( io_sheet ) .
  ENDMETHOD.


  METHOD add_new_comment.
    CREATE OBJECT eo_comment.

    comments->add( eo_comment ).
  ENDMETHOD.


  METHOD add_new_drawing.
* Create default blank worksheet
    CREATE OBJECT eo_drawing
      EXPORTING
        ip_type  = ip_type
        ip_title = ip_title.

    CASE ip_type.
      WHEN 'image'.
        drawings->add( eo_drawing ).
      WHEN 'hd_ft'.
        drawings->add( eo_drawing ).
      WHEN 'chart'.
        charts->add( eo_drawing ).
    ENDCASE.
  ENDMETHOD.


  METHOD add_new_range.
* Create default blank range
    CREATE OBJECT eo_range.
    ranges->add( eo_range ).
  ENDMETHOD.


  METHOD add_new_style.
* Start of deletion # issue 139 - Dateretention of cellstyles
*  CREATE OBJECT eo_style.
*  styles->add( eo_style ).
* End of deletion # issue 139 - Dateretention of cellstyles
* Start of insertion # issue 139 - Dateretention of cellstyles
* Create default style
    CREATE OBJECT eo_style
      EXPORTING
        ip_guid     = ip_guid
        io_clone_of = io_clone_of.
    styles->add( eo_style ).

    DATA: style2 TYPE zexcel_s_stylemapping.
* Copy to new representations
    style2 = stylemapping_dynamic_style( eo_style ).
    INSERT style2 INTO TABLE t_stylemapping1.
    INSERT style2 INTO TABLE t_stylemapping2.
* End of insertion # issue 139 - Dateretention of cellstyles

  ENDMETHOD.


  METHOD add_new_worksheet.

* Create default blank worksheet
    CREATE OBJECT eo_worksheet
      EXPORTING
        ip_excel = me
        ip_title = ip_title.

    worksheets->add( eo_worksheet ).
    worksheets->active_worksheet = worksheets->size( ).
  ENDMETHOD.


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


  METHOD constructor.
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
    CREATE OBJECT comments.
    CREATE OBJECT legacy_palette.
    CREATE OBJECT autofilters.

    me->zif_excel_book_protection~initialize( ).
    me->zif_excel_book_properties~initialize( ).

    TRY.
        me->add_new_worksheet( ).
      CATCH zcx_excel. " suppress syntax check error
        ASSERT 1 = 2.  " some error processing anyway
    ENDTRY.

    me->add_new_style( ). " Standard style
    lo_style = me->add_new_style( ). " Standard style with fill gray125
    lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_pattern_gray125.

  ENDMETHOD.


  METHOD delete_worksheet.

    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          l_size          TYPE i,
          lv_errormessage TYPE string.

    l_size = get_worksheets_size( ).
    IF l_size = 1.  " Only 1 worksheet left --> check whether this is the worksheet to be deleted
      lo_worksheet = me->get_worksheet_by_index( 1 ).
      IF lo_worksheet = io_worksheet.
        lv_errormessage = 'Deleting last remaining worksheet is not allowed'(002).
        zcx_excel=>raise_text( lv_errormessage ).
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
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.
    me->delete_worksheet( lo_worksheet ).

  ENDMETHOD.


  METHOD delete_worksheet_by_name.

    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lv_errormessage TYPE string.

    lo_worksheet = me->get_worksheet_by_name( iv_title ).
    IF lo_worksheet IS NOT BOUND.
      lv_errormessage = 'Worksheet not existing'(001).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.
    me->delete_worksheet( lo_worksheet ).

  ENDMETHOD.


  METHOD fill_template.

    DATA: lo_template_filler TYPE REF TO zcl_excel_fill_template.

    FIELD-SYMBOLS:
      <lv_sheet>     TYPE zexcel_sheet_title,
      <lv_data_line> TYPE zcl_excel_template_data=>ts_template_data_sheet.


    lo_template_filler = zcl_excel_fill_template=>create( me ).

    LOOP AT lo_template_filler->mt_sheet ASSIGNING <lv_sheet>.

      READ TABLE iv_data->mt_data ASSIGNING <lv_data_line> WITH KEY sheet = <lv_sheet>.
      CHECK sy-subrc = 0.
      lo_template_filler->fill_sheet( <lv_data_line> ).

    ENDLOOP.

  ENDMETHOD.


  METHOD get_active_sheet_index.
    r_active_worksheet = me->worksheets->active_worksheet.
  ENDMETHOD.


  METHOD get_active_worksheet.

    eo_worksheet = me->worksheets->get( me->worksheets->active_worksheet ).

  ENDMETHOD.


  METHOD get_autofilters_reference.

    ro_autofilters = autofilters.

  ENDMETHOD.


  METHOD get_default_style.
    ep_style = me->default_style.
  ENDMETHOD.


  METHOD get_drawings_iterator.

    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        eo_iterator = me->drawings->get_iterator( ).
      WHEN zcl_excel_drawing=>type_chart.
        eo_iterator = me->charts->get_iterator( ).
      WHEN OTHERS.
    ENDCASE.

  ENDMETHOD.


  METHOD get_next_table_id.
    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lv_tables_count TYPE i.

    lo_iterator = me->get_worksheets_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).

      lv_tables_count = lo_worksheet->get_tables_size( ).
      ADD lv_tables_count TO ep_id.

    ENDWHILE.

    ADD 1 TO ep_id.

  ENDMETHOD.


  METHOD get_ranges_iterator.

    eo_iterator = me->ranges->get_iterator( ).

  ENDMETHOD.


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
      style-guid = zcl_excel_obsolete_func_wrap=>guid_create( ). " ins issue #379 - replacement for outdated function call
      INSERT style INTO TABLE me->t_stylemapping1.
      INSERT style INTO TABLE me->t_stylemapping2.

    ENDIF.

    ep_guid = style-guid.
  ENDMETHOD.


  METHOD get_styles_iterator.

    eo_iterator = me->styles->get_iterator( ).

  ENDMETHOD.


  METHOD get_style_from_guid.

    DATA: lo_style    TYPE REF TO zcl_excel_style,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    lo_iterator = styles->get_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_style ?= lo_iterator->get_next( ).
      IF lo_style->get_guid( ) = ip_guid.
        eo_style = lo_style.
        RETURN.
      ENDIF.
    ENDWHILE.

  ENDMETHOD.


  METHOD get_style_index_in_styles.
    DATA: index TYPE i.
    DATA: lo_iterator TYPE REF TO zcl_excel_collection_iterator,
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
      zcx_excel=>raise_text( 'Index not found' ).
    ELSE.
      SUBTRACT 1 FROM ep_index.  " In excel list starts with "0"
    ENDIF.
  ENDMETHOD.


  METHOD get_style_to_guid.
    DATA: lo_style TYPE REF TO zcl_excel_style.
    " # issue 139
    READ TABLE me->t_stylemapping2 INTO ep_stylemapping WITH TABLE KEY guid = ip_guid.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( 'GUID not found' ).
    ENDIF.

    IF ep_stylemapping-dynamic_style_guid IS NOT INITIAL.
      lo_style = me->get_style_from_guid( ip_guid ).
      zcl_excel_common=>recursive_class_to_struct( EXPORTING i_source = lo_style
                                                   CHANGING  e_target =  ep_stylemapping-complete_style
                                                             e_targetx = ep_stylemapping-complete_stylex ).
    ENDIF.
  ENDMETHOD.


  METHOD get_theme.
    eo_theme = theme.
  ENDMETHOD.


  METHOD get_worksheets_iterator.

    eo_iterator = me->worksheets->get_iterator( ).

  ENDMETHOD.


  METHOD get_worksheets_name.

    ep_name = me->worksheets->name.

  ENDMETHOD.


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
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    me->worksheets->active_worksheet = i_active_worksheet.

  ENDMETHOD.


  METHOD set_active_sheet_index_by_name.

    DATA: ws_it    TYPE REF TO zcl_excel_collection_iterator,
          ws       TYPE REF TO zcl_excel_worksheet,
          lv_title TYPE zexcel_sheet_title,
          count    TYPE i VALUE 1.

    ws_it = me->worksheets->get_iterator( ).

    WHILE ws_it->has_next( ) = abap_true.
      ws ?= ws_it->get_next( ).
      lv_title = ws->get_title( ).
      IF lv_title = i_worksheet_name.
        me->worksheets->active_worksheet = count.
        EXIT.
      ENDIF.
      count = count + 1.
    ENDWHILE.

  ENDMETHOD.


  METHOD set_default_style.
    me->default_style = ip_style.
  ENDMETHOD.


  METHOD set_theme.
    theme = io_theme.
  ENDMETHOD.


  METHOD stylemapping_dynamic_style.
    " # issue 139
    eo_style2-dynamic_style_guid  = ip_style->get_guid( ).
    eo_style2-guid                = eo_style2-dynamic_style_guid.
    eo_style2-added_to_iterator   = abap_true.

* don't care about attributes here, since this data may change
* dynamically

  ENDMETHOD.


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


  METHOD zif_excel_book_protection~initialize.
    me->zif_excel_book_protection~protected      = zif_excel_book_protection=>c_unprotected.
    me->zif_excel_book_protection~lockrevision   = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockstructure  = zif_excel_book_protection=>c_unlocked.
    me->zif_excel_book_protection~lockwindows    = zif_excel_book_protection=>c_unlocked.
    CLEAR me->zif_excel_book_protection~workbookpassword.
    CLEAR me->zif_excel_book_protection~revisionspassword.
  ENDMETHOD.


  METHOD zif_excel_book_vba_project~set_codename.
    me->zif_excel_book_vba_project~codename = ip_codename.
  ENDMETHOD.


  METHOD zif_excel_book_vba_project~set_codename_pr.
    me->zif_excel_book_vba_project~codename_pr = ip_codename_pr.
  ENDMETHOD.


  METHOD zif_excel_book_vba_project~set_vbaproject.
    me->zif_excel_book_vba_project~vbaproject = ip_vbaproject.
  ENDMETHOD.
ENDCLASS.
