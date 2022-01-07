CLASS zcl_excel_ole DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    TYPES ty_doc_url TYPE c LENGTH 255.

    CLASS-METHODS bind_alv_ole2
      IMPORTING
        !i_document_url      TYPE ty_doc_url DEFAULT space
        !i_xls               TYPE c DEFAULT space
        !i_save_path         TYPE string
        !io_alv              TYPE REF TO cl_gui_alv_grid
        !it_listheader       TYPE slis_t_listheader OPTIONAL
        !i_top               TYPE i DEFAULT 1
        !i_left              TYPE i DEFAULT 1
        !i_columns_header    TYPE c DEFAULT 'X'
        !i_columns_autofit   TYPE c DEFAULT 'X'
        !i_format_col_header TYPE soi_format_item OPTIONAL
        !i_format_subtotal   TYPE soi_format_item OPTIONAL
        !i_format_total      TYPE soi_format_item OPTIONAL
      EXCEPTIONS
        miss_guide
        ex_transfer_kkblo_error
        fatal_error
        inv_data_range
        dim_mismatch_vkey
        dim_mismatch_sema
        error_in_sema .

  PROTECTED SECTION.
  PRIVATE SECTION.

    CLASS-METHODS close_document.
    CLASS-METHODS error_doi.

    CLASS-DATA: lo_spreadsheet TYPE REF TO i_oi_spreadsheet,
                lo_control     TYPE REF TO i_oi_container_control,
                lo_proxy       TYPE REF TO i_oi_document_proxy,
                lo_error       TYPE REF TO i_oi_error,
                lc_retcode     TYPE        soi_ret_string.

ENDCLASS.



CLASS zcl_excel_ole IMPLEMENTATION.

  METHOD bind_alv_ole2.
*--------------------------------------------------------------------*
* Method description:
*   Method use to export a CL_GUI_ALV_GRID object to xlsx/xls file
*   with list header and  characteristics of ALV field catalog such as:
*     + Total, group's subtotal
*     + Quantity fields, amount fields (dependent fields)
*     + No_out, no_zero, ...
* Technique use in method:
*   SAP Desktop Office Integration (DOI)
*--------------------------------------------------------------------*

* Data for session 0: DOI constructor
* ------------------------------------------

    DATA: li_has      TYPE i. "Proxy has spreadsheet interface?

* Data for session 1: Get LVC data from ALV object
* ------------------------------------------

    DATA: l_has_activex,
          l_doctype_excel_sheet(11) TYPE c.

* LVC
    DATA: lt_fieldcat_lvc       TYPE lvc_t_fcat.
    DATA: wa_fieldcat_lvc       TYPE lvc_s_fcat.
    DATA: lt_sort_lvc           TYPE lvc_t_sort.
    DATA: lt_filter_idx_lvc     TYPE lvc_t_fidx.
    DATA: lt_grouplevels_lvc    TYPE lvc_t_grpl.

* KKBLO
    DATA: lt_fieldcat_kkblo     TYPE  kkblo_t_fieldcat.
    DATA: lt_sort_kkblo         TYPE  kkblo_t_sortinfo.
    DATA: lt_grouplevels_kkblo  TYPE  kkblo_t_grouplevels.
    DATA: lt_filter_idx_kkblo   TYPE  kkblo_t_sfinfo.
    DATA: wa_listheader         LIKE LINE OF it_listheader.

* Subtotal
    DATA: lt_collect00          TYPE REF TO data.
    DATA: lt_collect01          TYPE REF TO data.
    DATA: lt_collect02          TYPE REF TO data.
    DATA: lt_collect03          TYPE REF TO data.
    DATA: lt_collect04          TYPE REF TO data.
    DATA: lt_collect05          TYPE REF TO data.
    DATA: lt_collect06          TYPE REF TO data.
    DATA: lt_collect07          TYPE REF TO data.
    DATA: lt_collect08          TYPE REF TO data.
    DATA: lt_collect09          TYPE REF TO data.

* data table name
    DATA: l_tabname             TYPE  kkblo_tabname.

* local object
    DATA: lo_grid               TYPE REF TO lcl_gui_alv_grid.

* data table get from ALV
    DATA: lt_alv                  TYPE REF TO data.

* total / subtotal data
    FIELD-SYMBOLS: <f_collect00>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect01>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect02>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect03>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect04>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect05>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect06>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect07>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect08>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect09>  TYPE STANDARD TABLE.

* table before append subtotal lines
    FIELD-SYMBOLS: <f_alv_tab>    TYPE STANDARD TABLE.

* data for session 2: sort, filter and calculate total/subtotal
* ------------------------------------------

* table to save index of subotal / total line in excel tanle
* this ideal to control index of subtotal / total line later
* for ex, when get subtotal / total line to format
    TYPES: BEGIN OF st_subtot_indexs,
             index TYPE i,
           END OF st_subtot_indexs.
    DATA: lt_subtot_indexs TYPE TABLE OF st_subtot_indexs.
    DATA: wa_subtot_indexs LIKE LINE OF lt_subtot_indexs.

* data table after append subtotal
    DATA: lt_excel                TYPE REF TO data.

    DATA: l_tabix                 TYPE i.
    DATA: l_save_index            TYPE i.

* dyn subtotal table name
    DATA: l_collect               TYPE string.

* subtotal range, to format subtotal (and total)
    DATA: subranges               TYPE soi_range_list.
    DATA: subrangeitem            TYPE soi_range_item.
    DATA: l_sub_index             TYPE i.


* table after append subtotal lines
    FIELD-SYMBOLS: <f_excel_tab>  TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_excel_line> TYPE any.

* dyn subtotal tables
    FIELD-SYMBOLS: <f_collect_tab>      TYPE STANDARD TABLE.
    FIELD-SYMBOLS: <f_collect_line>     TYPE any.

    FIELD-SYMBOLS: <f_filter_idx_line>  LIKE LINE OF lt_filter_idx_kkblo.
    FIELD-SYMBOLS: <f_fieldcat_line>    LIKE LINE OF lt_fieldcat_kkblo.
    FIELD-SYMBOLS: <f_grouplevels_line> LIKE LINE OF lt_grouplevels_kkblo.
    FIELD-SYMBOLS: <f_line>             TYPE any.

* Data for session 3: map data to semantic table
* ------------------------------------------

    TYPES: BEGIN OF st_column_index,
             fieldname TYPE kkblo_fieldname,
             tabname   TYPE kkblo_tabname,
             col       LIKE sy-index,
           END OF st_column_index.

* columns index
    DATA: lt_column_index   TYPE TABLE OF st_column_index.
    DATA: wa_column_index   LIKE LINE OF lt_column_index.

* table of dependent field ( currency and quantity unit field)
    DATA: lt_fieldcat_depf  TYPE kkblo_t_fieldcat.
    DATA: wa_fieldcat_depf  TYPE kkblo_fieldcat.

* XXL interface:
* -XXL: contain exporting columns characteristic
    DATA: lt_sema TYPE TABLE OF gxxlt_s INITIAL SIZE 0.
    DATA: wa_sema LIKE LINE OF lt_sema.

* -XXL interface: header
    DATA: lt_hkey TYPE TABLE OF gxxlt_h INITIAL SIZE 0.
    DATA: wa_hkey LIKE LINE OF lt_hkey.

* -XXL interface: header keys
    DATA: lt_vkey TYPE TABLE OF gxxlt_v INITIAL SIZE 0.
    DATA: wa_vkey LIKE LINE OF lt_vkey.

* Number of H Keys: number of key columns
    DATA: l_n_hrz_keys      TYPE  i.
* Number of data columns in the list object: non-key columns no
    DATA: l_n_att_cols      TYPE  i.
* Number of V Keys: number of header row
    DATA: l_n_vrt_keys      TYPE  i.

* curency to format amount
    DATA: lt_tcurx          TYPE TABLE OF tcurx.
    DATA: wa_tcurx          LIKE LINE OF lt_tcurx.
    DATA: l_def             TYPE flag. " currency / quantity flag
    DATA: wa_t006           TYPE t006. " decimal place of unit

    DATA: l_num             TYPE i. " table columns number
    DATA: l_typ             TYPE c. " table type
    DATA: wa                TYPE REF TO data.
    DATA: l_int             TYPE i.
    DATA: l_counter         TYPE i.

    FIELD-SYMBOLS: <f_excel_column>     TYPE any.
    FIELD-SYMBOLS: <f_fcat_column>      TYPE any.

* Data for session 4: write to excel
* ------------------------------------------

    DATA: sema_type         TYPE  c.

    DATA l_error           TYPE REF TO c_oi_proxy_error.
    DATA count              TYPE i.
    DATA datac              TYPE i.
    DATA datareal TYPE i. " exporting column number
    DATA vkeycount          TYPE i.
    DATA all TYPE i.
    DATA mit TYPE i         VALUE 1.  " index of recent row?
    DATA li_col_pos TYPE i  VALUE 1.  " column position
    DATA li_col_num TYPE i.           " table columns number
    FIELD-SYMBOLS: <line>   TYPE any.
    FIELD-SYMBOLS: <item>   TYPE any.

    DATA td                 TYPE sydes_desc.

    DATA: typ.
    DATA: ranges             TYPE soi_range_list.
    DATA: rangeitem          TYPE soi_range_item.
    DATA: contents           TYPE soi_generic_table.
    DATA: contentsitem       TYPE soi_generic_item.
    DATA: semaitem           TYPE gxxlt_s.
    DATA: hkeyitem           TYPE gxxlt_h.
    DATA: vkeyitem           TYPE gxxlt_v.
    DATA: li_commentary_rows TYPE i.  "row number of title lines + 1
    DATA: lo_error_w         TYPE REF TO  i_oi_error.
    DATA: l_retcode          TYPE soi_ret_string.
    DATA: no_flush           TYPE c VALUE 'X'.
    DATA: li_head_top        TYPE i. "header rows position

* Data for session 5: Save and clode document
* ------------------------------------------

    DATA: li_document_size   TYPE i.
    DATA: ls_path            TYPE rlgrap-filename.

*--------------------------------------------------------------------*
* SESSION 0: DOI CONSTRUCTOR
*--------------------------------------------------------------------*

* check active windown

    CALL FUNCTION 'GUI_HAS_ACTIVEX'
      IMPORTING
        return = l_has_activex.

    IF l_has_activex IS INITIAL.
      RAISE miss_guide.
    ENDIF.

*   Get Container Object of Screen

    CALL METHOD c_oi_container_control_creator=>get_container_control
      IMPORTING
        control = lo_control
        retcode = lc_retcode.

    error_doi( ).

* Initialize Container control

    CALL METHOD lo_control->init_control
      EXPORTING
        parent                   = cl_gui_container=>default_screen
        r3_application_name      = ''
        inplace_enabled          = 'X'
        no_flush                 = 'X'
        register_on_close_event  = 'X'
        register_on_custom_event = 'X'
      IMPORTING
        error                    = lo_error
        retcode                  = lc_retcode.

    error_doi( ).

* Get Proxy Document:
* check exist of document proxy, if exist -> close first

    IF NOT lo_proxy IS INITIAL.
      close_document( ).
    ENDIF.

    IF i_xls IS NOT INITIAL.
* xls format, doctype = soi_doctype_excel97_sheet
      l_doctype_excel_sheet = 'Excel.Sheet.8'.
    ELSE.
* xlsx format, doctype = soi_doctype_excel_sheet
      l_doctype_excel_sheet = 'Excel.Sheet'.
    ENDIF.

    CALL METHOD lo_control->get_document_proxy
      EXPORTING
        document_type      = l_doctype_excel_sheet
        register_container = 'X'
      IMPORTING
        document_proxy     = lo_proxy
        error              = lo_error
        retcode            = lc_retcode.

    error_doi( ).

    IF i_document_url IS INITIAL.

* create new excel document

      CALL METHOD lo_proxy->create_document
        EXPORTING
          create_view_data = 'X'
          open_inplace     = 'X'
          no_flush         = 'X'
        IMPORTING
          error            = lo_error
          retcode          = lc_retcode.

      error_doi( ).

    ELSE.

* Read excel template for i_DOCUMENT_URL
* this excel template can be store in local or server

      CALL METHOD lo_proxy->open_document
        EXPORTING
          document_url = i_document_url
          open_inplace = 'X'
          no_flush     = 'X'
        IMPORTING
          error        = lo_error
          retcode      = lc_retcode.

      error_doi( ).

    ENDIF.

* Check Spreadsheet Interface of Document Proxy

    CALL METHOD lo_proxy->has_spreadsheet_interface
      IMPORTING
        is_available = li_has
        error        = lo_error
        retcode      = lc_retcode.

    error_doi( ).

* create Spreadsheet object

    CHECK li_has IS NOT INITIAL.

    CALL METHOD lo_proxy->get_spreadsheet_interface
      IMPORTING
        sheet_interface = lo_spreadsheet
        error           = lo_error
        retcode         = lc_retcode.

    error_doi( ).

*--------------------------------------------------------------------*
* SESSION 1: GET LVC DATA FROM ALV OBJECT
*--------------------------------------------------------------------*

* data table

    CREATE OBJECT lo_grid
      EXPORTING
        i_parent = cl_gui_container=>screen0.

    CALL METHOD lo_grid->get_alv_attributes
      EXPORTING
        io_grid  = io_alv
      IMPORTING
        et_table = lt_alv.

    ASSIGN lt_alv->* TO <f_alv_tab>.

* fieldcat

    CALL METHOD io_alv->get_frontend_fieldcatalog
      IMPORTING
        et_fieldcatalog = lt_fieldcat_lvc.

* table name

    LOOP AT lt_fieldcat_lvc INTO wa_fieldcat_lvc
    WHERE NOT tabname IS INITIAL.
      l_tabname = wa_fieldcat_lvc-tabname.
      EXIT.
    ENDLOOP.

    IF sy-subrc NE 0.
      l_tabname = '1'.
    ENDIF.
    CLEAR: wa_fieldcat_lvc.

* sort table

    CALL METHOD io_alv->get_sort_criteria
      IMPORTING
        et_sort = lt_sort_lvc.


* filter index

    CALL METHOD io_alv->get_filtered_entries
      IMPORTING
        et_filtered_entries = lt_filter_idx_lvc.

* group level + subtotal

    CALL METHOD io_alv->get_subtotals
      IMPORTING
        ep_collect00   = lt_collect00
        ep_collect01   = lt_collect01
        ep_collect02   = lt_collect02
        ep_collect03   = lt_collect03
        ep_collect04   = lt_collect04
        ep_collect05   = lt_collect05
        ep_collect06   = lt_collect06
        ep_collect07   = lt_collect07
        ep_collect08   = lt_collect08
        ep_collect09   = lt_collect09
        et_grouplevels = lt_grouplevels_lvc.

    ASSIGN lt_collect00->* TO <f_collect00>.
    ASSIGN lt_collect01->* TO <f_collect01>.
    ASSIGN lt_collect02->* TO <f_collect02>.
    ASSIGN lt_collect03->* TO <f_collect03>.
    ASSIGN lt_collect04->* TO <f_collect04>.
    ASSIGN lt_collect05->* TO <f_collect05>.
    ASSIGN lt_collect06->* TO <f_collect06>.
    ASSIGN lt_collect07->* TO <f_collect07>.
    ASSIGN lt_collect08->* TO <f_collect08>.
    ASSIGN lt_collect09->* TO <f_collect09>.

* transfer to KKBLO struct

    CALL FUNCTION 'LVC_TRANSFER_TO_KKBLO'
      EXPORTING
        it_fieldcat_lvc           = lt_fieldcat_lvc
        it_sort_lvc               = lt_sort_lvc
        it_filter_index_lvc       = lt_filter_idx_lvc
        it_grouplevels_lvc        = lt_grouplevels_lvc
      IMPORTING
        et_fieldcat_kkblo         = lt_fieldcat_kkblo
        et_sort_kkblo             = lt_sort_kkblo
        et_filtered_entries_kkblo = lt_filter_idx_kkblo
        et_grouplevels_kkblo      = lt_grouplevels_kkblo
      TABLES
        it_data                   = <f_alv_tab>
      EXCEPTIONS
        it_data_missing           = 1
        it_fieldcat_lvc_missing   = 2
        OTHERS                    = 3.
    IF sy-subrc <> 0.
      RAISE ex_transfer_kkblo_error.
    ENDIF.

    CLEAR:
      wa_fieldcat_lvc,
      lt_fieldcat_lvc,
      lt_sort_lvc,
      lt_filter_idx_lvc,
      lt_grouplevels_lvc.

    CLEAR:
      lo_grid.


*--------------------------------------------------------------------*
* SESSION 2: SORT, FILTER AND CALCULATE TOTAL / SUBTOTAL
*--------------------------------------------------------------------*

* append subtotal & total line

    CREATE DATA lt_excel LIKE <f_alv_tab>.
    ASSIGN lt_excel->* TO <f_excel_tab>.

    LOOP AT <f_alv_tab> ASSIGNING <f_line>.
      l_save_index = sy-tabix.

* filter base on filter index table

      READ TABLE lt_filter_idx_kkblo ASSIGNING <f_filter_idx_line>
      WITH KEY index = l_save_index
      BINARY SEARCH.
      IF sy-subrc NE 0.
        APPEND <f_line> TO <f_excel_tab>.
      ENDIF.

* append subtotal lines

      READ TABLE lt_grouplevels_kkblo ASSIGNING <f_grouplevels_line>
      WITH KEY index_to = l_save_index
      BINARY SEARCH.
      IF sy-subrc = 0.
        l_tabix = sy-tabix.
        DO.
          IF <f_grouplevels_line>-subtot EQ 'X' AND
             <f_grouplevels_line>-hide_level IS INITIAL AND
             <f_grouplevels_line>-cindex_from NE 0.

* dynamic append subtotal line to excel table base on grouplevel table
* ex <f_GROUPLEVELS_line>-level = 1
* then <f_collect_tab> = '<F_COLLECT01>'

            l_collect = <f_grouplevels_line>-level.
            CONDENSE l_collect.
            CONCATENATE '<F_COLLECT0'
                        l_collect '>'
*                      '->*'
                        INTO l_collect.

            ASSIGN (l_collect) TO <f_collect_tab>.

* incase there're more than 1 total line of group, at the same level
* for example: subtotal of multi currency

            LOOP AT <f_collect_tab> ASSIGNING <f_collect_line>.
              IF  sy-tabix BETWEEN <f_grouplevels_line>-cindex_from
                              AND  <f_grouplevels_line>-cindex_to.


                APPEND <f_collect_line> TO <f_excel_tab>.

* save subtotal lines index

                wa_subtot_indexs-index = sy-tabix.
                APPEND wa_subtot_indexs TO lt_subtot_indexs.

* append sub total ranges table for format later

                ADD 1 TO l_sub_index.
                subrangeitem-name     =  l_sub_index.
                CONDENSE subrangeitem-name.
                CONCATENATE 'SUBTOT'
                            subrangeitem-name
                            INTO subrangeitem-name.

                subrangeitem-rows     = wa_subtot_indexs-index.
                subrangeitem-columns  = 1.            " start col
                APPEND subrangeitem TO subranges.
                CLEAR: subrangeitem.

              ENDIF.
            ENDLOOP.
            UNASSIGN: <f_collect_tab>.
            UNASSIGN: <f_collect_line>.
            CLEAR: l_collect.
          ENDIF.

* check next subtotal level of group

          UNASSIGN: <f_grouplevels_line>.
          ADD 1 TO l_tabix.

          READ TABLE lt_grouplevels_kkblo ASSIGNING <f_grouplevels_line>
          INDEX l_tabix.
          IF sy-subrc NE 0
          OR <f_grouplevels_line>-index_to NE l_save_index.
            EXIT.
          ENDIF.

          UNASSIGN:
            <f_collect_tab>,
            <f_collect_line>.

        ENDDO.
      ENDIF.

      CLEAR:
        l_tabix,
        l_save_index.

      UNASSIGN:
        <f_filter_idx_line>,
        <f_grouplevels_line>.

    ENDLOOP.

* free local data

    UNASSIGN:
      <f_line>,
      <f_collect_tab>,
      <f_collect_line>,
      <f_fieldcat_line>.

* append grand total line

    IF <f_collect00> IS ASSIGNED.
      ASSIGN <f_collect00> TO <f_collect_tab>.
      IF <f_collect_tab> IS NOT INITIAL.
        LOOP AT <f_collect_tab> ASSIGNING <f_collect_line>.

          APPEND <f_collect_line> TO <f_excel_tab>.

* save total line index

          wa_subtot_indexs-index = sy-tabix.
          APPEND wa_subtot_indexs TO lt_subtot_indexs.

* append grand total range (to format)

          ADD 1 TO l_sub_index.
          subrangeitem-name     =  l_sub_index.
          CONDENSE subrangeitem-name.
          CONCATENATE 'TOTAL'
                      subrangeitem-name
                      INTO subrangeitem-name.

          subrangeitem-rows     = wa_subtot_indexs-index.
          subrangeitem-columns  = 1.            " start col
          APPEND subrangeitem TO subranges.
        ENDLOOP.
      ENDIF.
    ENDIF.

    CLEAR:
      subrangeitem,
      lt_sort_kkblo,
      <f_collect00>,
      <f_collect01>,
      <f_collect02>,
      <f_collect03>,
      <f_collect04>,
      <f_collect05>,
      <f_collect06>,
      <f_collect07>,
      <f_collect08>,
      <f_collect09>.

    UNASSIGN:
      <f_collect00>,
      <f_collect01>,
      <f_collect02>,
      <f_collect03>,
      <f_collect04>,
      <f_collect05>,
      <f_collect06>,
      <f_collect07>,
      <f_collect08>,
      <f_collect09>,
      <f_collect_tab>,
      <f_collect_line>.

*--------------------------------------------------------------------*
* SESSION 3: MAP DATA TO SEMANTIC TABLE
*--------------------------------------------------------------------*

* get dependent field field: currency and quantity

    CREATE DATA wa LIKE LINE OF <f_excel_tab>.
    ASSIGN wa->* TO <f_excel_line>.

    DESCRIBE FIELD <f_excel_line> TYPE l_typ COMPONENTS l_num.

    DO l_num TIMES.
      l_save_index = sy-index.
      ASSIGN COMPONENT l_save_index OF STRUCTURE <f_excel_line>
      TO <f_excel_column>.
      IF sy-subrc NE 0.
        MESSAGE e801(zabap2xlsx) WITH 'FATAL ERROR' RAISING fatal_error.
      ENDIF.

      LOOP AT lt_fieldcat_kkblo ASSIGNING <f_fieldcat_line>
      WHERE tabname = l_tabname.
        ASSIGN COMPONENT <f_fieldcat_line>-fieldname
        OF STRUCTURE <f_excel_line> TO <f_fcat_column>.

        DESCRIBE DISTANCE BETWEEN <f_excel_column> AND <f_fcat_column>
        INTO l_int IN BYTE MODE.

* append column index
* this columns index is of table, not fieldcat

        IF l_int = 0.
          wa_column_index-fieldname = <f_fieldcat_line>-fieldname.
          wa_column_index-tabname   = <f_fieldcat_line>-tabname.
          wa_column_index-col       = l_save_index.
          APPEND wa_column_index TO lt_column_index.
        ENDIF.

* append dependent fields (currency and quantity unit)

        IF <f_fieldcat_line>-cfieldname IS NOT INITIAL.
          CLEAR wa_fieldcat_depf.
          wa_fieldcat_depf-fieldname = <f_fieldcat_line>-cfieldname.
          wa_fieldcat_depf-tabname   = <f_fieldcat_line>-ctabname.
          COLLECT wa_fieldcat_depf INTO lt_fieldcat_depf.
        ENDIF.

        IF <f_fieldcat_line>-qfieldname IS NOT INITIAL.
          CLEAR wa_fieldcat_depf.
          wa_fieldcat_depf-fieldname = <f_fieldcat_line>-qfieldname.
          wa_fieldcat_depf-tabname   = <f_fieldcat_line>-qtabname.
          COLLECT wa_fieldcat_depf INTO lt_fieldcat_depf.
        ENDIF.

* rewrite field data type

        IF <f_fieldcat_line>-inttype = 'X'
        AND <f_fieldcat_line>-datatype(3) = 'INT'.
          <f_fieldcat_line>-inttype = 'I'.
        ENDIF.

      ENDLOOP.

      CLEAR: l_save_index.
      UNASSIGN: <f_fieldcat_line>.

    ENDDO.

* build semantic tables

    l_n_hrz_keys = 1.

*   Get keyfigures

    LOOP AT lt_fieldcat_kkblo ASSIGNING <f_fieldcat_line>
    WHERE tabname = l_tabname
    AND tech NE 'X'
    AND no_out NE 'X'.

      CLEAR wa_sema.
      CLEAR wa_hkey.

*   Units belong to keyfigures -> display as str

      READ TABLE lt_fieldcat_depf INTO wa_fieldcat_depf WITH KEY
      fieldname = <f_fieldcat_line>-fieldname
      tabname   = <f_fieldcat_line>-tabname.

      IF sy-subrc = 0.
        wa_sema-col_typ = 'STR'.
        wa_sema-col_ops = 'DFT'.

*   Keyfigures

      ELSE.
        CASE <f_fieldcat_line>-datatype.
          WHEN 'QUAN'.
            wa_sema-col_typ = 'N03'.

            IF <f_fieldcat_line>-no_sum NE 'X'.
              wa_sema-col_ops = 'ADD'.
            ELSE.
              wa_sema-col_ops = 'NOP'. " no dependent field
            ENDIF.

          WHEN 'DATS'.
            wa_sema-col_typ = 'DAT'.
            wa_sema-col_ops = 'NOP'.

          WHEN 'CHAR' OR 'UNIT' OR 'CUKY'. " Added fieldformats UNIT and CUKY - dd. 26-10-2012 Wouter Heuvelmans
            wa_sema-col_typ = 'STR'.
            wa_sema-col_ops = 'DFT'.   " dependent field

*   incase numeric, ex '00120' -> display as '12'

          WHEN 'NUMC'.
            wa_sema-col_typ = 'STR'.
            wa_sema-col_ops = 'DFT'.

          WHEN OTHERS.
            wa_sema-col_typ = 'NUM'.

            IF <f_fieldcat_line>-no_sum NE 'X'.
              wa_sema-col_ops = 'ADD'.
            ELSE.
              wa_sema-col_ops = 'NOP'.
            ENDIF.
        ENDCASE.
      ENDIF.

      l_counter = l_counter + 1.
      l_n_att_cols = l_n_att_cols + 1.

      wa_sema-col_no = l_counter.

      READ TABLE lt_column_index INTO wa_column_index WITH KEY
      fieldname = <f_fieldcat_line>-fieldname
      tabname   = <f_fieldcat_line>-tabname.

      IF sy-subrc = 0.
        wa_sema-col_src = wa_column_index-col.
      ELSE.
        RAISE fatal_error.
      ENDIF.

* columns index of ref currency field in table

      IF NOT <f_fieldcat_line>-cfieldname IS INITIAL.
        READ TABLE lt_column_index INTO wa_column_index WITH KEY
        fieldname = <f_fieldcat_line>-cfieldname
        tabname   = <f_fieldcat_line>-ctabname.

        IF sy-subrc = 0.
          wa_sema-col_cur = wa_column_index-col.
        ENDIF.

* quantities fields
* treat as currency when display on excel

      ELSEIF NOT <f_fieldcat_line>-qfieldname IS INITIAL.
        READ TABLE lt_column_index INTO wa_column_index WITH KEY
        fieldname = <f_fieldcat_line>-qfieldname
        tabname   = <f_fieldcat_line>-qtabname.
        IF sy-subrc = 0.
          wa_sema-col_cur = wa_column_index-col.
        ENDIF.

      ENDIF.

*   Treat of fixed currency in the fieldcatalog for column

      DATA: l_num_help(2) TYPE n.

      IF NOT <f_fieldcat_line>-currency IS INITIAL.

        SELECT * FROM tcurx INTO TABLE lt_tcurx.
        SORT lt_tcurx.
        READ TABLE lt_tcurx INTO wa_tcurx
                   WITH KEY currkey = <f_fieldcat_line>-currency.
        IF sy-subrc = 0.
          l_num_help = wa_tcurx-currdec.
          CONCATENATE 'N' l_num_help INTO wa_sema-col_typ.
          wa_sema-col_cur = sy-tabix * ( -1 ).
        ENDIF.

      ENDIF.

      wa_hkey-col_no    = l_n_att_cols.
      wa_hkey-row_no    = l_n_hrz_keys.
      wa_hkey-col_name  = <f_fieldcat_line>-reptext.
      APPEND wa_hkey TO lt_hkey.
      APPEND wa_sema TO lt_sema.

    ENDLOOP.

* free local data

    CLEAR:
      lt_column_index,
      wa_column_index,
      lt_fieldcat_depf,
      wa_fieldcat_depf,
      lt_tcurx,
      wa_tcurx,
      l_num,
      l_typ,
      wa,
      l_int,
      l_counter.

    UNASSIGN:
      <f_fieldcat_line>,
      <f_excel_line>,
      <f_excel_column>,
      <f_fcat_column>.

*--------------------------------------------------------------------*
* SESSION 4: WRITE TO EXCEL
*--------------------------------------------------------------------*

    CLEAR wa_tcurx.
    CLEAR lt_tcurx.

*   if spreadsheet dun have proxy yet

    IF li_has IS INITIAL.
      l_retcode = c_oi_errors=>ret_interface_not_supported.
      CALL METHOD c_oi_errors=>create_error_for_retcode
        EXPORTING
          retcode  = l_retcode
          no_flush = no_flush
        IMPORTING
          error    = lo_error_w.
      RETURN.
    ENDIF.

    CREATE OBJECT l_error
      EXPORTING
        object_name = 'OLE_DOCUMENT_PROXY'
        method_name = 'get_ranges_names'.

    CALL METHOD c_oi_errors=>add_error
      EXPORTING
        error = l_error.


    DESCRIBE TABLE lt_sema LINES datareal.
    DESCRIBE TABLE <f_excel_tab> LINES datac.
    DESCRIBE TABLE lt_vkey LINES vkeycount.

    IF datac = 0.
      RAISE inv_data_range.
    ENDIF.


    IF vkeycount NE l_n_vrt_keys.
      RAISE dim_mismatch_vkey.
    ENDIF.

    all = l_n_vrt_keys + l_n_att_cols.

    IF datareal NE all.
      RAISE dim_mismatch_sema.
    ENDIF.

    DATA: decimal TYPE c.

* get decimal separator format ('.', ',', ...) in Office config

    CALL METHOD lo_proxy->get_application_property
      EXPORTING
        property_name    = 'INTERNATIONAL'
        subproperty_name = 'DECIMAL_SEPARATOR'
      CHANGING
        retvalue         = decimal.

    DATA date_format TYPE usr01-datfm.
    SELECT SINGLE datfm FROM usr01 INTO date_format WHERE bname = sy-uname.

    DATA: comma_elim(4) TYPE c.
    FIELD-SYMBOLS <g> TYPE any.
    DATA search_item(4) VALUE '   #'.

    CONCATENATE ',' decimal '.' decimal INTO comma_elim.

    DATA help TYPE i. " table (with subtotal) line number

    help = datac.

    DATA: rowmax TYPE i VALUE 1.    " header row number
    DATA: columnmax TYPE i VALUE 0. " header columns number

    LOOP AT lt_hkey INTO hkeyitem.
      IF hkeyitem-col_no > columnmax.
        columnmax = hkeyitem-col_no.
      ENDIF.

      IF hkeyitem-row_no > rowmax.
        rowmax = hkeyitem-row_no.
      ENDIF.
    ENDLOOP.

    DATA: hkeycolumns TYPE i. " header columns no

    hkeycolumns = columnmax.

    IF hkeycolumns <   l_n_att_cols.
      hkeycolumns = l_n_att_cols.
    ENDIF.

    columnmax = 0.

    LOOP AT lt_vkey INTO vkeyitem.
      IF vkeyitem-col_no > columnmax.
        columnmax = vkeyitem-col_no.
      ENDIF.
    ENDLOOP.

    DATA overflow TYPE i VALUE 1.
    DATA testname(10) TYPE c.
    DATA temp2 TYPE i.                " 1st item row position in excel
    DATA realmit TYPE i VALUE 1.
    DATA realoverflow TYPE i VALUE 1. " row index in content

    CALL METHOD lo_spreadsheet->screen_update
      EXPORTING
        updating = ''.

    CALL METHOD lo_spreadsheet->load_lib.

    DATA: str(40) TYPE c. " range names of columns range (w/o col header)
    DATA: rows TYPE i.    " row postion of 1st item line in ecxel

* calculate row position of data table

    DESCRIBE TABLE it_listheader LINES li_commentary_rows.

* if grid had title, add 1 empy line between title and table

    IF li_commentary_rows NE 0.
      ADD 1 TO li_commentary_rows.
    ENDIF.

* add top position of block data

    li_commentary_rows = li_commentary_rows + i_top - 1.

* write header (commentary rows)

    DATA: li_commentary_row_index TYPE i VALUE 1.
    DATA: li_content_index TYPE i VALUE 1.
    DATA: ls_index(10) TYPE c.
    DATA  ls_commentary_range(40) TYPE c VALUE 'TITLE'.
    DATA: li_font_bold    TYPE i.
    DATA: li_font_italic  TYPE i.
    DATA: li_font_size    TYPE i.

    LOOP AT it_listheader INTO wa_listheader.
      li_commentary_row_index = i_top + li_content_index - 1.
      ls_index = li_content_index.
      CONDENSE ls_index.
      CONCATENATE ls_commentary_range(5) ls_index
                  INTO ls_commentary_range.
      CONDENSE ls_commentary_range.

* insert title range

      CALL METHOD lo_spreadsheet->insert_range_dim
        EXPORTING
          name     = ls_commentary_range
          top      = li_commentary_row_index
          left     = i_left
          rows     = 1
          columns  = 1
          no_flush = no_flush.

* format range

      CASE wa_listheader-typ.
        WHEN 'H'. "title
          li_font_size    = 16.
          li_font_bold    = 1.
          li_font_italic  = -1.
        WHEN 'S'. "subtile
          li_font_size = -1.
          li_font_bold    = 1.
          li_font_italic  = -1.
        WHEN OTHERS. "'A' comment
          li_font_size = -1.
          li_font_bold    = -1.
          li_font_italic  = 1.
      ENDCASE.

      CALL METHOD lo_spreadsheet->set_font
        EXPORTING
          rangename = ls_commentary_range
          family    = ''
          size      = li_font_size
          bold      = li_font_bold
          italic    = li_font_italic
          align     = 0
          no_flush  = no_flush.

* title: range content

      rangeitem-name = ls_commentary_range.
      rangeitem-columns = 1.
      rangeitem-rows = 1.
      APPEND rangeitem TO ranges.

      contentsitem-row    = li_content_index.
      contentsitem-column = 1.
      CONCATENATE wa_listheader-key
                  wa_listheader-info
                  INTO contentsitem-value
                  SEPARATED BY space.
      CONDENSE contentsitem-value.
      APPEND contentsitem TO contents.

      ADD 1 TO li_content_index.

      CLEAR:
        rangeitem,
        contentsitem,
        ls_index.

    ENDLOOP.

* set range data title

    CALL METHOD lo_spreadsheet->set_ranges_data
      EXPORTING
        ranges   = ranges
        contents = contents
        no_flush = no_flush.

    CLEAR: ranges, contents.

    rows = rowmax + li_commentary_rows + 1.

    all = date_format.
    all = all + 3.

    LOOP AT lt_sema INTO semaitem.
      IF semaitem-col_typ = 'DAT' OR semaitem-col_typ = 'MON' OR
         semaitem-col_typ = 'N00' OR semaitem-col_typ = 'N01' OR
         semaitem-col_typ = 'N02' OR
         semaitem-col_typ = 'N03' OR semaitem-col_typ = 'PCT' OR
         semaitem-col_typ = 'STR' OR semaitem-col_typ = 'NUM'.
        CLEAR str.
        str = semaitem-col_no.
        CONDENSE str.
        CONCATENATE 'DATA' str INTO str.
        mit = semaitem-col_no.
        li_col_pos = semaitem-col_no + i_left - 1.

* range from data1 to data(n), for each columns of table

        CALL METHOD lo_spreadsheet->insert_range_dim
          EXPORTING
            name     = str
            top      = rows
            left     = li_col_pos
            rows     = help
            columns  = 1
            no_flush = no_flush.

        DATA dec TYPE i VALUE -1.
        DATA typeinfo TYPE sydes_typeinfo.
        LOOP AT <f_excel_tab> ASSIGNING <line>.
          ASSIGN COMPONENT semaitem-col_no OF STRUCTURE <line> TO <item>.
          DESCRIBE FIELD <item> INTO td.
          READ TABLE td-types INDEX 1 INTO typeinfo.
          IF typeinfo-type = 'P'.
            dec = typeinfo-decimals.
          ELSEIF typeinfo-type = 'I'.
            dec = 0.
          ENDIF.

          DESCRIBE FIELD <line> TYPE typ COMPONENTS count.
          mit = 1.
          DO count TIMES.
            IF mit = semaitem-col_src.
              ASSIGN COMPONENT sy-index OF STRUCTURE <line> TO <item>.
              DESCRIBE FIELD <item> INTO td.
              READ TABLE td-types INDEX 1 INTO typeinfo.
              IF typeinfo-type = 'P'.
                dec = typeinfo-decimals.
              ENDIF.
              EXIT.
            ENDIF.
            mit = mit + 1.
          ENDDO.
          EXIT.
        ENDLOOP.

* format for each columns of table (w/o columns headers)

        IF semaitem-col_typ = 'DAT'.
          IF semaitem-col_no > vkeycount.
            CALL METHOD lo_spreadsheet->set_format
              EXPORTING
                rangename = str
                currency  = ''
                typ       = all
                no_flush  = no_flush.
          ELSE.
            CALL METHOD lo_spreadsheet->set_format
              EXPORTING
                rangename = str
                currency  = ''
                typ       = 0
                no_flush  = no_flush.
          ENDIF.
        ELSEIF semaitem-col_typ = 'STR'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 0
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'MON'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 10
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'N00'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 0
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'N01'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 1
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'N02'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 2
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'N03'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 3
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'N04'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 1
              decimals  = 4
              no_flush  = no_flush.
        ELSEIF semaitem-col_typ = 'NUM'.
          IF dec EQ -1.
            CALL METHOD lo_spreadsheet->set_format
              EXPORTING
                rangename = str
                currency  = ''
                typ       = 1
                decimals  = 2
                no_flush  = no_flush.
          ELSE.
            CALL METHOD lo_spreadsheet->set_format
              EXPORTING
                rangename = str
                currency  = ''
                typ       = 1
                decimals  = dec
                no_flush  = no_flush.
          ENDIF.
        ELSEIF semaitem-col_typ = 'PCT'.
          CALL METHOD lo_spreadsheet->set_format
            EXPORTING
              rangename = str
              currency  = ''
              typ       = 3
              decimals  = 0
              no_flush  = no_flush.
        ENDIF.

      ENDIF.
    ENDLOOP.

* get item contents for set_range_data method
* get currency cell also

    mit = 1.

    DATA: currcells TYPE soi_cell_table.
    DATA: curritem  TYPE soi_cell_item.

    curritem-rows = 1.
    curritem-columns = 1.
    curritem-front = -1.
    curritem-back = -1.
    curritem-font = ''.
    curritem-size = -1.
    curritem-bold = -1.
    curritem-italic = -1.
    curritem-align = -1.
    curritem-frametyp = -1.
    curritem-framecolor = -1.
    curritem-currency = ''.
    curritem-number = 1.
    curritem-input = -1.

    DATA: const TYPE i.

*   Change for Correction request
*    Initial 10000 lines are missing in Excel Export
*    if there are only 2 columns in exported List object.

    IF datareal GT 2.
      const = 20000 / datareal.
    ELSE.
      const = 20000 / ( datareal + 2 ).
    ENDIF.

    DATA: lines TYPE i.
    DATA: innerlines TYPE i.
    DATA: counter TYPE i.
    DATA: curritem2 LIKE curritem.
    DATA: curritem3 LIKE curritem.
    DATA: length TYPE i.
    DATA: found.

* append content table (for method set_range_content)

    LOOP AT <f_excel_tab> ASSIGNING <line>.

* save line index to compare with lt_subtot_indexs,
* to discover line is a subtotal / totale line or not
* ex use to set 'dun display zero in subtotal / total line'

      l_save_index = sy-tabix.

      DO datareal TIMES.
        READ TABLE lt_sema INTO semaitem WITH KEY col_no = sy-index.
        IF semaitem-col_src NE 0.
          ASSIGN COMPONENT semaitem-col_src
                 OF STRUCTURE <line> TO <item>.
        ELSE.
          ASSIGN COMPONENT sy-index
                 OF STRUCTURE <line> TO <item>.
        ENDIF.

        contentsitem-row = realoverflow.

        IF sy-subrc = 0.
          search_item(3) = semaitem-col_ops.
          SEARCH 'ADD#CNT#MIN#MAX#AVG#NOP#DFT#'
                            FOR search_item.
          IF sy-subrc NE 0.
            RAISE error_in_sema.
          ENDIF.
          search_item(3) = semaitem-col_typ.
          SEARCH 'NUM#N00#N01#N02#N03#N04#PCT#DAT#MON#STR#'
                            FOR search_item.
          IF sy-subrc NE 0.
            RAISE error_in_sema.
          ENDIF.
          contentsitem-column = sy-index.
          IF semaitem-col_typ EQ 'DAT' OR semaitem-col_typ EQ 'MON'.
            IF semaitem-col_no > vkeycount.

              " Hinweis 512418
              " EXCEL bezieht Datumsangaben
              " auf den 31.12.1899, behandelt
              " aber 1900 als ein Schaltjahr
              " d.h. ab 1.3.1900 korrekt
              " 1.3.1900 als Zahl = 61

              DATA: genesis TYPE d VALUE '18991230'.
              DATA: number_of_days TYPE p.
* change for date in char format & sema_type = X
              DATA: temp_date TYPE d.

              IF NOT <item> IS INITIAL AND NOT <item> CO ' ' AND NOT
              <item> CO '0'.
* change for date in char format & sema_type = X starts
                IF sema_type = 'X'.
                  DESCRIBE FIELD <item> TYPE typ.
                  IF typ = 'C'.
                    temp_date = <item>.
                    number_of_days = temp_date - genesis.
                  ELSE.
                    number_of_days = <item> - genesis.
                  ENDIF.
                ELSE.
                  number_of_days = <item> - genesis.
                ENDIF.
* change for date in char format & sema_type = X ends
                IF number_of_days < 61.
                  number_of_days = number_of_days - 1.
                ENDIF.

                SET COUNTRY 'DE'.
                WRITE number_of_days TO contentsitem-value
                NO-GROUPING
                                          LEFT-JUSTIFIED.
                SET COUNTRY space.
                TRANSLATE contentsitem-value USING comma_elim.
              ELSE.
                CLEAR contentsitem-value.
              ENDIF.
            ELSE.
              contentsitem-value = <item>.
            ENDIF.
          ELSEIF semaitem-col_typ EQ 'NUM' OR
                 semaitem-col_typ EQ 'N00' OR
                 semaitem-col_typ EQ 'N01' OR
                 semaitem-col_typ EQ 'N02' OR
                 semaitem-col_typ EQ 'N03' OR
                 semaitem-col_typ EQ 'N04' OR
                 semaitem-col_typ EQ 'PCT'.
            SET COUNTRY 'DE'.
            DESCRIBE FIELD <item> TYPE typ.

            IF semaitem-col_cur IS INITIAL.
              IF typ NE 'F'.
                WRITE <item> TO contentsitem-value NO-GROUPING
                                                   NO-SIGN DECIMALS 14.
              ELSE.
                WRITE <item> TO contentsitem-value NO-GROUPING
                                                   NO-SIGN.
              ENDIF.
            ELSE.
* Treat of fixed curreny for column >>Y9CK007319
              IF semaitem-col_cur < 0.
                semaitem-col_cur = semaitem-col_cur * ( -1 ).
                SELECT * FROM tcurx INTO TABLE lt_tcurx.
                SORT lt_tcurx.
                READ TABLE lt_tcurx INTO
                                    wa_tcurx INDEX semaitem-col_cur.
                IF sy-subrc = 0.
                  IF typ NE 'F'.
                    WRITE <item> TO contentsitem-value NO-GROUPING
                     CURRENCY wa_tcurx-currkey NO-SIGN DECIMALS 14.
                  ELSE.
                    WRITE <item> TO contentsitem-value NO-GROUPING
                     CURRENCY wa_tcurx-currkey NO-SIGN.
                  ENDIF.
                ENDIF.
              ELSE.
                ASSIGN COMPONENT semaitem-col_cur
                     OF STRUCTURE <line> TO <g>.
* mit = index of recent row
                curritem-top  = rowmax + mit + li_commentary_rows.

                li_col_pos =  sy-index + i_left - 1.
                curritem-left = li_col_pos.

* if filed is quantity field (qfieldname ne space)
* or amount field (cfieldname ne space), then format decimal place
* corresponding with config

                CLEAR: l_def.
                READ TABLE lt_fieldcat_kkblo ASSIGNING <f_fieldcat_line>
                WITH KEY  tabname = l_tabname
                          tech    = space
                          no_out  = space
                          col_pos = semaitem-col_no.
                IF sy-subrc = 0.
                  IF <f_fieldcat_line>-cfieldname IS NOT INITIAL.
                    l_def = 'C'.
                  ELSE."if <f_fieldcat_line>-qfieldname is not initial.
                    l_def = 'Q'.
                  ENDIF.
                ENDIF.

* if field is amount field
* exporting of amount field base on currency decimal table: TCURX
                IF l_def = 'C'. "field is amount field
                  SELECT SINGLE * FROM tcurx INTO wa_tcurx
                    WHERE currkey = <g>.
* if amount ref to un-know currency -> default decimal  = 2
                  IF sy-subrc EQ 0.
                    curritem-decimals = wa_tcurx-currdec.
                  ELSE.
                    curritem-decimals = 2.
                  ENDIF.

                  APPEND curritem TO currcells.
                  IF typ NE 'F'.
                    WRITE <item> TO contentsitem-value
                                        CURRENCY <g>
                       NO-SIGN NO-GROUPING.
                  ELSE.
                    WRITE <item> TO contentsitem-value
                       DECIMALS 14      CURRENCY <g>
                       NO-SIGN NO-GROUPING.
                  ENDIF.

* if field is quantity field
* exporting of quantity field base on quantity decimal table: T006

                ELSE."if l_def = 'Q'. " field is quantity field
                  CLEAR: wa_t006.
                  SELECT SINGLE * FROM t006 INTO wa_t006
                    WHERE msehi = <g>.
* if quantity ref to un-know unit-> default decimal  = 2
                  IF sy-subrc EQ 0.
                    curritem-decimals = wa_t006-decan.
                  ELSE.
                    curritem-decimals = 2.
                  ENDIF.
                  APPEND curritem TO currcells.

                  WRITE <item> TO contentsitem-value
                                      UNIT <g>
                     NO-SIGN NO-GROUPING.
                  CONDENSE contentsitem-value.

                ENDIF.

              ENDIF.                                        "Y9CK007319
            ENDIF.
            CONDENSE contentsitem-value.

* add function fieldcat-no zero display

            LOOP AT lt_fieldcat_kkblo ASSIGNING <f_fieldcat_line>
            WHERE tabname = l_tabname
            AND   tech NE 'X'
            AND   no_out NE 'X'.
              IF <f_fieldcat_line>-col_pos = semaitem-col_no.
                IF <f_fieldcat_line>-no_zero = 'X'.
                  IF <item> = '0'.
                    CLEAR: contentsitem-value.
                  ENDIF.

* dun display zero in total/subtotal line too

                ELSE.
                  CLEAR: wa_subtot_indexs.
                  READ TABLE lt_subtot_indexs INTO wa_subtot_indexs
                  WITH KEY index = l_save_index.
                  IF sy-subrc = 0 AND <item> = '0'.
                    CLEAR: contentsitem-value.
                  ENDIF.
                ENDIF.
              ENDIF.
            ENDLOOP.
            UNASSIGN: <f_fieldcat_line>.

            IF <item> LT 0.
              SEARCH contentsitem-value FOR 'E'.
              IF sy-fdpos EQ 0.

* use prefix notation for signed numbers

                TRANSLATE contentsitem-value USING '- '.
                CONDENSE contentsitem-value NO-GAPS.
                CONCATENATE '-' contentsitem-value
                           INTO contentsitem-value.
              ELSE.
                CONCATENATE '-' contentsitem-value
                           INTO contentsitem-value.
              ENDIF.
            ENDIF.
            SET COUNTRY space.
* Hier wird nur die korrekte Kommaseparatierung gemacht, wenn die
* Zeichen einer
* Zahl enthalten sind. Das ist für Timestamps, die auch ":" enthalten.
* Für die
* darf keine Kommaseparierung stattfinden.
* Changing for correction request - Y6BK041073
            IF contentsitem-value CO '0123456789.,-+E '.
              TRANSLATE contentsitem-value USING comma_elim.
            ENDIF.
          ELSE.
            CLEAR contentsitem-value.

* if type is not numeric -> dun display with zero

            WRITE <item> TO contentsitem-value NO-ZERO.

            SHIFT contentsitem-value LEFT DELETING LEADING space.

          ENDIF.
          APPEND contentsitem TO contents.
        ENDIF.
      ENDDO.

      realmit = realmit + 1.
      realoverflow = realoverflow + 1.

      mit = mit + 1.
*   overflow = current row index in content table
      overflow = overflow + 1.
    ENDLOOP.

    UNASSIGN: <f_fieldcat_line>.

* set item range for set_range_data method

    testname = mit / const.
    CONDENSE testname.

    CONCATENATE 'TEST' testname INTO testname.

    realoverflow = realoverflow - 1.
    realmit = realmit - 1.
    help = realoverflow.

    rangeitem-name = testname.
    rangeitem-columns = datareal.
    rangeitem-rows = help.
    APPEND rangeitem TO ranges.

* insert item range dim

    temp2 = rowmax + 1 + li_commentary_rows + realmit - realoverflow.

* items data

    CALL METHOD lo_spreadsheet->insert_range_dim
      EXPORTING
        name     = testname
        top      = temp2
        left     = i_left
        rows     = help
        columns  = datareal
        no_flush = no_flush.

* get columns header contents for set_range_data method
* export columns header only if no columns header option = space

    DATA: rowcount TYPE i.
    DATA: columncount TYPE i.

    IF i_columns_header = 'X'.

* append columns header to contents: hkey

      rowcount = 1.
      DO rowmax TIMES.
        columncount = 1.
        DO hkeycolumns TIMES.
          LOOP AT lt_hkey INTO hkeyitem WHERE col_no = columncount
                                           AND row_no   = rowcount.
          ENDLOOP.
          IF sy-subrc = 0.
            str = hkeyitem-col_name.
            contentsitem-value = hkeyitem-col_name.
          ELSE.
            contentsitem-value = str.
          ENDIF.
          contentsitem-column = columncount.
          contentsitem-row = rowcount.
          APPEND contentsitem TO contents.
          columncount = columncount + 1.
        ENDDO.
        rowcount = rowcount + 1.
      ENDDO.

* incase columns header in multiline

      DATA: rowmaxtemp TYPE i.
      IF rowmax > 1.
        rowmaxtemp = rowmax - 1.
        rowcount = 1.
        DO rowmaxtemp TIMES.
          columncount = 1.
          DO columnmax TIMES.
            contentsitem-column = columncount.
            contentsitem-row    = rowcount.
            contentsitem-value  = ''.
            APPEND contentsitem TO contents.
            columncount = columncount + 1.
          ENDDO.
          rowcount = rowcount + 1.
        ENDDO.
      ENDIF.

* append columns header to contents: vkey

      columncount = 1.
      DO columnmax TIMES.
        LOOP AT lt_vkey INTO vkeyitem WHERE col_no = columncount.
        ENDLOOP.
        contentsitem-value = vkeyitem-col_name.
        contentsitem-row = rowmax.
        contentsitem-column = columncount.
        APPEND contentsitem TO contents.
        columncount = columncount + 1.
      ENDDO.
*--------------------------------------------------------------------*
* set header range for method set_range_data
* insert header keys range dim

      li_head_top = li_commentary_rows + 1.
      li_col_pos = i_left.

* insert range headers

      IF hkeycolumns NE 0.
        rangeitem-name = 'TESTHKEY'.
        rangeitem-rows = rowmax.
        rangeitem-columns = hkeycolumns.
        APPEND rangeitem TO ranges.
        CLEAR: rangeitem.

        CALL METHOD lo_spreadsheet->insert_range_dim
          EXPORTING
            name     = 'TESTHKEY'
            top      = li_head_top
            left     = li_col_pos
            rows     = rowmax
            columns  = hkeycolumns
            no_flush = no_flush.
      ENDIF.
    ENDIF.

* format for columns header + total + subtotal
* ------------------------------------------

    help = rowmax + realmit. " table + header lines

    DATA: lt_format     TYPE soi_format_table.
    DATA: wa_format     LIKE LINE OF lt_format.
    DATA: wa_format_temp LIKE LINE OF lt_format.

    FIELD-SYMBOLS: <f_source> TYPE any.
    FIELD-SYMBOLS: <f_des>    TYPE any.

* columns header format

    wa_format-front       = -1.
    wa_format-back        = 15. "grey
    wa_format-font        = space.
    wa_format-size        = -1.
    wa_format-bold        = 1.
    wa_format-align       = 0.
    wa_format-frametyp    = -1.
    wa_format-framecolor  = -1.

* get column header format from input record
* -> map input format

    IF i_columns_header = 'X'.
      wa_format-name        = 'TESTHKEY'.
      IF i_format_col_header IS NOT INITIAL.
        DESCRIBE FIELD i_format_col_header TYPE l_typ COMPONENTS
        li_col_num.
        DO li_col_num TIMES.
          IF sy-index NE 1. " dun map range name
            ASSIGN COMPONENT sy-index OF STRUCTURE i_format_col_header
            TO <f_source>.
            IF <f_source> IS NOT INITIAL.
              ASSIGN COMPONENT sy-index OF STRUCTURE wa_format TO <f_des>.
              <f_des> = <f_source>.
              UNASSIGN: <f_des>.
            ENDIF.
            UNASSIGN: <f_source>.
          ENDIF.
        ENDDO.

        CLEAR: li_col_num.
      ENDIF.

      APPEND wa_format TO lt_format.
    ENDIF.

* Zusammenfassen der Spalten mit gleicher Nachkommastellenzahl
* collect vertical cells (col)  with the same number of decimal places
* to increase perfomance in currency cell format

    DESCRIBE TABLE currcells LINES lines.
    lines = lines - 1.
    DO lines TIMES.
      DESCRIBE TABLE currcells LINES innerlines.
      innerlines = innerlines - 1.
      SORT currcells BY left top.
      CLEAR found.
      DO innerlines TIMES.
        READ TABLE currcells INDEX sy-index INTO curritem.
        counter = sy-index + 1.
        READ TABLE currcells INDEX counter INTO curritem2.
        IF curritem-left EQ curritem2-left.
          length = curritem-top + curritem-rows.
          IF length EQ curritem2-top AND curritem-decimals EQ curritem2-decimals.
            curritem3 = curritem.
            curritem3-rows = curritem3-rows + curritem2-rows.
            curritem-left = -1.
            MODIFY currcells INDEX sy-index FROM curritem.
            curritem2-left = -1.
            MODIFY currcells INDEX counter FROM curritem2.
            APPEND curritem3 TO currcells.
            found = 'X'.
          ENDIF.
        ENDIF.
      ENDDO.
      IF found IS INITIAL.
        EXIT.
      ENDIF.
      DELETE currcells WHERE left = -1.
    ENDDO.

* Zusammenfassen der Zeilen mit gleicher Nachkommastellenzahl
* collect horizontal cells (row) with the same number of decimal places
* to increase perfomance in currency cell format

    DESCRIBE TABLE currcells LINES lines.
    lines = lines - 1.
    DO lines TIMES.
      DESCRIBE TABLE currcells LINES innerlines.
      innerlines = innerlines - 1.
      SORT currcells BY top left.
      CLEAR found.
      DO innerlines TIMES.
        READ TABLE currcells INDEX sy-index INTO curritem.
        counter = sy-index + 1.
        READ TABLE currcells INDEX counter INTO curritem2.
        IF curritem-top EQ curritem2-top AND curritem-rows EQ
        curritem2-rows.
          length = curritem-left + curritem-columns.
          IF length EQ curritem2-left AND curritem-decimals EQ curritem2-decimals.
            curritem3 = curritem.
            curritem3-columns = curritem3-columns + curritem2-columns.
            curritem-left = -1.
            MODIFY currcells INDEX sy-index FROM curritem.
            curritem2-left = -1.
            MODIFY currcells INDEX counter FROM curritem2.
            APPEND curritem3 TO currcells.
            found = 'X'.
          ENDIF.
        ENDIF.
      ENDDO.
      IF found IS INITIAL.
        EXIT.
      ENDIF.
      DELETE currcells WHERE left = -1.
    ENDDO.
* Ende der Zusammenfassung


* item data: format for currency cell, corresponding with currency

    CALL METHOD lo_spreadsheet->cell_format
      EXPORTING
        cells    = currcells
        no_flush = no_flush.

* item data: write item table content

    CALL METHOD lo_spreadsheet->set_ranges_data
      EXPORTING
        ranges   = ranges
        contents = contents
        no_flush = no_flush.

* whole table range to format all table

    IF i_columns_header = 'X'.
      li_head_top = li_commentary_rows + 1.
    ELSE.
      li_head_top = li_commentary_rows + 2.
      help = help - 1.
    ENDIF.

    CALL METHOD lo_spreadsheet->insert_range_dim
      EXPORTING
        name     = 'WHOLE_TABLE'
        top      = li_head_top
        left     = i_left
        rows     = help
        columns  = datareal
        no_flush = no_flush.

* columns width auto fix
* this parameter = space in case use with exist template

    IF i_columns_autofit = 'X'.
      CALL METHOD lo_spreadsheet->fit_widest
        EXPORTING
          name     = 'WHOLE_TABLE'
          no_flush = no_flush.
    ENDIF.

* frame
* The parameter has 8 bits
*0 Left margin
*1 Top marginT
*2 Bottom margin
*3 Right margin
*4 Horizontal line
*5 Vertical line
*6 Thinness
*7 Thickness
* here 127 = 1111111 6-5-4-3-2-1 mean Thin-ver-hor-right-bot-top-left

* ( final DOI method call, set no_flush = space
* equal to call method CL_GUI_CFW=>FLUSH )

    CALL METHOD lo_spreadsheet->set_frame
      EXPORTING
        rangename = 'WHOLE_TABLE'
        typ       = 127
        color     = 1
        no_flush  = space
      IMPORTING
        error     = lo_error
        retcode   = lc_retcode.

    error_doi( ).

* reformat subtotal / total line after format wholw table

    LOOP AT subranges INTO subrangeitem.
      l_sub_index = subrangeitem-rows + li_commentary_rows + rowmax.

      CALL METHOD lo_spreadsheet->insert_range_dim
        EXPORTING
          name     = subrangeitem-name
          left     = i_left
          top      = l_sub_index
          rows     = 1
          columns  = datareal
          no_flush = no_flush.

      wa_format-name    = subrangeitem-name.

*   default format:
*     - clolor: subtotal = light yellow, subtotal = yellow
*     - frame: box

      IF  subrangeitem-name(3) = 'SUB'.
        wa_format-back = 36. "subtotal line
        wa_format_temp = i_format_subtotal.
      ELSE.
        wa_format-back = 27. "total line
        wa_format_temp = i_format_total.
      ENDIF.
      wa_format-frametyp = 79.
      wa_format-framecolor = 1.
      wa_format-number  = -1.
      wa_format-align   = -1.

*   get subtoal + total format from intput parameter
*   overwrite default format

      IF wa_format_temp IS NOT INITIAL.
        DESCRIBE FIELD wa_format_temp TYPE l_typ COMPONENTS li_col_num.
        DO li_col_num TIMES.
          IF sy-index NE 1. " dun map range name
            ASSIGN COMPONENT sy-index OF STRUCTURE wa_format_temp
            TO <f_source>.
            IF <f_source> IS NOT INITIAL.
              ASSIGN COMPONENT sy-index OF STRUCTURE wa_format TO <f_des>.
              <f_des> = <f_source>.
              UNASSIGN: <f_des>.
            ENDIF.
            UNASSIGN: <f_source>.
          ENDIF.
        ENDDO.

        CLEAR: li_col_num.
      ENDIF.

      APPEND wa_format TO lt_format.
      CLEAR: wa_format-name.
      CLEAR: l_sub_index.
      CLEAR: wa_format_temp.

    ENDLOOP.

    IF lt_format[] IS NOT INITIAL.
      CALL METHOD lo_spreadsheet->set_ranges_format
        EXPORTING
          formattable = lt_format
          no_flush    = no_flush.
      CLEAR lt_format.
    ENDIF.
*--------------------------------------------------------------------*
    CALL METHOD lo_spreadsheet->screen_update
      EXPORTING
        updating = 'X'.

    CALL METHOD c_oi_errors=>flush_errors.

    lo_error_w = l_error.
    lc_retcode = lo_error_w->error_code.

    CLEAR:
      lt_sema,
      wa_sema,
      lt_hkey,
      wa_hkey,
      lt_vkey,
      wa_vkey,
      l_n_hrz_keys,
      l_n_att_cols,
      l_n_vrt_keys,
      count,
      datac,
      datareal,
      vkeycount,
      all,
      mit,
      li_col_pos,
      li_col_num,
      ranges,
      rangeitem,
      contents,
      contentsitem,
      semaitem,
      hkeyitem,
      vkeyitem,
      li_commentary_rows,
      l_retcode,
      li_head_top,
      <f_excel_tab>.

    CLEAR:
       lo_error_w.

    UNASSIGN:
    <line>,
    <item>,
    <f_excel_tab>.

*--------------------------------------------------------------------*
* SESSION 5: SAVE AND CLOSE FILE
*--------------------------------------------------------------------*

* ex of save path: 'FILE://C:\temp\test.xlsx'
    CONCATENATE 'FILE://' i_save_path
                INTO ls_path.

    CALL METHOD lo_proxy->save_document_to_url
      EXPORTING
        no_flush      = 'X'
        url           = ls_path
      IMPORTING
        error         = lo_error
        retcode       = lc_retcode
      CHANGING
        document_size = li_document_size.

    error_doi( ).

* if save successfully -> raise successful message
    MESSAGE i400(zabap2xlsx).

    CLEAR:
      ls_path,
      li_document_size.

    close_document( ).
  ENDMETHOD.                    "BIND_ALV_OLE2


  METHOD close_document.

    DATA: l_is_closed TYPE i.

    CLEAR: l_is_closed.
    IF lo_proxy IS NOT INITIAL.

* check proxy detroyed adi

      CALL METHOD lo_proxy->is_destroyed
        IMPORTING
          ret_value = l_is_closed.

* if dun detroyed yet: close -> release proxy

      IF l_is_closed IS INITIAL.
        CALL METHOD lo_proxy->close_document
          IMPORTING
            error   = lo_error
            retcode = lc_retcode.
      ENDIF.

      CALL METHOD lo_proxy->release_document
        IMPORTING
          error   = lo_error
          retcode = lc_retcode.

    ELSE.
      lc_retcode = c_oi_errors=>ret_document_not_open.
    ENDIF.

* Detroy control container

    IF lo_control IS NOT INITIAL.
      CALL METHOD lo_control->destroy_control.
    ENDIF.

    CLEAR:
      lo_spreadsheet,
      lo_proxy,
      lo_control.

  ENDMETHOD.


  METHOD error_doi.

    IF lc_retcode NE c_oi_errors=>ret_ok.
      close_document( ).
      CALL METHOD lo_error->raise_message
        EXPORTING
          type = 'E'.
      CLEAR: lo_error.
    ENDIF.

  ENDMETHOD.


ENDCLASS.
