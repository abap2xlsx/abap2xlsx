*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL11
*& Export Organisation and Contact Persons using ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT  zdemo_excel11.

TYPE-POOLS: abap.

DATA: central_search TYPE bapibus1006_central_search,
      addressdata_search TYPE bapibus1006_addr_search,
      others_search TYPE bapibus1006_other_data.
DATA: searchresult TYPE TABLE OF bapibus1006_bp_addr,
      return TYPE TABLE OF bapiret2.
DATA: lines TYPE i.
FIELD-SYMBOLS: <searchresult_line> LIKE LINE OF searchresult.
DATA: centraldata             TYPE bapibus1006_central,
      centraldataperson       TYPE bapibus1006_central_person,
      centraldataorganization TYPE bapibus1006_central_organ.
DATA: addressdata TYPE bapibus1006_address.
DATA: relationships TYPE TABLE OF bapibus1006_relations.
FIELD-SYMBOLS: <relationship> LIKE LINE OF relationships.
DATA: relationship_centraldata TYPE bapibus1006002_central.
DATA: relationship_addresses TYPE TABLE OF bapibus1006002_addresses.
FIELD-SYMBOLS: <relationship_address> LIKE LINE OF relationship_addresses.

DATA: lt_download TYPE TABLE OF zexcel_s_org_rel.
FIELD-SYMBOLS: <download> LIKE LINE OF lt_download.

CONSTANTS: gc_save_file_name TYPE string VALUE '11_Export_Org_and_Contact.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.


PARAMETERS: md TYPE flag RADIOBUTTON GROUP act.

SELECTION-SCREEN BEGIN OF BLOCK a WITH FRAME TITLE text-00a.
PARAMETERS: partnerc TYPE bu_type   DEFAULT 2, " Organizations
            postlcod TYPE ad_pstcd1 DEFAULT '8334*',
            country  TYPE land1     DEFAULT 'DE',
            maxsel   TYPE bu_maxsel DEFAULT 100.
SELECTION-SCREEN END OF BLOCK a.

PARAMETERS: rel TYPE flag RADIOBUTTON GROUP act DEFAULT 'X'.

SELECTION-SCREEN BEGIN OF BLOCK b WITH FRAME TITLE text-00b.
PARAMETERS: reltyp TYPE bu_reltyp DEFAULT 'BUR011',
            partner TYPE bu_partner DEFAULT '191'.
SELECTION-SCREEN END OF BLOCK b.

START-OF-SELECTION.
  IF md = abap_true.
    " Read all Companies by Master Data
    central_search-partnercategory = partnerc.
    addressdata_search-postl_cod1  = postlcod.
    addressdata_search-country     = country.
    others_search-maxsel           = maxsel.
    others_search-no_search_for_contactperson = 'X'.

    CALL FUNCTION 'BAPI_BUPA_SEARCH_2'
      EXPORTING
        centraldata  = central_search
        addressdata  = addressdata_search
        OTHERS       = others_search
      TABLES
        searchresult = searchresult
        return       = return.

    SORT searchresult BY partner.
    DELETE ADJACENT DUPLICATES FROM searchresult COMPARING partner.
  ELSEIF rel = abap_true.
    " Read by Relationship
    SELECT but050~partner1 AS partner FROM but050
      INNER JOIN but000 ON but000~partner = but050~partner1 AND but000~type = '2'
      INTO CORRESPONDING FIELDS OF TABLE searchresult
      WHERE but050~partner2 = partner
        AND but050~reltyp   = reltyp.
  ENDIF.

  DESCRIBE TABLE searchresult LINES lines.
  WRITE: / 'Number of search results: ', lines.

  LOOP AT searchresult ASSIGNING <searchresult_line>.
    " Read Details of Organization
    CALL FUNCTION 'BAPI_BUPA_CENTRAL_GETDETAIL'
      EXPORTING
        businesspartner         = <searchresult_line>-partner
      IMPORTING
        centraldataorganization = centraldataorganization.
    " Read Standard Address of Organization
    CALL FUNCTION 'BAPI_BUPA_ADDRESS_GETDETAIL'
      EXPORTING
        businesspartner = <searchresult_line>-partner
      IMPORTING
        addressdata     = addressdata.

    " Add Organization to Download
    APPEND INITIAL LINE TO lt_download ASSIGNING <download>.
    " Fill Organization Partner Numbers
    CALL FUNCTION 'BAPI_BUPA_GET_NUMBERS'
      EXPORTING
        businesspartner        = <searchresult_line>-partner
      IMPORTING
        businesspartnerout     = <download>-org_number
        businesspartnerguidout = <download>-org_guid.

    MOVE-CORRESPONDING centraldataorganization TO <download>.
    MOVE-CORRESPONDING addressdata TO <download>.
    CLEAR: addressdata.

    " Read all Relationships
    CLEAR: relationships.
    CALL FUNCTION 'BAPI_BUPA_RELATIONSHIPS_GET'
      EXPORTING
        businesspartner = <searchresult_line>-partner
      TABLES
        relationships   = relationships.
    DELETE relationships WHERE relationshipcategory <> 'BUR001'.
    LOOP AT relationships ASSIGNING <relationship>.
      " Read details of Contact person
      CALL FUNCTION 'BAPI_BUPA_CENTRAL_GETDETAIL'
        EXPORTING
          businesspartner   = <relationship>-partner2
        IMPORTING
          centraldata       = centraldata
          centraldataperson = centraldataperson.
      " Read details of the Relationship
      CALL FUNCTION 'BAPI_BUPR_CONTP_GETDETAIL'
        EXPORTING
          businesspartner = <relationship>-partner1
          contactperson   = <relationship>-partner2
        IMPORTING
          centraldata     = relationship_centraldata.
      " Read relationship address
      CLEAR: relationship_addresses.

      CALL FUNCTION 'BAPI_BUPR_CONTP_ADDRESSES_GET'
        EXPORTING
          businesspartner = <relationship>-partner1
          contactperson   = <relationship>-partner2
        TABLES
          addresses       = relationship_addresses.

      READ TABLE relationship_addresses
        ASSIGNING <relationship_address>
        WITH KEY standardaddress = 'X'.

      IF <relationship_address> IS ASSIGNED.
        " Read Relationship Address
        CLEAR addressdata.
        CALL FUNCTION 'BAPI_BUPA_ADDRESS_GETDETAIL'
          EXPORTING
            businesspartner = <searchresult_line>-partner
            addressguid     = <relationship_address>-addressguid
          IMPORTING
            addressdata     = addressdata.

        APPEND INITIAL LINE TO lt_download ASSIGNING <download>.
        CALL FUNCTION 'BAPI_BUPA_GET_NUMBERS'
          EXPORTING
            businesspartner        = <relationship>-partner1
          IMPORTING
            businesspartnerout     = <download>-org_number
            businesspartnerguidout = <download>-org_guid.

        CALL FUNCTION 'BAPI_BUPA_GET_NUMBERS'
          EXPORTING
            businesspartner        = <relationship>-partner2
          IMPORTING
            businesspartnerout     = <download>-contpers_number
            businesspartnerguidout = <download>-contpers_guid.

        MOVE-CORRESPONDING centraldataorganization TO <download>.
        MOVE-CORRESPONDING addressdata TO <download>.
        MOVE-CORRESPONDING centraldataperson TO <download>.
        MOVE-CORRESPONDING relationship_centraldata TO <download>.

        WRITE: / <relationship>-partner1, <relationship>-partner2.
        WRITE: centraldataorganization-name1(20), centraldataorganization-name2(10).
        WRITE: centraldataperson-firstname(15), centraldataperson-lastname(15).
        WRITE: addressdata-street(25), addressdata-house_no,
               addressdata-postl_cod1, addressdata-city(25).
      ENDIF.
    ENDLOOP.

  ENDLOOP.

  DATA: lo_excel                TYPE REF TO zcl_excel,
        lo_worksheet            TYPE REF TO zcl_excel_worksheet,
        lo_style_body           TYPE REF TO zcl_excel_style,
        lo_border_dark          TYPE REF TO zcl_excel_style_border,
        lo_column               TYPE REF TO zcl_excel_column,
        lo_row                  TYPE REF TO zcl_excel_row.

  DATA: lv_style_body_even_guid   TYPE zexcel_cell_style,
        lv_style_body_green       TYPE zexcel_cell_style.

  DATA: row TYPE zexcel_cell_row.

  DATA: lv_file                 TYPE xstring,
        lv_bytecount            TYPE i,
        lt_file_tab             TYPE solix_tab.

  DATA: lt_field_catalog        TYPE zexcel_t_fieldcatalog,
        ls_table_settings       TYPE zexcel_s_table_settings.

  DATA: column        TYPE zexcel_cell_column,
        column_alpha  TYPE zexcel_cell_column_alpha,
        value	        TYPE zexcel_cell_value.

  FIELD-SYMBOLS: <fs_field_catalog> TYPE zexcel_s_fieldcatalog.

  " Creates active sheet
  CREATE OBJECT lo_excel.

  " Create border object
  CREATE OBJECT lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
  "Create style with border even
  lo_style_body                       = lo_excel->add_new_style( ).
  lo_style_body->fill->fgcolor-rgb    = zcl_excel_style_color=>c_yellow.
  lo_style_body->borders->allborders  = lo_border_dark.
  lv_style_body_even_guid             = lo_style_body->get_guid( ).
  "Create style with border and green fill
  lo_style_body                       = lo_excel->add_new_style( ).
  lo_style_body->fill->fgcolor-rgb    = zcl_excel_style_color=>c_green.
  lo_style_body->borders->allborders  = lo_border_dark.
  lv_style_body_green                 = lo_style_body->get_guid( ).

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_download ).

  LOOP AT lt_field_catalog ASSIGNING <fs_field_catalog>.
    CASE <fs_field_catalog>-fieldname.
      WHEN 'ORG_NUMBER'.
        <fs_field_catalog>-position   = 1.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'CONTPERS_NUMBER'.
        <fs_field_catalog>-position   = 2.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'NAME1'.
        <fs_field_catalog>-position   = 3.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'NAME2'.
        <fs_field_catalog>-position   = 4.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'STREET'.
        <fs_field_catalog>-position   = 5.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'HOUSE_NO'.
        <fs_field_catalog>-position   = 6.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'POSTL_COD1'.
        <fs_field_catalog>-position   = 7.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'CITY'.
        <fs_field_catalog>-position   = 8.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'COUNTRYISO'.
        <fs_field_catalog>-position   = 9.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'FIRSTNAME'.
        <fs_field_catalog>-position   = 10.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'LASTNAME'.
        <fs_field_catalog>-position   = 11.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'FUNCTIONNAME'.
        <fs_field_catalog>-position   = 12.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'DEPARTMENTNAME'.
        <fs_field_catalog>-position   = 13.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'TEL1_NUMBR'.
        <fs_field_catalog>-position   = 14.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'TEL1_EXT'.
        <fs_field_catalog>-position   = 15.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'FAX_NUMBER'.
        <fs_field_catalog>-position   = 16.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'FAX_EXTENS'.
        <fs_field_catalog>-position   = 17.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN 'E_MAIL'.
        <fs_field_catalog>-position   = 18.
        <fs_field_catalog>-dynpfld    = abap_true.
      WHEN OTHERS.
        <fs_field_catalog>-dynpfld = abap_false.
    ENDCASE.
  ENDLOOP.

  ls_table_settings-top_left_column = 'A'.
  ls_table_settings-top_left_row    = '2'.
  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium5.

  lo_worksheet->bind_table( ip_table          = lt_download
                            is_table_settings = ls_table_settings
                            it_field_catalog  = lt_field_catalog ).
  LOOP AT lt_download ASSIGNING <download>.
    row = sy-tabix + 2.
    IF    NOT <download>-org_number IS INITIAL
      AND <download>-contpers_number IS INITIAL.
      " Mark fields of Organization which can be changed green
      lo_worksheet->set_cell_style(
          ip_column = 'C'
          ip_row    = row
          ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'D'
          ip_row    = row
          ip_style  = lv_style_body_green
      ).
*        CATCH zcx_excel.    " Exceptions for ABAP2XLSX
    ELSEIF NOT <download>-org_number IS INITIAL
       AND NOT <download>-contpers_number IS INITIAL.
      " Mark fields of Relationship which can be changed green
      lo_worksheet->set_cell_style(
          ip_column = 'L' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'M' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'N' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'O' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'P' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'Q' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'R' ip_row    = row ip_style  = lv_style_body_green
      ).
    ENDIF.
  ENDLOOP.
  " Add Fieldnames in first row and hide the row
  LOOP AT lt_field_catalog ASSIGNING <fs_field_catalog>
    WHERE position <> '' AND dynpfld = abap_true.
    column = <fs_field_catalog>-position.
    column_alpha = zcl_excel_common=>convert_column2alpha( column ).
    value = <fs_field_catalog>-fieldname.
    lo_worksheet->set_cell( ip_column = column_alpha
                            ip_row    = 1
                            ip_value  = value
                            ip_style  = lv_style_body_even_guid ).
  ENDLOOP.
  " Hide first row
  lo_row = lo_worksheet->get_row( 1 ).
  lo_row->set_visible( abap_false ).

  DATA: highest_column TYPE zexcel_cell_column,
        count TYPE int4,
        col_alpha TYPE zexcel_cell_column_alpha.

  highest_column = lo_worksheet->get_highest_column( ).
  count = 1.
  WHILE count <= highest_column.
    col_alpha = zcl_excel_common=>convert_column2alpha( ip_column = count ).
    lo_column = lo_worksheet->get_column( ip_column = col_alpha ).
    lo_column->set_auto_size( ip_auto_size = abap_true ).
    count = count + 1.
  ENDWHILE.

*** Create output
  lcl_output=>output( lo_excel ).
