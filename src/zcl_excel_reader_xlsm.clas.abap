CLASS zcl_excel_reader_xlsm DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_reader_2007
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_READER_XLSM
*"* do not include other source files here!!!
  PROTECTED SECTION.

*"* protected components of class ZCL_EXCEL_READER_XLSM
*"* do not include other source files here!!!
    METHODS load_workbook
        REDEFINITION .
    METHODS load_worksheet
        REDEFINITION .
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_READER_XLSM
*"* do not include other source files here!!!
    METHODS load_vbaproject
      IMPORTING
        !ip_path  TYPE string
        !ip_excel TYPE REF TO zcl_excel
      RAISING
        zcx_excel .
ENDCLASS.



CLASS zcl_excel_reader_xlsm IMPLEMENTATION.


  METHOD load_vbaproject.

    DATA lv_content TYPE xstring.

    lv_content = me->get_from_zip_archive( ip_path ).

    ip_excel->zif_excel_book_vba_project~set_vbaproject( lv_content ).

  ENDMETHOD.


  METHOD load_workbook.
    super->load_workbook( EXPORTING iv_workbook_full_filename = iv_workbook_full_filename
                                    io_excel                  = io_excel ).

    CONSTANTS: lc_vba_project  TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'.

    DATA: rels_workbook_path TYPE string,
          rels_workbook      TYPE REF TO if_ixml_document,
          path               TYPE string,
          node               TYPE REF TO if_ixml_element,
          workbook           TYPE REF TO if_ixml_document,
          stripped_name      TYPE chkfile,
          dirname            TYPE string,
          relationship       TYPE t_relationship,
          fileversion        TYPE t_fileversion,
          workbookpr         TYPE t_workbookpr.

    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = iv_workbook_full_filename
      IMPORTING
        stripped_name = stripped_name
        file_path     = dirname.

    " Read Workbook Relationships
    CONCATENATE dirname '_rels/' stripped_name '.rels'
      INTO rels_workbook_path.

    rels_workbook = me->get_ixml_from_zip_archive( rels_workbook_path ).

    node ?= rels_workbook->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
    WHILE node IS BOUND.
      me->fill_struct_from_attributes( EXPORTING ip_element = node CHANGING cp_structure = relationship ).

      CASE relationship-type.
        WHEN lc_vba_project.
          " Read VBA  binary
          CONCATENATE dirname relationship-target INTO path.
          me->load_vbaproject( ip_path  = path
                               ip_excel = io_excel ).
        WHEN OTHERS.
      ENDCASE.

      node ?= node->get_next( ).
    ENDWHILE.

    " Read Workbook codeName
    workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).
    node ?=  workbook->find_from_name_ns( name = 'fileVersion' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = fileversion  ).

      io_excel->zif_excel_book_vba_project~set_codename( fileversion-codename ).
    ENDIF.

    " Read Workbook codeName
    workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).
    node ?=  workbook->find_from_name_ns( name = 'workbookPr' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = workbookpr  ).

      io_excel->zif_excel_book_vba_project~set_codename_pr( workbookpr-codename ).
    ENDIF.

  ENDMETHOD.


  METHOD load_worksheet.

    super->load_worksheet( EXPORTING ip_path      = ip_path
                                     io_worksheet = io_worksheet ).

    DATA: node      TYPE REF TO if_ixml_element,
          worksheet TYPE REF TO if_ixml_document,
          sheetpr   TYPE t_sheetpr.

    " Read Workbook codeName
    worksheet = me->get_ixml_from_zip_archive( ip_path ).
    node ?=  worksheet->find_from_name_ns( name = 'sheetPr' uri = namespace-main ).
    IF node IS BOUND.

      fill_struct_from_attributes( EXPORTING ip_element   = node
                                   CHANGING  cp_structure = sheetpr  ).
      IF sheetpr-codename IS NOT INITIAL.
        io_worksheet->zif_excel_sheet_vba_project~set_codename_pr( sheetpr-codename ).
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
