*&---------------------------------------------------------------------*
*& Report  Fill Template
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel_fill_template.

TYPES
: begin of t_TABLE1
,     PERSON type string
,     SALARY type i
, end of t_TABLE1

, tt_TABLE1 type table of  t_TABLE1 with empty key

, begin of t_LINE1
,     CARRID type string
,     CONNID type string
,     FLDATE type string
,     PRICE type i
, end of t_LINE1

, tt_LINE1 type table of  t_LINE1 with empty key

, begin of t_TABLE2
,     CARRID type string
,     PRICE type i
,     LINE1 type tt_LINE1
, end of t_TABLE2

, tt_TABLE2 type table of  t_TABLE2 with empty key

, begin of t_Sheet1
,     DATE type string
,     TIME type string
,     USER type string
,     TOTAL type i
,     PRICE type i
,     TABLE1 type tt_TABLE1
,     TABLE2 type tt_TABLE2
, end of t_Sheet1


, begin of t_TABLE3
,     PERSON type string
,     SALARY type i
, end of t_TABLE3

, tt_TABLE3 type table of  t_TABLE3 with empty key

, begin of t_Sheet2
,     DATE type string
,     TIME type string
,     USER type string
,     TOTAL type i
,     TABLE3 type tt_TABLE3
, end of t_Sheet2
.



DATA
: lo_data type ref to ZCL_EXCEL_TEMPLATE_DATA
, gs_sheet1 TYPE   t_Sheet1
, gs_sheet2 TYPE   t_Sheet2
.

* define variables
data: lo_excel type ref to zcl_excel,
      reader   type ref to zif_excel_reader.


constants: gc_save_file_name type string value 'fill_template_example.xlsx'.
include zdemo_excel_outputopt_incl.

parameters: p_fpath type string obligatory lower case default 'C:\Users\sadfasdf\Desktop\abap2xlsx\ZABAP2XLSX_EXAMPLE.xlsx'.


parameters: p_file radiobutton group rad1 default 'X'
          , p_smw0 radiobutton group rad1
          .

at selection-screen on value-request for p_fpath.
  perform get_file_path changing p_fpath.


start-of-selection.

  create object lo_data.

* generate data
  gs_sheet1 = value #(
  date = |{ sy-datum date = environment }|
  time = |{ sy-uzeit time = environment }|
  user  = |{ sy-uname }|

  table1 = value #(
                    ( person = 'Lurch Schpellchek' salary = '1200' )
                    ( person = 'Russell Sprout'    salary = '1300' )
                    ( person = 'Fergus Douchebag'  salary = '3000' )
                    ( person = 'Bartholomew Shoe'  salary = '100' )
                  )

  total = '5600'
  table2 = value #(
                      ( line1 = value #(
                                         (  carrid = 'AC' connid = '0820'  fldate = '20.12.2002' price = '1222'  )
                                       )
                        carrid ='AC'
                        price = '1222'
                      )
                      ( line1 = value #(
                                          (  carrid = 'AF' connid = '0820'  fldate = '23.12.2002' price = '2222'  )
                                       )
                        carrid ='AF'
                        price = '2222'
                      )

                      ( line1 = value #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = value #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = value #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = value #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = value #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = value #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = value #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = value #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = value #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = value #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

  )

  price = '14003'
  ).


  gs_sheet2 = value #(
  date = |{ sy-datum date = environment }|
  time = |{ sy-uzeit time = environment }|
  user  = |{ sy-uname }|

  table3 = value #(
                    ( person = 'Lurch Schpellchek' salary = '1200' )
                    ( person = 'Russell Sprout'    salary = '1300' )
                    ( person = 'Fergus Douchebag'  salary = '3000' )
                    ( person = 'Bartholomew Shoe'  salary = '100' )
                  )

  total = '5600' ).

* add data
  lo_data->add( iv_sheet = 'Sheet1' iv_data = gs_sheet1 ).
  lo_data->add( iv_sheet = 'Sheet2' iv_data = gs_sheet2 ).

* create reader

  create object reader type zcl_excel_reader_2007.
* load template

  IF p_file is NOT INITIAL.
    lo_excel = reader->load_file( p_fpath ).
  else.
    lo_excel = reader->load_smw0( 'ZEXCEL_DEMO_TEMPLATE' ).
  ENDIF.


* merge data with template
  lo_excel->fill_template( lo_data ).

*** Create output
  lcl_output=>output( lo_excel ).


*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
form get_file_path changing cv_path type string.
  clear cv_path.

  data:
    lv_rc          type  i,
    lv_user_action type  i,
    lt_file_table  type  filetable,
    ls_file_table  like line of lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  exporting
    window_title        = 'select template  xlsx'
    multiselection      = ''
    default_extension   = '*.xlsx'
    file_filter         = 'Text file (*.xlsx)|*.xlsx|All (*.*)|*.*'
  changing
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  exceptions
    others              = 1
    ).
  if sy-subrc = 0.
    if lv_user_action = cl_gui_frontend_services=>action_ok.
      if lt_file_table is not initial.
        read table lt_file_table into ls_file_table index 1.
        if sy-subrc = 0.
          cv_path = ls_file_table-filename.
        endif.
      endif.
    endif.
  endif.
endform.                    " Get_file_path
