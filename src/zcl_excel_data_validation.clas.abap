CLASS zcl_excel_data_validation DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC
  INHERITING FROM zcl_excel_base.

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!

    DATA errorstyle TYPE zexcel_data_val_error_style .
    DATA operator TYPE zexcel_data_val_operator .
    DATA allowblank TYPE flag VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA cell_column TYPE zexcel_cell_column_alpha .
    DATA cell_column_to TYPE zexcel_cell_column_alpha .
    DATA cell_row TYPE zexcel_cell_row .
    DATA cell_row_to TYPE zexcel_cell_row .
    CONSTANTS c_type_custom TYPE zexcel_data_val_type VALUE 'custom'. "#EC NOTEXT
    CONSTANTS c_type_list TYPE zexcel_data_val_type VALUE 'list'. "#EC NOTEXT
    DATA showerrormessage TYPE abap_bool VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA showinputmessage TYPE abap_bool VALUE 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA type TYPE zexcel_data_val_type .
    DATA formula1 TYPE zexcel_validation_formula1 .
    DATA formula2 TYPE zexcel_validation_formula1 .
    CONSTANTS c_type_none TYPE zexcel_data_val_type VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_type_date TYPE zexcel_data_val_type VALUE 'date'. "#EC NOTEXT
    CONSTANTS c_type_decimal TYPE zexcel_data_val_type VALUE 'decimal'. "#EC NOTEXT
    CONSTANTS c_type_textlength TYPE zexcel_data_val_type VALUE 'textLength'. "#EC NOTEXT
    CONSTANTS c_type_time TYPE zexcel_data_val_type VALUE 'time'. "#EC NOTEXT
    CONSTANTS c_type_whole TYPE zexcel_data_val_type VALUE 'whole'. "#EC NOTEXT
    CONSTANTS c_style_stop TYPE zexcel_data_val_error_style VALUE 'stop'. "#EC NOTEXT
    CONSTANTS c_style_warning TYPE zexcel_data_val_error_style VALUE 'warning'. "#EC NOTEXT
    CONSTANTS c_style_information TYPE zexcel_data_val_error_style VALUE 'information'. "#EC NOTEXT
    CONSTANTS c_operator_between TYPE zexcel_data_val_operator VALUE 'between'. "#EC NOTEXT
    CONSTANTS c_operator_equal TYPE zexcel_data_val_operator VALUE 'equal'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthan TYPE zexcel_data_val_operator VALUE 'greaterThan'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthanorequal TYPE zexcel_data_val_operator VALUE 'greaterThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_lessthan TYPE zexcel_data_val_operator VALUE 'lessThan'. "#EC NOTEXT
    CONSTANTS c_operator_lessthanorequal TYPE zexcel_data_val_operator VALUE 'lessThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_notbetween TYPE zexcel_data_val_operator VALUE 'notBetween'. "#EC NOTEXT
    CONSTANTS c_operator_notequal TYPE zexcel_data_val_operator VALUE 'notEqual'. "#EC NOTEXT
    DATA showdropdown TYPE abap_bool .
    DATA errortitle TYPE string .
    DATA error TYPE string .
    DATA prompttitle TYPE string .
    DATA prompt TYPE string .

    METHODS constructor .
    METHODS clone REDEFINITION.
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_data_validation IMPLEMENTATION.


  METHOD constructor.
    super->constructor( ).
    " Initialise instance variables
    formula1          = ''.
    formula2          = ''.
    type              = me->c_type_none.
    errorstyle        = me->c_style_stop.
    operator          = ''.
    allowblank        = abap_false.
    showdropdown      = abap_false.
    showinputmessage  = abap_true.
    showerrormessage  = abap_true.
    errortitle        = ''.
    error             = ''.
    prompttitle       = ''.
    prompt            = ''.
* inizialize dimension range
    cell_row     = 1.
    cell_column  = 'A'.
  ENDMETHOD.


  METHOD clone.
    DATA lo_excel_data_validation TYPE REF TO zcl_excel_data_validation.

    CREATE OBJECT lo_excel_data_validation.

    lo_excel_data_validation->allowblank        = allowblank.
    lo_excel_data_validation->cell_column       = cell_column.
    lo_excel_data_validation->cell_column_to    = cell_column_to.
    lo_excel_data_validation->cell_row          = cell_row.
    lo_excel_data_validation->cell_row_to       = cell_row_to.
    lo_excel_data_validation->error             = error.
    lo_excel_data_validation->errorstyle        = errorstyle.
    lo_excel_data_validation->errortitle        = errortitle.
    lo_excel_data_validation->formula1          = formula1.
    lo_excel_data_validation->formula2          = formula2.
    lo_excel_data_validation->operator          = operator.
    lo_excel_data_validation->prompt            = prompt.
    lo_excel_data_validation->prompttitle       = prompttitle.
    lo_excel_data_validation->showdropdown      = showdropdown.
    lo_excel_data_validation->showerrormessage  = showerrormessage.
    lo_excel_data_validation->showinputmessage  = showinputmessage.
    lo_excel_data_validation->type              = type.

    ro_object = lo_excel_data_validation.
  ENDMETHOD.

ENDCLASS.
