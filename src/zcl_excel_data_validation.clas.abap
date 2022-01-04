CLASS zcl_excel_data_validation DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

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
ENDCLASS.
