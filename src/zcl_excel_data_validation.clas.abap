class ZCL_EXCEL_DATA_VALIDATION definition
  public
  final
  create public .

public section.
*"* public components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
  type-pools ABAP .

  data ERRORSTYLE type ZEXCEL_DATA_VAL_ERROR_STYLE .
  data OPERATOR type ZEXCEL_DATA_VAL_OPERATOR .
  data ALLOWBLANK type FLAG value 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data CELL_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA .
  data CELL_COLUMN_TO type ZEXCEL_CELL_COLUMN_ALPHA .
  data CELL_ROW type ZEXCEL_CELL_ROW .
  data CELL_ROW_TO type ZEXCEL_CELL_ROW .
  constants C_TYPE_CUSTOM type ZEXCEL_DATA_VAL_TYPE value 'custom'. "#EC NOTEXT
  constants C_TYPE_LIST type ZEXCEL_DATA_VAL_TYPE value 'list'. "#EC NOTEXT
  data SHOWERRORMESSAGE type FLAG value 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data SHOWINPUTMESSAGE type FLAG value 'X'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data TYPE type ZEXCEL_DATA_VAL_TYPE .
  data FORMULA1 type ZEXCEL_VALIDATION_FORMULA1 .
  data FORMULA2 type ZEXCEL_VALIDATION_FORMULA1 .
  constants C_TYPE_NONE type ZEXCEL_DATA_VAL_TYPE value 'none'. "#EC NOTEXT
  constants C_TYPE_DATE type ZEXCEL_DATA_VAL_TYPE value 'date'. "#EC NOTEXT
  constants C_TYPE_DECIMAL type ZEXCEL_DATA_VAL_TYPE value 'decimal'. "#EC NOTEXT
  constants C_TYPE_TEXTLENGTH type ZEXCEL_DATA_VAL_TYPE value 'textLength'. "#EC NOTEXT
  constants C_TYPE_TIME type ZEXCEL_DATA_VAL_TYPE value 'time'. "#EC NOTEXT
  constants C_TYPE_WHOLE type ZEXCEL_DATA_VAL_TYPE value 'whole'. "#EC NOTEXT
  constants C_STYLE_STOP type ZEXCEL_DATA_VAL_ERROR_STYLE value 'stop'. "#EC NOTEXT
  constants C_STYLE_WARNING type ZEXCEL_DATA_VAL_ERROR_STYLE value 'warning'. "#EC NOTEXT
  constants C_STYLE_INFORMATION type ZEXCEL_DATA_VAL_ERROR_STYLE value 'information'. "#EC NOTEXT
  constants C_OPERATOR_BETWEEN type ZEXCEL_DATA_VAL_OPERATOR value 'between'. "#EC NOTEXT
  constants C_OPERATOR_EQUAL type ZEXCEL_DATA_VAL_OPERATOR value 'equal'. "#EC NOTEXT
  constants C_OPERATOR_GREATERTHAN type ZEXCEL_DATA_VAL_OPERATOR value 'greaterThan'. "#EC NOTEXT
  constants C_OPERATOR_GREATERTHANOREQUAL type ZEXCEL_DATA_VAL_OPERATOR value 'greaterThanOrEqual'. "#EC NOTEXT
  constants C_OPERATOR_LESSTHAN type ZEXCEL_DATA_VAL_OPERATOR value 'lessThan'. "#EC NOTEXT
  constants C_OPERATOR_LESSTHANOREQUAL type ZEXCEL_DATA_VAL_OPERATOR value 'lessThanOrEqual'. "#EC NOTEXT
  constants C_OPERATOR_NOTBETWEEN type ZEXCEL_DATA_VAL_OPERATOR value 'notBetween'. "#EC NOTEXT
  constants C_OPERATOR_NOTEQUAL type ZEXCEL_DATA_VAL_OPERATOR value 'notEqual'. "#EC NOTEXT
  data SHOWDROPDOWN type FLAG .
  data ERRORTITLE type STRING .
  data ERROR type STRING .
  data PROMPTTITLE type STRING .
  data PROMPT type STRING .

  methods CONSTRUCTOR .
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
protected section.
private section.
*"* private components of class ZCL_EXCEL_DATA_VALIDATION
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_DATA_VALIDATION IMPLEMENTATION.


method CONSTRUCTOR.
  " Initialise instance variables
  formula1          = ''.
  formula2          = ''.
  type              = me->c_type_none.
  errorstyle        = me->c_style_stop.
  operator          = ''.
  allowblank        = abap_false.
  showdropdown      = abap_false.
  showinputmessage 	= abap_true.
  showerrormessage 	= abap_true.
  errortitle        = ''.
  error             = ''.
  prompttitle       = ''.
  prompt            = ''.
* inizialize dimension range
  cell_row     = 1.
  cell_column  = 'A'.
  endmethod.
ENDCLASS.
