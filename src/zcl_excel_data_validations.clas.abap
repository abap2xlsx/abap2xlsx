class ZCL_EXCEL_DATA_VALIDATIONS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  methods ADD
    importing
      !IP_DATA_VALIDATION type ref to ZCL_EXCEL_DATA_VALIDATION .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IP_DATA_VALIDATION type ref to ZCL_EXCEL_DATA_VALIDATION .
  methods SIZE
    returning
      value(EP_SIZE) type I .
*"* protected components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_DATA_VALIDATIONS
*"* do not include other source files here!!!
private section.

  data DATA_VALIDATIONS type ref to CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_DATA_VALIDATIONS IMPLEMENTATION.


method ADD.
  data_validations->add( ip_data_validation ).
  endmethod.


method CLEAR.
  data_validations->clear( ).
  endmethod.


method CONSTRUCTOR.

  CREATE OBJECT data_validations.

  endmethod.


method GET_ITERATOR.
  eo_iterator ?= data_validations->if_object_collection~get_iterator( ).
  endmethod.


method IS_EMPTY.
  is_empty = data_validations->if_object_collection~is_empty( ).
  endmethod.


method REMOVE.
  data_validations->remove( ip_data_validation ).
  endmethod.


method SIZE.
  ep_size = data_validations->if_object_collection~size( ).
  endmethod.
ENDCLASS.
