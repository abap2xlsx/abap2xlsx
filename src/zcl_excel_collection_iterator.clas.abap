CLASS zcl_excel_collection_iterator DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    METHODS get_index
      RETURNING
        VALUE(index) TYPE i.
    METHODS has_next
      RETURNING
        VALUE(has_next) TYPE abap_bool.
    METHODS get_next
      RETURNING
        VALUE(object) TYPE REF TO object.
    METHODS constructor
      IMPORTING
        collection TYPE REF TO zcl_excel_collection.
  PROTECTED SECTION.
  PRIVATE SECTION.
    DATA index TYPE i VALUE 0.
    DATA collection TYPE REF TO zcl_excel_collection.
ENDCLASS.



CLASS zcl_excel_collection_iterator IMPLEMENTATION.


  METHOD constructor .
    me->collection = collection.
  ENDMETHOD.


  METHOD get_index .
    index = me->index.
  ENDMETHOD.


  METHOD get_next .
    DATA obj TYPE REF TO object.
    index = index + 1.
    object = collection->get( index ).
  ENDMETHOD.


  METHOD has_next.
    DATA obj TYPE REF TO object.
    obj = collection->get( index + 1 ).
    has_next = boolc( obj IS NOT INITIAL ).
  ENDMETHOD.
ENDCLASS.
