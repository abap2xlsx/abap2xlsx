CLASS zcl_excel_collection DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      ty_collection TYPE STANDARD TABLE OF REF TO object .

    DATA collection TYPE ty_collection READ-ONLY .

    METHODS size
      RETURNING
        VALUE(size) TYPE i .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE abap_bool .
    METHODS get
      IMPORTING
        !index        TYPE i
      RETURNING
        VALUE(object) TYPE REF TO object .
    METHODS get_iterator
      RETURNING
        VALUE(iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS add
      IMPORTING
        !element TYPE REF TO object .
    METHODS remove
      IMPORTING
        !element TYPE REF TO object .
    METHODS clear .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_collection IMPLEMENTATION.


  METHOD add .
    APPEND element TO collection.
  ENDMETHOD.


  METHOD clear .
    CLEAR collection.
  ENDMETHOD.


  METHOD get .
    READ TABLE collection INDEX index INTO object.
  ENDMETHOD.


  METHOD get_iterator .
    CREATE OBJECT iterator
      EXPORTING
        collection = me.
  ENDMETHOD.


  METHOD is_empty.
    is_empty = boolc( size( ) = 0 ).
  ENDMETHOD.


  METHOD remove .
    DATA obj TYPE REF TO object.
    LOOP AT collection INTO obj.
      IF obj = element.
        DELETE collection.
        RETURN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD size.
    size = lines( collection ).
  ENDMETHOD.
ENDCLASS.
