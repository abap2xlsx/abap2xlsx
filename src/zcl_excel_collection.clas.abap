CLASS zcl_excel_collection DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC
  INHERITING FROM zcl_excel_base.

  PUBLIC SECTION.

    TYPES:
      ty_collection TYPE STANDARD TABLE OF REF TO zcl_excel_base.

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
        VALUE(object) TYPE REF TO zcl_excel_base.
    METHODS get_iterator
      RETURNING
        VALUE(iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS add
      IMPORTING
        !element TYPE REF TO zcl_excel_base.
    METHODS remove
      IMPORTING
        !element TYPE REF TO zcl_excel_base.
    METHODS clear .
    METHODS clone REDEFINITION.
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

  METHOD clone.
    DATA lo_excel_collection TYPE REF TO zcl_excel_collection.
    DATA lo_element TYPE REF TO zcl_excel_base.
    DATA lo_clone TYPE REF TO zcl_excel_base.

    CREATE OBJECT lo_excel_collection.

    LOOP AT me->collection INTO lo_element.
      lo_clone = lo_element->clone( ).
      INSERT lo_clone INTO TABLE lo_excel_collection->collection.
    ENDLOOP.

    ro_object = lo_excel_collection.
  ENDMETHOD.

ENDCLASS.
