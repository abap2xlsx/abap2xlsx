CLASS ltcl_test DEFINITION FOR TESTING DURATION SHORT RISK LEVEL HARMLESS FINAL.

  PRIVATE SECTION.
    METHODS test01 FOR TESTING RAISING cx_static_check.
ENDCLASS.


CLASS ltcl_test IMPLEMENTATION.

  METHOD test01.

    DATA lo_collection TYPE REF TO zcl_excel_collection.

    CREATE OBJECT lo_collection.

    cl_abap_unit_assert=>assert_equals(
      act = lo_collection->size( )
      exp = 0 ).

    cl_abap_unit_assert=>assert_equals(
      act = lo_collection->is_empty( )
      exp = abap_true ).

* heh, yea, add the collection to itself :)
    lo_collection->add( lo_collection ).

    cl_abap_unit_assert=>assert_equals(
      act = lo_collection->size( )
      exp = 1 ).

    cl_abap_unit_assert=>assert_equals(
      act = lo_collection->is_empty( )
      exp = abap_false ).

  ENDMETHOD.

ENDCLASS.
