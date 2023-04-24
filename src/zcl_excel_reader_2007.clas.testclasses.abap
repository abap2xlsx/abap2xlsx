*"* use this source file for your ABAP unit test classes

CLASS ltc_unescape_string_value DEFINITION DEFERRED.
CLASS zcl_excel_reader_2007 DEFINITION LOCAL FRIENDS
    ltc_unescape_string_value.


CLASS ltc_unescape_string_value DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.
  PRIVATE SECTION.

    METHODS escaped_character_inside_text FOR TESTING.
    METHODS no_escaping FOR TESTING.
    METHODS one_escaped_character FOR TESTING.
    METHODS two_escaped_characters FOR TESTING.

    METHODS run_cut
      IMPORTING
        input TYPE string
        exp   TYPE string.

ENDCLASS.


CLASS ltc_unescape_string_value IMPLEMENTATION.

  METHOD escaped_character_inside_text.
    run_cut( input = 'start _x0000_ end' exp = |start { cl_abap_conv_in_ce=>uccp( '0000' ) } end| ).
  ENDMETHOD.

  METHOD no_escaping.
    run_cut( input = 'no escaping' exp = 'no escaping' ).
  ENDMETHOD.

  METHOD one_escaped_character.
    run_cut( input = '_x0000_' exp = cl_abap_conv_in_ce=>uccp( '0000' ) ).
  ENDMETHOD.

  METHOD run_cut.

    DATA: lo_excel TYPE REF TO zcl_excel_reader_2007.

    CREATE OBJECT lo_excel.

    cl_abap_unit_assert=>assert_equals( act = lo_excel->unescape_string_value( input ) exp = exp msg = |input: { input }| ).

  ENDMETHOD.

  METHOD two_escaped_characters.
    run_cut( input = '_x0000_ and _xFFFF_' exp = |{ cl_abap_conv_in_ce=>uccp( '0000' ) } and { cl_abap_conv_in_ce=>uccp( 'FFFF' ) }| ).
  ENDMETHOD.

ENDCLASS.
