*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations

* Signal "not found"
CLASS lcx_not_found IMPLEMENTATION.
  METHOD constructor.
    super->constructor( textid = textid previous = previous ).
    me->error = error.
  ENDMETHOD.                    "constructor
  METHOD if_message~get_text.
    result = error.
  ENDMETHOD.                    "if_message~get_text
ENDCLASS.                    "lcx_not_found IMPLEMENTATION
