*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations

* Signal "not found"
class lcx_not_found implementation.
  method constructor.
    super->constructor( textid = textid previous = previous ).
    me->error = error.
  endmethod.                    "constructor
  method if_message~get_text.
    result = error.
  endmethod.                    "if_message~get_text
endclass.                    "lcx_not_found IMPLEMENTATION
