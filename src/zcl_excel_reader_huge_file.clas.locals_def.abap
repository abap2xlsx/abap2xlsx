*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section

* Signal for "Not found"
class lcx_not_found definition inheriting from cx_static_check.
  public section.
    data error type string.
    methods constructor
      importing error type string
                textid type sotr_conc optional
                previous type ref to cx_root optional.
    methods if_message~get_text redefinition.
endclass.
