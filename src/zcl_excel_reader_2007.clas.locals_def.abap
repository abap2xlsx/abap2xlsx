*"* use this source file for any type declarations (class
*"* definitions, interfaces or data types) you need for method
*"* implementation or private method's signature

*
class lcl_zip_archive definition abstract.
  public section.
    methods read abstract
      importing i_filename type csequence
      returning value(r_content) type xstring " Remember copy-on-write!
      raising zcx_excel.
endclass.
