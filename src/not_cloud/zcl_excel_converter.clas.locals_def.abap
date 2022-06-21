*"* use this source file for any type declarations (class
*"* definitions, interfaces or data types) you need for method
*"* implementation or private method's signature

TYPES ty_style_type TYPE c LENGTH 1.

TYPES: BEGIN OF ts_alv_types,
         seoclass TYPE seoclsname,
         clsname  TYPE seoclsname,
       END OF ts_alv_types,
       tt_alv_types TYPE HASHED TABLE OF ts_alv_types WITH UNIQUE KEY seoclass.

TYPES: BEGIN OF ts_sort_values,
         fieldname    TYPE fieldname,
         row_int      TYPE zexcel_cell_row,
         value        TYPE REF TO data,
         new          TYPE flag,
         sort_level   TYPE int4,
         is_collapsed TYPE flag,
       END OF ts_sort_values,
* Begin of ATC fix-issue-1014-part1
*       tt_sort_values TYPE HASHED TABLE OF ts_sort_values WITH UNIQUE KEY fieldname .
        tt_sort_values TYPE HASHED TABLE OF ts_sort_values
                       WITH UNIQUE KEY primary_key COMPONENTS fieldname
                       WITH NON-UNIQUE SORTED KEY sort_level  COMPONENTS  sort_level is_collapsed .

* End of ATC fix-issue-1014-part1
TYPES: BEGIN OF ts_subtotal_rows,
         row_int       TYPE zexcel_cell_row,
         row_int_start TYPE zexcel_cell_row,
         columnname    TYPE fieldname,
       END OF ts_subtotal_rows,

       tt_subtotal_rows TYPE HASHED TABLE OF ts_subtotal_rows WITH UNIQUE KEY row_int.

TYPES: BEGIN OF ts_styles,
         type      TYPE ty_style_type,
         alignment TYPE zexcel_alignment,
         inttype   TYPE abap_typekind,
         decimals  TYPE int1,
         style     TYPE REF TO zcl_excel_style,
         guid      TYPE zexcel_cell_style,
       END OF ts_styles,
* Begin of ATC fix-issue-1014-part1
*       tt_styles TYPE HASHED TABLE OF ts_styles  WITH UNIQUE KEY type alignment inttype decimals.
        tt_styles TYPE HASHED TABLE OF ts_styles with UNIQUE KEY primary_key
                                COMPONENTS type alignment inttype decimals
                                WITH NON-UNIQUE SORTED KEY guid  COMPONENTS  guid .
        .
* End of ATC fix-issue-1014-part1

TYPES: BEGIN OF ts_color_styles,
         guid_old  TYPE zexcel_cell_style,
         fontcolor TYPE zexcel_style_color_argb,
         fillcolor TYPE zexcel_style_color_argb,
         style_new TYPE REF TO zcl_excel_style,
       END OF ts_color_styles,

       tt_color_styles TYPE HASHED TABLE OF ts_color_styles  WITH UNIQUE KEY guid_old fontcolor fillcolor.
