*"* use this source file for any type declarations (class
*"* definitions, interfaces or data types) you need for method
*"* implementation or private method's signature
TYPES: BEGIN OF ts_col_converter,
           col       TYPE lvc_col,
           int       TYPE lvc_int,
           inv       TYPE lvc_inv,
           fontcolor TYPE zexcel_style_color_argb,
           fillcolor TYPE zexcel_style_color_argb,
       END OF ts_col_converter,

       tt_col_converter TYPE HASHED TABLE OF ts_col_converter WITH UNIQUE KEY col int inv.
