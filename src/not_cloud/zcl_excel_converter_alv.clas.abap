CLASS zcl_excel_converter_alv DEFINITION
  PUBLIC
  ABSTRACT
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!
  PUBLIC SECTION.

    INTERFACES zif_excel_converter
      ALL METHODS ABSTRACT .

    CLASS-METHODS class_constructor .
*"* protected components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!
  PROTECTED SECTION.

    DATA wt_sort TYPE lvc_t_sort .
    DATA wt_filt TYPE lvc_t_filt .
    DATA wt_fcat TYPE lvc_t_fcat .
    DATA ws_layo TYPE lvc_s_layo .
    DATA ws_option TYPE zexcel_s_converter_option .

    METHODS update_catalog
      CHANGING
        !cs_layout       TYPE zexcel_s_converter_layo
        !ct_fieldcatalog TYPE zexcel_t_converter_fcat .
    METHODS apply_sort
      IMPORTING
        !it_table TYPE STANDARD TABLE
      EXPORTING
        !eo_table TYPE REF TO data .
    METHODS get_color
      IMPORTING
        !io_table  TYPE REF TO data
      EXPORTING
        !et_colors TYPE zexcel_t_converter_col .
    METHODS get_filter
      EXPORTING
        !et_filter TYPE zexcel_t_converter_fil
      CHANGING
        !xo_table  TYPE REF TO data .
*"* private components of class ZCL_EXCEL_CONVERTER_ALV
*"* do not include other source files here!!!
  PRIVATE SECTION.

    CLASS-DATA wt_colors TYPE tt_col_converter .
ENDCLASS.



CLASS zcl_excel_converter_alv IMPLEMENTATION.


  METHOD apply_sort.
    DATA: lt_otab TYPE abap_sortorder_tab,
          ls_otab TYPE abap_sortorder.

    FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE,
                   <fs_sort>  TYPE lvc_s_sort.

    CREATE DATA eo_table LIKE it_table.
    ASSIGN eo_table->* TO <fs_table>.

    <fs_table> = it_table.

    SORT wt_sort BY spos.
    LOOP AT wt_sort ASSIGNING <fs_sort>.
      IF <fs_sort>-up = abap_true.
        ls_otab-name = <fs_sort>-fieldname.
        ls_otab-descending = abap_false.
*     ls_otab-astext     = abap_true. " not only text fields
        INSERT ls_otab INTO TABLE lt_otab.
      ENDIF.
      IF <fs_sort>-down = abap_true.
        ls_otab-name = <fs_sort>-fieldname.
        ls_otab-descending = abap_true.
*     ls_otab-astext     = abap_true. " not only text fields
        INSERT ls_otab INTO TABLE lt_otab.
      ENDIF.
    ENDLOOP.
    IF lt_otab IS NOT INITIAL.
      SORT <fs_table> BY (lt_otab).
    ENDIF.

  ENDMETHOD.


  METHOD class_constructor.
* let's fill the color conversion routines.
    DATA: ls_color TYPE ts_col_converter.
* 0 all combination the same
    ls_color-col       = 0.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Blue
    ls_color-col       = 1.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFB0E4FC'. " 176 228 252 blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFB0E4FC'. " 176 228 252 blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF5FCBFE'. " 095 203 254 Int blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF5FCBFE'. " 095 203 254
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255
    INSERT ls_color INTO TABLE wt_colors.

* Gray
    ls_color-col       = 2.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'.
    ls_color-fillcolor = 'FFE5EAF0'. " 229 234 240 gray
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFE5EAF0'. " 229 234 240 gray
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFD8E8F4'. " 216 234 244 int gray
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFD8E8F4'. " 216 234 244 int gray
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

*Yellow
    ls_color-col       = 3.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFEFEB8'. " 254 254 184 yellow
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFEFEB8'. " 254 254 184 yellow
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFF9ED5D'. " 249 237 093 int yellow
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFF9ED5D'. " 249 237 093 int yellow
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* light blue
    ls_color-col       = 4.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFCEE7FB'. " 206 231 251 light blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFCEE7FB'. " 206 231 251 light blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF9ACCEF'. " 154 204 239 int light blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF9ACCEF'. " 154 204 239 int light blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Green
    ls_color-col       = 5.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFCEF8AE'. " 206 248 174 Green
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFCEF8AE'. " 206 248 174 Green
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF7AC769'. " 122 199 105 int Green
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF7AC769'. " 122 199 105 int Green
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Red
    ls_color-col       = 6.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFDBBBC'. " 253 187 188 Red
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFDBBBC'. " 253 187 188 Red
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFB6B6B'. " 251 107 107 int Red
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFB6B6B'. " 251 107 107 int Red
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

  ENDMETHOD.


  METHOD get_color.
    DATA: ls_con_col TYPE zexcel_s_converter_col,
          ls_color   TYPE ts_col_converter,
          l_line     TYPE i,
          l_color(4) TYPE c.
    FIELD-SYMBOLS: <fs_tab>  TYPE STANDARD TABLE,
                   <fs_stab> TYPE any,
                   <fs>      TYPE any,
                   <ft_slis> TYPE STANDARD TABLE,
                   <fs_slis> TYPE any.

* Loop trough the table to set the color properties of each line. The color properties field is
* Char 4 and the characters is set as follows:
* Char 1 = C = This is a color property
* Char 2 = 6 = Color code (1 - 7)
* Char 3 = Intensified on/of = 1 = on
* Char 4 = Inverse display = 0 = of

    ASSIGN io_table->* TO <fs_tab>.

    IF ws_layo-info_fname IS NOT INITIAL OR
       ws_layo-ctab_fname IS NOT INITIAL.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_line = sy-tabix.
        IF ws_layo-info_fname IS NOT INITIAL.
          ASSIGN COMPONENT ws_layo-info_fname OF STRUCTURE <fs_stab> TO <fs>.
          IF sy-subrc = 0 AND <fs> IS NOT INITIAL.
            l_color = <fs>.
            IF l_color(1) = 'C'.
              READ TABLE wt_colors INTO ls_color WITH TABLE KEY col = l_color+1(1)
                                                                int = l_color+2(1)
                                                                inv = l_color+3(1).
              IF sy-subrc = 0.
                ls_con_col-rownumber  = l_line.
                ls_con_col-columnname = space.
                ls_con_col-fontcolor  = ls_color-fontcolor.
                ls_con_col-fillcolor  = ls_color-fillcolor.
                INSERT ls_con_col INTO TABLE et_colors.
              ENDIF.
            ENDIF.
          ENDIF.
        ENDIF.
        IF ws_layo-ctab_fname IS NOT INITIAL.

          ASSIGN COMPONENT ws_layo-ctab_fname OF STRUCTURE <fs_stab> TO <ft_slis>.
          IF sy-subrc = 0.
            LOOP AT <ft_slis> ASSIGNING <fs_slis>.
              ASSIGN COMPONENT 'COLOR' OF STRUCTURE <fs_slis> TO <fs>.
              IF sy-subrc = 0.
                IF <fs> IS NOT INITIAL.
                  FIELD-SYMBOLS: <col>      TYPE any,
                                 <int>      TYPE any,
                                 <inv>      TYPE any,
                                 <fname>    TYPE any,
                                 <nokeycol> TYPE any.
                  ASSIGN COMPONENT 'COL' OF STRUCTURE <fs> TO <col>.
                  ASSIGN COMPONENT 'INT' OF STRUCTURE <fs> TO <int>.
                  ASSIGN COMPONENT 'INV' OF STRUCTURE <fs> TO <inv>.
                  READ TABLE wt_colors INTO ls_color WITH TABLE KEY col = <col>
                                                                    int = <int>
                                                                    inv = <inv>.
                  IF sy-subrc = 0.
                    ls_con_col-rownumber  = l_line.
                    ASSIGN COMPONENT 'FNAME' OF STRUCTURE <fs_slis> TO <fname>.
                    IF sy-subrc NE 0.
                      ASSIGN COMPONENT 'FIELDNAME' OF STRUCTURE <fs_slis> TO <fname>.
                      IF sy-subrc EQ 0.
                        ls_con_col-columnname = <fname>.
                      ENDIF.
                    ELSE.
                      ls_con_col-columnname = <fname>.
                    ENDIF.

                    ls_con_col-fontcolor  = ls_color-fontcolor.
                    ls_con_col-fillcolor  = ls_color-fillcolor.
                    ASSIGN COMPONENT 'NOKEYCOL' OF STRUCTURE <fs_slis> TO <nokeycol>.
                    IF sy-subrc EQ 0.
                      ls_con_col-nokeycol   = <nokeycol>.
                    ENDIF.
                    INSERT ls_con_col INTO TABLE et_colors.
                  ENDIF.
                ENDIF.
              ENDIF.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.


  METHOD get_filter.
    DATA: ls_filt   TYPE lvc_s_filt,
          l_line    TYPE i,
          ls_filter TYPE zexcel_s_converter_fil.
    DATA: lo_addit          TYPE REF TO cl_abap_elemdescr,
          lt_components_tab TYPE cl_abap_structdescr=>component_table,
          ls_components     TYPE abap_componentdescr,
          lo_table          TYPE REF TO cl_abap_tabledescr,
          lo_struc          TYPE REF TO cl_abap_structdescr,
          lo_trange         TYPE REF TO data,
          lo_srange         TYPE REF TO data,
          lo_ltabdata       TYPE REF TO data.

    FIELD-SYMBOLS: <fs_tab>    TYPE STANDARD TABLE,
                   <fs_ltab>   TYPE STANDARD TABLE,
                   <fs_stab>   TYPE any,
                   <fs>        TYPE any,
                   <fs1>       TYPE any,
                   <fs_srange> TYPE any,
                   <fs_trange> TYPE STANDARD TABLE.

    IF ws_option-filter = abap_false.
      CLEAR et_filter.
      RETURN.
    ENDIF.

    ASSIGN xo_table->* TO <fs_tab>.

    CREATE DATA lo_ltabdata LIKE <fs_tab>.
    ASSIGN lo_ltabdata->* TO <fs_ltab>.

    LOOP AT wt_filt INTO ls_filt.
      LOOP AT <fs_tab> ASSIGNING <fs_stab>.
        l_line = sy-tabix.
        ASSIGN COMPONENT ls_filt-fieldname OF STRUCTURE <fs_stab> TO <fs>.
        IF sy-subrc = 0.
          IF l_line = 1.
            CLEAR lt_components_tab.
            ls_components-name   = 'SIGN'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( ls_filt-sign ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'OPTION'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( ls_filt-option ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'LOW'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            ls_components-name   = 'HIGH'.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
            ls_components-type   = lo_addit           .
            INSERT ls_components INTO TABLE lt_components_tab.
            "create new line type
            TRY.
                lo_struc = cl_abap_structdescr=>create( p_components = lt_components_tab
                                                        p_strict     = abap_false ).
              CATCH cx_sy_struct_creation.
                CONTINUE.
            ENDTRY.
            lo_table = cl_abap_tabledescr=>create( lo_struc ).

            CREATE DATA lo_trange  TYPE HANDLE lo_table.
            CREATE DATA lo_srange  TYPE HANDLE lo_struc.

            ASSIGN lo_trange->* TO <fs_trange>.
            ASSIGN lo_srange->* TO <fs_srange>.
          ENDIF.
          CLEAR <fs_trange>.
          ASSIGN COMPONENT 'SIGN'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-sign.
          ASSIGN COMPONENT 'OPTION' OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-option.
          ASSIGN COMPONENT 'LOW'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-low.
          ASSIGN COMPONENT 'HIGH'   OF STRUCTURE  <fs_srange> TO <fs1>.
          <fs1> = ls_filt-high.
          INSERT <fs_srange> INTO TABLE <fs_trange>.
          IF <fs> IN <fs_trange>.
            IF ws_option-filter = abap_true.
              ls_filter-rownumber   = l_line.
              ls_filter-columnname  = ls_filt-fieldname.
              INSERT ls_filter INTO TABLE et_filter.
            ELSE.
              INSERT <fs_stab> INTO TABLE <fs_ltab>.
            ENDIF.
          ENDIF.
        ENDIF.
      ENDLOOP.
      IF ws_option-filter = abap_undefined.
        <fs_tab> = <fs_ltab>.
        CLEAR <fs_ltab>.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD update_catalog.
    DATA: ls_fieldcatalog TYPE zexcel_s_converter_fcat,
          ls_fcat         TYPE lvc_s_fcat,
          ls_sort         TYPE lvc_s_sort,
          l_decimals      TYPE lvc_decmls.

    FIELD-SYMBOLS: <fs_scat> TYPE zexcel_s_converter_fcat.

    IF ws_layo-zebra IS NOT INITIAL.
      cs_layout-is_stripped = abap_true.
    ENDIF.
    IF ws_layo-no_keyfix IS INITIAL OR
       ws_layo-no_keyfix = '0'.
      cs_layout-is_fixed = abap_true.
    ENDIF.

    LOOP AT wt_fcat INTO ls_fcat.
      CLEAR: ls_fieldcatalog,
             l_decimals.
      CASE ws_option-hidenc.
        WHEN abap_false.     " We make hiden columns visible
          CLEAR ls_fcat-no_out.
        WHEN abap_true.
* We convert column and hide it.
        WHEN abap_undefined. "We don't convert hiden columns
          IF ls_fcat-no_out = abap_true.
            ls_fcat-tech = abap_true.
          ENDIF.
      ENDCASE.
      IF ls_fcat-tech = abap_false.
        ls_fieldcatalog-tabname         = ls_fcat-tabname.
        ls_fieldcatalog-fieldname       = ls_fcat-fieldname .
        ls_fieldcatalog-columnname      = ls_fcat-fieldname .
        ls_fieldcatalog-position        = ls_fcat-col_pos.
        ls_fieldcatalog-col_id          = ls_fcat-col_id.
        ls_fieldcatalog-convexit        = ls_fcat-convexit.
        ls_fieldcatalog-inttype         = ls_fcat-inttype.
        ls_fieldcatalog-scrtext_s       = ls_fcat-scrtext_s .
        ls_fieldcatalog-scrtext_m       = ls_fcat-scrtext_m .
        ls_fieldcatalog-scrtext_l       = ls_fcat-scrtext_l.
        l_decimals = ls_fcat-decimals_o.
        IF l_decimals IS NOT INITIAL.
          ls_fieldcatalog-decimals = l_decimals.
        ELSE.
          ls_fieldcatalog-decimals = ls_fcat-decimals .
        ENDIF.
        CASE ws_option-subtot.
          WHEN abap_false.     " We ignore subtotals
            CLEAR ls_fcat-do_sum.
          WHEN abap_true.      " We convert subtotals and detail

          WHEN abap_undefined. " We should only take subtotals and displayed detail
* for now abap_true
        ENDCASE.
        CASE ls_fcat-do_sum.
          WHEN abap_true.
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_sum.
          WHEN 'A'.
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_min.
          WHEN 'B' .
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_max.
          WHEN 'C' .
            ls_fieldcatalog-totals_function =  zcl_excel_table=>totals_function_average.
          WHEN OTHERS.
            CLEAR ls_fieldcatalog-totals_function .
        ENDCASE.
        ls_fieldcatalog-fix_column       = ls_fcat-fix_column.
        IF ws_layo-cwidth_opt IS INITIAL.
          IF ls_fcat-col_opt IS NOT INITIAL.
            ls_fieldcatalog-is_optimized = abap_true.
          ENDIF.
        ELSE.
          ls_fieldcatalog-is_optimized = abap_true.
        ENDIF.
        IF ls_fcat-no_out IS NOT INITIAL.
          ls_fieldcatalog-is_hidden = abap_true.
          ls_fieldcatalog-position  = ls_fieldcatalog-col_id. " We hide based on orginal data structure
        ENDIF.
* Alignment in each cell
        CASE ls_fcat-just.
          WHEN 'R'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_right.
          WHEN 'L'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_left.
          WHEN 'C'.
            ls_fieldcatalog-alignment = zcl_excel_style_alignment=>c_horizontal_center.
          WHEN OTHERS.
            CLEAR ls_fieldcatalog-alignment.
        ENDCASE.
* Check for subtotals.
        READ TABLE wt_sort INTO ls_sort WITH KEY fieldname = ls_fcat-fieldname.
        IF sy-subrc = 0 AND  ws_option-subtot <> abap_false.
          ls_fieldcatalog-sort_level      = 0 .
          ls_fieldcatalog-is_subtotalled  = ls_sort-subtot.
          ls_fieldcatalog-is_collapsed    = ls_sort-expa.
          IF ls_fieldcatalog-is_subtotalled = abap_true.
            ls_fieldcatalog-sort_level      = ls_sort-spos.
            ls_fieldcatalog-totals_function = zcl_excel_table=>totals_function_sum. " we need function for text
          ENDIF.
        ENDIF.
        APPEND ls_fieldcatalog TO ct_fieldcatalog.
      ENDIF.
    ENDLOOP.

    SORT ct_fieldcatalog BY sort_level ASCENDING.
    cs_layout-max_subtotal_level  = 0.
    LOOP AT ct_fieldcatalog ASSIGNING <fs_scat> WHERE sort_level > 0.
      cs_layout-max_subtotal_level = cs_layout-max_subtotal_level + 1.
      <fs_scat>-sort_level = cs_layout-max_subtotal_level.
    ENDLOOP.

  ENDMETHOD.
ENDCLASS.
