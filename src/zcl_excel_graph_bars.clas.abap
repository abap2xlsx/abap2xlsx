CLASS zcl_excel_graph_bars DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_graph
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
*"* public components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
      tv_type TYPE c LENGTH 5,
      BEGIN OF s_ax,
        axid          TYPE string,
        type          TYPE tv_type,
        orientation   TYPE string,
        delete        TYPE string,
        axpos         TYPE string,
        formatcode    TYPE string,
        sourcelinked  TYPE string,
        majortickmark TYPE string,
        minortickmark TYPE string,
        ticklblpos    TYPE string,
        crossax       TYPE string,
        crosses       TYPE string,
        auto          TYPE string,
        lblalgn       TYPE string,
        lbloffset     TYPE string,
        nomultilvllbl TYPE string,
        crossbetween  TYPE string,
      END OF s_ax .
    TYPES:
      t_ax TYPE STANDARD TABLE OF s_ax .

    DATA ns_bardirval TYPE string VALUE 'col'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_groupingval_clustered TYPE string VALUE 'clustered'. "#EC NOTEXT
    CONSTANTS c_groupingval_stacked TYPE string VALUE 'stacked'. "#EC NOTEXT
    DATA ns_groupingval TYPE string VALUE c_groupingval_clustered. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_varycolorsval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showlegendkeyval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showvalval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showcatnameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showsernameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showpercentval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showbubblesizeval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_gapwidthval TYPE string VALUE '150'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA axes TYPE t_ax .
    CONSTANTS:
      c_valax TYPE c LENGTH 5 VALUE 'VALAX'.                "#EC NOTEXT
    CONSTANTS:
      c_catax TYPE c LENGTH 5 VALUE 'CATAX'.                "#EC NOTEXT
    DATA ns_legendposval TYPE string VALUE 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA ns_overlayval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_invertifnegative_yes TYPE string VALUE '1'. "#EC NOTEXT
    CONSTANTS c_invertifnegative_no TYPE string VALUE '0'.  "#EC NOTEXT

    METHODS create_ax
      IMPORTING
        !ip_axid          TYPE string OPTIONAL
        !ip_type          TYPE tv_type
        !ip_orientation   TYPE string OPTIONAL
        !ip_delete        TYPE string OPTIONAL
        !ip_axpos         TYPE string OPTIONAL
        !ip_formatcode    TYPE string OPTIONAL
        !ip_sourcelinked  TYPE string OPTIONAL
        !ip_majortickmark TYPE string OPTIONAL
        !ip_minortickmark TYPE string OPTIONAL
        !ip_ticklblpos    TYPE string OPTIONAL
        !ip_crossax       TYPE string OPTIONAL
        !ip_crosses       TYPE string OPTIONAL
        !ip_auto          TYPE string OPTIONAL
        !ip_lblalgn       TYPE string OPTIONAL
        !ip_lbloffset     TYPE string OPTIONAL
        !ip_nomultilvllbl TYPE string OPTIONAL
        !ip_crossbetween  TYPE string OPTIONAL .
    METHODS set_show_legend_key
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_values
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_cat_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_ser_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_percent
      IMPORTING
        !ip_value TYPE c .
    METHODS set_varycolor
      IMPORTING
        !ip_value TYPE c .
  PROTECTED SECTION.
*"* protected components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
ENDCLASS.



CLASS zcl_excel_graph_bars IMPLEMENTATION.


  METHOD create_ax.
    DATA ls_ax TYPE s_ax.
    ls_ax-type = ip_type.

    IF ip_type = c_catax.
      IF ip_axid IS SUPPLIED.
        ls_ax-axid = ip_axid.
      ELSE.
        ls_ax-axid = '1'.
      ENDIF.
      IF ip_orientation IS SUPPLIED.
        ls_ax-orientation = ip_orientation.
      ELSE.
        ls_ax-orientation = 'minMax'.
      ENDIF.
      IF ip_delete IS SUPPLIED.
        ls_ax-delete = ip_delete.
      ELSE.
        ls_ax-delete = '0'.
      ENDIF.
      IF ip_axpos IS SUPPLIED.
        ls_ax-axpos = ip_axpos.
      ELSE.
        ls_ax-axpos = 'b'.
      ENDIF.
      IF ip_formatcode IS SUPPLIED.
        ls_ax-formatcode = ip_formatcode.
      ELSE.
        ls_ax-formatcode = 'General'.
      ENDIF.
      IF ip_sourcelinked IS SUPPLIED.
        ls_ax-sourcelinked = ip_sourcelinked.
      ELSE.
        ls_ax-sourcelinked = '1'.
      ENDIF.
      IF ip_majortickmark IS SUPPLIED.
        ls_ax-majortickmark = ip_majortickmark.
      ELSE.
        ls_ax-majortickmark = 'out'.
      ENDIF.
      IF ip_minortickmark IS SUPPLIED.
        ls_ax-minortickmark = ip_minortickmark.
      ELSE.
        ls_ax-minortickmark = 'none'.
      ENDIF.
      IF ip_ticklblpos IS SUPPLIED.
        ls_ax-ticklblpos = ip_ticklblpos.
      ELSE.
        ls_ax-ticklblpos = 'nextTo'.
      ENDIF.
      IF ip_crossax IS SUPPLIED.
        ls_ax-crossax = ip_crossax.
      ELSE.
        ls_ax-crossax = '2'.
      ENDIF.
      IF ip_crosses IS SUPPLIED.
        ls_ax-crosses = ip_crosses.
      ELSE.
        ls_ax-crosses = 'autoZero'.
      ENDIF.
      IF ip_auto IS SUPPLIED.
        ls_ax-auto = ip_auto.
      ELSE.
        ls_ax-auto = '1'.
      ENDIF.
      IF ip_lblalgn IS SUPPLIED.
        ls_ax-lblalgn = ip_lblalgn.
      ELSE.
        ls_ax-lblalgn = 'ctr'.
      ENDIF.
      IF ip_lbloffset IS SUPPLIED.
        ls_ax-lbloffset = ip_lbloffset.
      ELSE.
        ls_ax-lbloffset = '100'.
      ENDIF.
      IF ip_nomultilvllbl IS SUPPLIED.
        ls_ax-nomultilvllbl = ip_nomultilvllbl.
      ELSE.
        ls_ax-nomultilvllbl = '0'.
      ENDIF.
    ELSEIF ip_type = c_valax.
      IF ip_axid IS SUPPLIED.
        ls_ax-axid = ip_axid.
      ELSE.
        ls_ax-axid = '2'.
      ENDIF.
      IF ip_orientation IS SUPPLIED.
        ls_ax-orientation = ip_orientation.
      ELSE.
        ls_ax-orientation = 'minMax'.
      ENDIF.
      IF ip_delete IS SUPPLIED.
        ls_ax-delete = ip_delete.
      ELSE.
        ls_ax-delete = '0'.
      ENDIF.
      IF ip_axpos IS SUPPLIED.
        ls_ax-axpos = ip_axpos.
      ELSE.
        ls_ax-axpos = 'l'.
      ENDIF.
      IF ip_formatcode IS SUPPLIED.
        ls_ax-formatcode = ip_formatcode.
      ELSE.
        ls_ax-formatcode = 'General'.
      ENDIF.
      IF ip_sourcelinked IS SUPPLIED.
        ls_ax-sourcelinked = ip_sourcelinked.
      ELSE.
        ls_ax-sourcelinked = '1'.
      ENDIF.
      IF ip_majortickmark IS SUPPLIED.
        ls_ax-majortickmark = ip_majortickmark.
      ELSE.
        ls_ax-majortickmark = 'out'.
      ENDIF.
      IF ip_minortickmark IS SUPPLIED.
        ls_ax-minortickmark = ip_minortickmark.
      ELSE.
        ls_ax-minortickmark = 'none'.
      ENDIF.
      IF ip_ticklblpos IS SUPPLIED.
        ls_ax-ticklblpos = ip_ticklblpos.
      ELSE.
        ls_ax-ticklblpos = 'nextTo'.
      ENDIF.
      IF ip_crossax IS SUPPLIED.
        ls_ax-crossax = ip_crossax.
      ELSE.
        ls_ax-crossax = '1'.
      ENDIF.
      IF ip_crosses IS SUPPLIED.
        ls_ax-crosses = ip_crosses.
      ELSE.
        ls_ax-crosses = 'autoZero'.
      ENDIF.
      IF ip_crossbetween IS SUPPLIED.
        ls_ax-crossbetween = ip_crossbetween.
      ELSE.
        ls_ax-crossbetween = 'between'.
      ENDIF.
    ENDIF.

    APPEND ls_ax TO me->axes.
    SORT me->axes BY axid ASCENDING.
  ENDMETHOD.


  METHOD set_show_cat_name.
    ns_showcatnameval = ip_value.
  ENDMETHOD.


  METHOD set_show_legend_key.
    ns_showlegendkeyval = ip_value.
  ENDMETHOD.


  METHOD set_show_percent.
    ns_showpercentval = ip_value.
  ENDMETHOD.


  METHOD set_show_ser_name.
    ns_showsernameval = ip_value.
  ENDMETHOD.


  METHOD set_show_values.
    ns_showvalval = ip_value.
  ENDMETHOD.


  METHOD set_varycolor.
    ns_varycolorsval = ip_value.
  ENDMETHOD.
ENDCLASS.
