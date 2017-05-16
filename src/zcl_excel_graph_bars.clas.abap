class ZCL_EXCEL_GRAPH_BARS definition
  public
  inheriting from ZCL_EXCEL_GRAPH
  final
  create public .

public section.

  types:
*"* public components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
    BEGIN OF s_ax,
                axid TYPE string,
                type TYPE char5,
                orientation TYPE string,
                delete TYPE string,
                axpos TYPE string,
                formatcode TYPE string,
                sourcelinked TYPE string,
                majortickmark TYPE string,
                minortickmark TYPE string,
                ticklblpos TYPE string,
                crossax TYPE string,
                crosses TYPE string,
                auto TYPE string,
                lblalgn TYPE string,
                lbloffset TYPE string,
                nomultilvllbl TYPE string,
                crossbetween TYPE string,
             END OF s_ax .
  types:
    t_ax TYPE STANDARD TABLE OF s_ax .

  data NS_BARDIRVAL type STRING value 'col'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  constants C_GROUPINGVAL_CLUSTERED type STRING value 'clustered'. "#EC NOTEXT
  constants C_GROUPINGVAL_STACKED type STRING value 'stacked'. "#EC NOTEXT
  data NS_GROUPINGVAL type STRING value C_GROUPINGVAL_CLUSTERED. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_VARYCOLORSVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWLEGENDKEYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWVALVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWCATNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWSERNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWPERCENTVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWBUBBLESIZEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_GAPWIDTHVAL type STRING value '150'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data AXES type T_AX .
  constants:
    C_VALAX type c length 5 value 'VALAX'. "#EC NOTEXT
  constants:
    C_CATAX type c length 5 value 'CATAX'. "#EC NOTEXT
  data NS_LEGENDPOSVAL type STRING value 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  data NS_OVERLAYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
  constants C_INVERTIFNEGATIVE_YES type STRING value '1'. "#EC NOTEXT
  constants C_INVERTIFNEGATIVE_NO type STRING value '0'. "#EC NOTEXT

  methods CREATE_AX
    importing
      !IP_AXID type STRING optional
      !IP_TYPE type CHAR5
      !IP_ORIENTATION type STRING optional
      !IP_DELETE type STRING optional
      !IP_AXPOS type STRING optional
      !IP_FORMATCODE type STRING optional
      !IP_SOURCELINKED type STRING optional
      !IP_MAJORTICKMARK type STRING optional
      !IP_MINORTICKMARK type STRING optional
      !IP_TICKLBLPOS type STRING optional
      !IP_CROSSAX type STRING optional
      !IP_CROSSES type STRING optional
      !IP_AUTO type STRING optional
      !IP_LBLALGN type STRING optional
      !IP_LBLOFFSET type STRING optional
      !IP_NOMULTILVLLBL type STRING optional
      !IP_CROSSBETWEEN type STRING optional .
  methods SET_SHOW_LEGEND_KEY
    importing
      !IP_VALUE type C .
  methods SET_SHOW_VALUES
    importing
      !IP_VALUE type C .
  methods SET_SHOW_CAT_NAME
    importing
      !IP_VALUE type C .
  methods SET_SHOW_SER_NAME
    importing
      !IP_VALUE type C .
  methods SET_SHOW_PERCENT
    importing
      !IP_VALUE type C .
  methods SET_VARYCOLOR
    importing
      !IP_VALUE type C .
protected section.
*"* protected components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
private section.
*"* private components of class ZCL_EXCEL_GRAPH_BARS
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_GRAPH_BARS IMPLEMENTATION.


method CREATE_AX.
  DATA ls_ax TYPE s_ax.
  ls_ax-type = ip_type.

  if ip_type = c_catax.
    if ip_axid is SUPPLIED.
      ls_ax-axid = ip_axid.
    else.
      ls_ax-axid = '1'.
    endif.
    if ip_orientation is SUPPLIED.
      ls_ax-orientation = ip_orientation.
    else.
      ls_ax-orientation = 'minMax'.
    endif.
    if ip_delete is SUPPLIED.
      ls_ax-delete = ip_delete.
    else.
      ls_ax-delete = '0'.
    endif.
    if ip_axpos is SUPPLIED.
      ls_ax-axpos = ip_axpos.
    else.
      ls_ax-axpos = 'b'.
    endif.
    if ip_formatcode is SUPPLIED.
      ls_ax-formatcode = ip_formatcode.
    else.
      ls_ax-formatcode = 'General'.
    endif.
    if ip_sourcelinked is SUPPLIED.
      ls_ax-sourcelinked = ip_sourcelinked.
    else.
      ls_ax-sourcelinked = '1'.
    endif.
    if ip_majorTickMark is SUPPLIED.
      ls_ax-majorTickMark = ip_majorTickMark.
    else.
      ls_ax-majorTickMark = 'out'.
    endif.
    if ip_minorTickMark is SUPPLIED.
      ls_ax-minorTickMark = ip_minorTickMark.
    else.
      ls_ax-minorTickMark = 'none'.
    endif.
    if ip_ticklblpos is SUPPLIED.
      ls_ax-ticklblpos = ip_ticklblpos.
    else.
      ls_ax-ticklblpos = 'nextTo'.
    endif.
    if ip_crossax is SUPPLIED.
      ls_ax-crossax = ip_crossax.
    else.
      ls_ax-crossax = '2'.
    endif.
    if ip_crosses is SUPPLIED.
      ls_ax-crosses = ip_crosses.
    else.
      ls_ax-crosses = 'autoZero'.
    endif.
    if ip_auto is SUPPLIED.
      ls_ax-auto = ip_auto.
    else.
      ls_ax-auto = '1'.
    endif.
    if ip_lblAlgn is SUPPLIED.
      ls_ax-lblAlgn = ip_lblAlgn.
    else.
      ls_ax-lblAlgn = 'ctr'.
    endif.
    if ip_lblOffset is SUPPLIED.
      ls_ax-lblOffset = ip_lblOffset.
    else.
      ls_ax-lblOffset = '100'.
    endif.
    if ip_noMultiLvlLbl is SUPPLIED.
      ls_ax-noMultiLvlLbl = ip_noMultiLvlLbl.
    else.
      ls_ax-noMultiLvlLbl = '0'.
    endif.
  elseif ip_type = c_valax.
    if ip_axid is SUPPLIED.
      ls_ax-axid = ip_axid.
    else.
      ls_ax-axid = '2'.
    endif.
    if ip_orientation is SUPPLIED.
      ls_ax-orientation = ip_orientation.
    else.
      ls_ax-orientation = 'minMax'.
    endif.
    if ip_delete is SUPPLIED.
      ls_ax-delete = ip_delete.
    else.
      ls_ax-delete = '0'.
    endif.
    if ip_axpos is SUPPLIED.
      ls_ax-axpos = ip_axpos.
    else.
      ls_ax-axpos = 'l'.
    endif.
    if ip_formatcode is SUPPLIED.
      ls_ax-formatcode = ip_formatcode.
    else.
      ls_ax-formatcode = 'General'.
    endif.
    if ip_sourcelinked is SUPPLIED.
      ls_ax-sourcelinked = ip_sourcelinked.
    else.
      ls_ax-sourcelinked = '1'.
    endif.
    if ip_majorTickMark is SUPPLIED.
      ls_ax-majorTickMark = ip_majorTickMark.
    else.
      ls_ax-majorTickMark = 'out'.
    endif.
    if ip_minorTickMark is SUPPLIED.
      ls_ax-minorTickMark = ip_minorTickMark.
    else.
      ls_ax-minorTickMark = 'none'.
    endif.
    if ip_ticklblpos is SUPPLIED.
      ls_ax-ticklblpos = ip_ticklblpos.
    else.
      ls_ax-ticklblpos = 'nextTo'.
    endif.
    if ip_crossax is SUPPLIED.
      ls_ax-crossax = ip_crossax.
    else.
      ls_ax-crossax = '1'.
    endif.
    if ip_crosses is SUPPLIED.
      ls_ax-crosses = ip_crosses.
    else.
      ls_ax-crosses = 'autoZero'.
    endif.
    if ip_crossBetween is SUPPLIED.
      ls_ax-crossBetween = ip_crossBetween.
    else.
      ls_ax-crossBetween = 'between'.
    endif.
  endif.

  APPEND ls_ax TO me->axes.
  sort me->axes by axid ascending.
  endmethod.


method SET_SHOW_CAT_NAME.
    ns_showcatnameval = ip_value.
  endmethod.


method SET_SHOW_LEGEND_KEY.
  ns_showlegendkeyval = ip_value.
  endmethod.


method SET_SHOW_PERCENT.
  ns_showpercentval = ip_value.
  endmethod.


method SET_SHOW_SER_NAME.
  ns_showsernameval = ip_value.
  endmethod.


method SET_SHOW_VALUES.
  ns_showvalval = ip_value.
  endmethod.


method SET_VARYCOLOR.
  ns_varycolorsval = ip_value.
  endmethod.
ENDCLASS.
