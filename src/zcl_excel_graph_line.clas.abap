class ZCL_EXCEL_GRAPH_LINE definition
  public
  inheriting from ZCL_EXCEL_GRAPH
  final
  create public .

public section.

  types:
*"* public components of class ZCL_EXCEL_GRAPH_LINE
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

  data NS_GROUPINGVAL type STRING value 'standard'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_VARYCOLORSVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWLEGENDKEYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWVALVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWCATNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWSERNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWPERCENTVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWBUBBLESIZEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_MARKERVAL type STRING value '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SMOOTHVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data AXES type T_AX .
  constants:
    C_VALAX type c length 5 value 'VALAX'. "#EC NOTEXT
  constants:
    C_CATAX type c length 5 value 'CATAX'. "#EC NOTEXT
  data NS_LEGENDPOSVAL type STRING value 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_OVERLAYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  constants C_SYMBOL_AUTO type STRING value 'auto'. "#EC NOTEXT
  constants C_SYMBOL_NONE type STRING value 'none'. "#EC NOTEXT

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
*"* protected components of class ZCL_EXCEL_GRAPH_LINE
*"* do not include other source files here!!!
private section.
*"* private components of class ZCL_EXCEL_GRAPH_LINE
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_GRAPH_LINE IMPLEMENTATION.


method CREATE_AX.
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
