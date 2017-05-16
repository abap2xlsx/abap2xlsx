class ZCL_EXCEL_GRAPH_PIE definition
  public
  inheriting from ZCL_EXCEL_GRAPH
  final
  create public .

public section.

*"* public components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
  data NS_LEGENDPOSVAL type STRING value 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_OVERLAYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_PPRRTL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_ENDPARARPRLANG type STRING value 'it-IT'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_VARYCOLORSVAL type STRING value '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_FIRSTSLICEANGVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWLEGENDKEYVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWVALVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWCATNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWSERNAMEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWPERCENTVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWBUBBLESIZEVAL type STRING value '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data NS_SHOWLEADERLINESVAL type STRING value '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

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
  methods SET_SHOW_LEADER_LINES
    importing
      !IP_VALUE type C .
  methods SET_VARYCOLOR
    importing
      !IP_VALUE type C .
protected section.
*"* protected components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
private section.
*"* private components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_GRAPH_PIE IMPLEMENTATION.


method SET_SHOW_CAT_NAME.
  ns_showcatnameval = ip_value.
  endmethod.


method SET_SHOW_LEADER_LINES.
  ns_showleaderlinesval = ip_value.
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
