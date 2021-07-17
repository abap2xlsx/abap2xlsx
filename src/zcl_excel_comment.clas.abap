class ZCL_EXCEL_COMMENT definition
  public
  final
  create public .

public section.
  type-pools ABAP .

  methods CONSTRUCTOR .
  methods GET_NAME
    returning
      value(R_NAME) type STRING .
  methods GET_INDEX
    returning
      value(RP_INDEX) type STRING .
  methods GET_REF
    returning
      value(RP_REF) type STRING .
  methods GET_TEXT
    returning
      value(RP_TEXT) type STRING .
  methods SET_TEXT
    importing
      !IP_TEXT type STRING
      !IP_REF type STRING optional .
protected section.
private section.

  data INDEX type STRING .
  data REF type STRING .
  data TEXT type STRING .
ENDCLASS.



CLASS ZCL_EXCEL_COMMENT IMPLEMENTATION.


METHOD constructor.

  ENDMETHOD.


METHOD get_index.
    rp_index = me->index.
  ENDMETHOD.


METHOD get_name.

  ENDMETHOD.


METHOD get_ref.
  rp_ref = me->ref.
ENDMETHOD.


method GET_TEXT.
  rp_text = me->text.
endmethod.


METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
