class ZCL_EXCEL_COMMENT definition
  public
  final
  create public .

public section.

  methods CONSTRUCTOR .
  methods GET_INDEX
    returning
      value(RP_INDEX) type STRING .
  methods GET_NAME
    returning
      value(GET_REF) type STRING .
  methods GET_TEXT
    returning
      value(RP_TEXT) type STRING .
  methods SET_TEXT
    importing
      !IP_TEXT type STRING
      !IP_REF type STRING optional .
  methods GET_REF
    returning
      value(RP_REF) type STRING .
  methods SET_CONTENT
    importing
      !IP_CONTENT type XSTRING .
  methods GET_CONTENT
    returning
      value(RP_CONTENT) type XSTRING .
  methods SET_FILENAME
    importing
      !IP_FILENAME type STRING .
  methods GET_FILENAME
    returning
      value(RP_FILENAME) type STRING .
  methods SET_REF
    importing
      !IP_REF type STRING .
protected section.
private section.

  data INDEX type STRING .
  data REF type STRING .
  data TEXT type STRING .
  data CONTENT type XSTRING .
  data FILENAME type STRING .
ENDCLASS.



CLASS ZCL_EXCEL_COMMENT IMPLEMENTATION.


  method CONSTRUCTOR.
  endmethod.


  method GET_CONTENT.
    rp_content = me->content.
  endmethod.


  method GET_FILENAME.
    rp_filename = me->filename.
  endmethod.


  method GET_INDEX.
      rp_index = me->index.
  endmethod.


  method GET_NAME.
  endmethod.


  method GET_REF.
     rp_ref = me->ref.
  endmethod.


    method GET_TEXT.

      rp_text = me->text.
  endmethod.


  method SET_CONTENT.
    me->content = ip_content.
  endmethod.


  method SET_FILENAME.
    me->filename = ip_filename.
  endmethod.


  method SET_REF.
    me->ref = ip_ref.
  endmethod.


  method SET_TEXT.
    me->text = ip_text.
     IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.
  endmethod.
ENDCLASS.
