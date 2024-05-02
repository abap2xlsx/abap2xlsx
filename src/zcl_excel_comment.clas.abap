CLASS zcl_excel_comment DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

  CONSTANTS default_width TYPE i VALUE 2. "#EC NOTEXT
  CONSTANTS default_height TYPE i VALUE 15. "#EC NOTEXT

    METHODS constructor .
    METHODS get_name
      RETURNING
        VALUE(r_name) TYPE string .
    METHODS get_index
      RETURNING
        VALUE(rp_index) TYPE string .
    METHODS get_ref
      RETURNING
        VALUE(rp_ref) TYPE string .
    METHODS get_text
      RETURNING
        VALUE(rp_text) TYPE string .
    METHODS set_text
      IMPORTING
        !ip_text TYPE string
        !ip_ref  TYPE string OPTIONAL .
  methods SET_SIZE
    importing
      !IP_WIDTH type I optional
      !IP_HEIGHT type I optional .
  methods GET_SIZE
    exporting
      !EP_WIDTH type I
      !EP_HEIGHT type I .

  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA index TYPE string .
    DATA ref TYPE string .
    DATA text TYPE string .
  data WIDTH type I .
  data HEIGHT type I .
ENDCLASS.



CLASS zcl_excel_comment IMPLEMENTATION.


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


  METHOD get_text.
    rp_text = me->text.
  ENDMETHOD.


  METHOD set_text.
    me->text = ip_text.

    IF ip_ref IS SUPPLIED.
      me->ref = ip_ref.
    ENDIF.
  ENDMETHOD.

METHOD set_size.

  IF ip_width IS SUPPLIED.
    width = ip_width.
  ENDIF.

  IF ip_height IS SUPPLIED.
    height = ip_height.
  ENDIF.

ENDMETHOD.

method GET_SIZE.

  IF width IS NOT INITIAL.
    ep_width = width.
  ELSE.
    ep_width = default_width. "Default width
  ENDIF.

  IF height IS NOT INITIAL.
    ep_height = height.
  ELSE.
    ep_height = default_height. "Default height
  ENDIF.

endmethod.

ENDCLASS.
