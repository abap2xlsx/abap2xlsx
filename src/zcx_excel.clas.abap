CLASS zcx_excel DEFINITION
  PUBLIC
  INHERITING FROM cx_static_check
  CREATE PUBLIC .

*"* public components of class ZCX_EXCEL
*"* do not include other source files here!!!
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
  PUBLIC SECTION.

    CONSTANTS zcx_excel TYPE sotr_conc VALUE '028C0ED2B5601ED78EB6F3368B1E4F9B' ##NO_TEXT.
    DATA error TYPE string .
    DATA syst_at_raise TYPE syst .

    METHODS constructor
      IMPORTING
        !textid        LIKE textid OPTIONAL
        !previous      LIKE previous OPTIONAL
        !error         TYPE string OPTIONAL
        !syst_at_raise TYPE syst OPTIONAL .
    CLASS-METHODS raise_text
      IMPORTING
        !iv_text TYPE clike
      RAISING
        zcx_excel .
    CLASS-METHODS raise_symsg
      RAISING
        zcx_excel .

    METHODS if_message~get_longtext
        REDEFINITION .
    METHODS if_message~get_text
        REDEFINITION .
  PROTECTED SECTION.
*"* private components of class ZCX_EXCEL
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcx_excel IMPLEMENTATION.


  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    CALL METHOD super->constructor
      EXPORTING
        textid   = textid
        previous = previous.
    IF textid IS INITIAL.
      me->textid = zcx_excel .
    ENDIF.
    me->error = error .
    me->syst_at_raise = syst_at_raise .
  ENDMETHOD.


  METHOD if_message~get_longtext.

    IF   me->error         IS NOT INITIAL
      OR me->syst_at_raise IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied explicitly use this as longtext as well
*--------------------------------------------------------------------*
      result = me->get_text( ).
    ELSE.
*--------------------------------------------------------------------*
* otherwise use standard method to derive text
*--------------------------------------------------------------------*
      result = super->if_message~get_longtext( preserve_newlines = preserve_newlines ).
    ENDIF.
  ENDMETHOD.


  METHOD if_message~get_text.

    IF me->error IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied explicitly use this
*--------------------------------------------------------------------*
      result = me->error .
    ELSEIF me->syst_at_raise IS NOT INITIAL.
*--------------------------------------------------------------------*
* If message was supplied by syst create messagetext now
*--------------------------------------------------------------------*
      MESSAGE ID syst_at_raise-msgid TYPE syst_at_raise-msgty NUMBER syst_at_raise-msgno
           WITH  syst_at_raise-msgv1 syst_at_raise-msgv2 syst_at_raise-msgv3 syst_at_raise-msgv4
           INTO  result.
    ELSE.
*--------------------------------------------------------------------*
* otherwise use standard method to derive text
*--------------------------------------------------------------------*
      CALL METHOD super->if_message~get_text
        RECEIVING
          result = result.
    ENDIF.
  ENDMETHOD.


  METHOD raise_symsg.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        syst_at_raise = syst.
  ENDMETHOD.


  METHOD raise_text.
    RAISE EXCEPTION TYPE zcx_excel
      EXPORTING
        error = iv_text.
  ENDMETHOD.
ENDCLASS.
