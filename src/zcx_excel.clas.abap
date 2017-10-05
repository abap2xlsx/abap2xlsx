class ZCX_EXCEL definition
  public
  inheriting from CX_STATIC_CHECK
  create public .

*"* public components of class ZCX_EXCEL
*"* do not include other source files here!!!
public section.

  constants ZCX_EXCEL type SOTR_CONC value '028C0ED2B5601ED78EB6F3368B1E4F9B'. "#EC NOTEXT
  data ERROR type STRING .
  data SYST_AT_RAISE type SYST .

  methods CONSTRUCTOR
    importing
      !TEXTID like TEXTID optional
      !PREVIOUS like PREVIOUS optional
      !ERROR type STRING optional
      !SYST_AT_RAISE type SYST optional .

  methods IF_MESSAGE~GET_LONGTEXT
    redefinition .
  methods IF_MESSAGE~GET_TEXT
    redefinition .
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
*"* protected components of class ZCX_EXCEL
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCX_EXCEL
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCX_EXCEL IMPLEMENTATION.


method CONSTRUCTOR.
CALL METHOD SUPER->CONSTRUCTOR
EXPORTING
TEXTID = TEXTID
PREVIOUS = PREVIOUS
.
 IF textid IS INITIAL.
   me->textid = ZCX_EXCEL .
 ENDIF.
me->ERROR = ERROR .
me->SYST_AT_RAISE = SYST_AT_RAISE .
endmethod.


method IF_MESSAGE~GET_LONGTEXT.

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
    super->if_message~get_longtext( EXPORTING
                                      preserve_newlines = preserve_newlines
                                    RECEIVING
                                      result            = result ).
  ENDIF.
  endmethod.


method IF_MESSAGE~GET_TEXT.

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
  endmethod.
ENDCLASS.
