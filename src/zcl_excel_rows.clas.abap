*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
CLASS ZCL_EXCEL_ROWS DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .


*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.
    TYPES:
      BEGIN OF T_ROWS,
        ROW_INDEX TYPE INT4,
        ROW       TYPE REF TO ZCL_EXCEL_ROW,
      END OF T_ROWS.

    DATA:
          DT_ROWS TYPE HASHED TABLE OF T_ROWS WITH UNIQUE KEY ROW_INDEX.

    METHODS ADD
      IMPORTING
        !IO_ROW TYPE REF TO ZCL_EXCEL_ROW .
    METHODS CLEAR .
    METHODS CONSTRUCTOR .
    METHODS GET
      IMPORTING
        !IP_INDEX     TYPE I
      RETURNING
        VALUE(EO_ROW) TYPE REF TO ZCL_EXCEL_ROW .
    METHODS GET_ITERATOR
      RETURNING
        VALUE(EO_ITERATOR) TYPE REF TO CL_OBJECT_COLLECTION_ITERATOR .
    METHODS IS_EMPTY
      RETURNING
        VALUE(IS_EMPTY) TYPE FLAG .
    METHODS REMOVE
      IMPORTING
        !IO_ROW TYPE REF TO ZCL_EXCEL_ROW .
    METHODS SIZE
      RETURNING
        VALUE(EP_SIZE) TYPE I .
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA ROWS TYPE REF TO CL_OBJECT_COLLECTION .
ENDCLASS.



CLASS ZCL_EXCEL_ROWS IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->ADD
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_ROW                         TYPE REF TO ZCL_EXCEL_ROW
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD ADD.
*    IF SY-UNAME = 'BILEN'.
    DATA:
          LS_ROW TYPE T_ROWS.
    LS_ROW-ROW_INDEX = IO_ROW->GET_ROW_INDEX( ).
    LS_ROW-ROW = IO_ROW.
    INSERT LS_ROW INTO TABLE DT_ROWS.
*
*    ELSE.
*      ROWS->ADD( IO_ROW ).
*
*    ENDIF.

  ENDMETHOD.                    "ADD


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->CLEAR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD CLEAR.
*    ROWS->CLEAR( ).
    CLEAR DT_ROWS[].

  ENDMETHOD.                    "CLEAR


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD CONSTRUCTOR.

    CREATE OBJECT ROWS.

  ENDMETHOD.                    "CONSTRUCTOR


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->GET
* +-------------------------------------------------------------------------------------------------+
* | [--->] IP_INDEX                       TYPE        I
* | [<-()] EO_ROW                         TYPE REF TO ZCL_EXCEL_ROW
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD GET.
*    IF SY-UNAME = 'BILEN'.
    READ TABLE DT_ROWS ASSIGNING FIELD-SYMBOL(<LS_ROW>) WITH TABLE KEY ROW_INDEX = IP_INDEX.
    IF SY-SUBRC = 0.
      EO_ROW ?= <LS_ROW>-ROW.
    ENDIF.
*    ELSE.
*      EO_ROW ?= ROWS->IF_OBJECT_COLLECTION~GET( IP_INDEX ).
*
*    ENDIF.
  ENDMETHOD.                    "GET


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->GET_ITERATOR
* +-------------------------------------------------------------------------------------------------+
* | [<-()] EO_ITERATOR                    TYPE REF TO CL_OBJECT_COLLECTION_ITERATOR
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD GET_ITERATOR.
    EO_ITERATOR ?= ROWS->IF_OBJECT_COLLECTION~GET_ITERATOR( ).
  ENDMETHOD.                    "GET_ITERATOR


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->IS_EMPTY
* +-------------------------------------------------------------------------------------------------+
* | [<-()] IS_EMPTY                       TYPE        FLAG
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD IS_EMPTY.
    DATA(LV_LINES) = LINES( DT_ROWS ).

    IF LV_LINES = 0.
      IS_EMPTY = 'X'.
    ELSE.
      IS_EMPTY = ' '.
    ENDIF.
*    IS_EMPTY = ROWS->IF_OBJECT_COLLECTION~IS_EMPTY( ).
  ENDMETHOD.                    "IS_EMPTY


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->REMOVE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_ROW                         TYPE REF TO ZCL_EXCEL_ROW
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD REMOVE.
*    ROWS->REMOVE( IO_ROW ).
    DELETE DT_ROWS WHERE ROW_INDEX = IO_ROW->GET_ROW_INDEX( ).
  ENDMETHOD.                    "REMOVE


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_ROWS->SIZE
* +-------------------------------------------------------------------------------------------------+
* | [<-()] EP_SIZE                        TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD SIZE.
*    IF sy-uname = 'BILEN'.
    EP_SIZE = LINES( DT_ROWS )..
*
*    else.
*    EP_SIZE = ROWS->IF_OBJECT_COLLECTION~SIZE( ).
*
*    ENDIF.
  ENDMETHOD.                    "SIZE
ENDCLASS.
