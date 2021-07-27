CLASS zcl_excel_aunit DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.

    CLASS-METHODS class_constructor.

    CLASS-METHODS assert_differs
      IMPORTING
        !exp                    TYPE simple
        !act                    TYPE simple
        !msg                    TYPE csequence OPTIONAL
        !level                  TYPE aunit_level DEFAULT if_aunit_constants=>critical
        !tol                    TYPE f OPTIONAL
        !quit                   TYPE aunit_flowctrl DEFAULT if_aunit_constants=>method
      RETURNING
        VALUE(assertion_failed) TYPE abap_bool.

    CLASS-METHODS assert_equals
      IMPORTING
        !exp                    TYPE any
        !act                    TYPE any
        !msg                    TYPE csequence OPTIONAL
        !level                  TYPE aunit_level DEFAULT if_aunit_constants=>critical
        !tol                    TYPE f OPTIONAL
        !quit                   TYPE aunit_flowctrl DEFAULT if_aunit_constants=>method
        !ignore_hash_sequence   TYPE abap_bool DEFAULT abap_false
      RETURNING
        VALUE(assertion_failed) TYPE abap_bool.

    CLASS-METHODS fail
      IMPORTING
        msg    TYPE csequence OPTIONAL
        level  TYPE aunit_level DEFAULT if_aunit_constants=>critical
        quit   TYPE aunit_flowctrl DEFAULT if_aunit_constants=>method
        detail TYPE csequence OPTIONAL.

  PROTECTED SECTION.
  PRIVATE SECTION.
    TYPES tv_clsname TYPE seoclass-clsname.

    CONSTANTS:
      BEGIN OF en_clsname,
        new  TYPE tv_clsname VALUE 'CL_ABAP_UNIT_ASSERT',
        old  TYPE tv_clsname VALUE 'CL_AUNIT_ASSERT',
        none TYPE tv_clsname VALUE '',
      END OF en_clsname.

    CLASS-DATA clsname TYPE tv_clsname.
ENDCLASS.



CLASS zcl_excel_aunit IMPLEMENTATION.

  METHOD class_constructor.
    " Let see >=7.02
    SELECT SINGLE clsname INTO clsname
      FROM seoclass
      WHERE clsname = en_clsname-new.

    CHECK sy-subrc <> 0.

    " Let see >=7.00 or even lower
    SELECT SINGLE clsname INTO clsname
      FROM seoclass
      WHERE clsname = en_clsname-old.

    CHECK sy-subrc <> 0.

    " We do nothing for now not supported

  ENDMETHOD.

  METHOD assert_differs.
    CHECK clsname = en_clsname-new OR clsname = en_clsname-old.

    CALL METHOD (clsname)=>assert_differs
      EXPORTING
        exp              = exp
        act              = act
        msg              = msg
        level            = level
        tol              = tol
        quit             = quit
      RECEIVING
        assertion_failed = assertion_failed.

  ENDMETHOD.

  METHOD assert_equals.
    CHECK clsname = en_clsname-new OR clsname = en_clsname-old.

    CALL METHOD (clsname)=>assert_equals
      EXPORTING
        exp                  = exp
        act                  = act
        msg                  = msg
        level                = level
        tol                  = tol
        quit                 = quit
        ignore_hash_sequence = ignore_hash_sequence
      RECEIVING
        assertion_failed     = assertion_failed.
  ENDMETHOD.

  METHOD fail.
    CHECK clsname = en_clsname-new OR clsname = en_clsname-old.

    CALL METHOD (clsname)=>fail
      EXPORTING
        msg    = msg
        level  = level
        quit   = quit
        detail = detail.

  ENDMETHOD.
ENDCLASS.
