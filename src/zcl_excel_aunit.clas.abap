class zcl_excel_aunit definition
  public
  final
  create private.

  public section.

    class-methods class_constructor.

    class-methods assert_differs
      importing
        !exp                    type simple
        !act                    type simple
        !msg                    type csequence optional
        !level                  type aunit_level default if_aunit_constants=>critical
        !tol                    type f optional
        !quit                   type aunit_flowctrl default if_aunit_constants=>method
      returning
        value(assertion_failed) type abap_bool.

    class-methods assert_equals
      importing
        !exp                    type any
        !act                    type any
        !msg                    type csequence optional
        !level                  type aunit_level default if_aunit_constants=>critical
        !tol                    type f optional
        !quit                   type aunit_flowctrl default if_aunit_constants=>method
        !ignore_hash_sequence   type abap_bool default abap_false
      returning
        value(assertion_failed) type abap_bool.

    class-methods fail
      importing
        msg    type csequence optional
        level  type aunit_level default if_aunit_constants=>critical
        quit   type aunit_flowctrl default if_aunit_constants=>method
        detail type csequence optional.

  protected section.
  private section.
    types tv_clsname type seoclass-clsname.

    constants:
      begin of en_clsname,
        new  type tv_clsname value 'CL_ABAP_UNIT_ASSERT',
        old  type tv_clsname value 'CL_AUNIT_ASSERT',
        none type tv_clsname value '',
      end of en_clsname.

    class-data clsname type tv_clsname.
endclass.



class zcl_excel_aunit implementation.

  method class_constructor.
    " Let see >=7.02
    select single clsname into clsname
      from seoclass
      where clsname = en_clsname-new.

    check sy-subrc <> 0.

    " Let see >=7.00 or even lower
    select single clsname into clsname
      from seoclass
      where clsname = en_clsname-old.

    check sy-subrc <> 0.

    " We do nothing for now not supported

  endmethod.

  method assert_differs.
    check clsname = en_clsname-new or clsname = en_clsname-old.

    call method (clsname)=>assert_differs
      exporting
        exp              = exp
        act              = act
        msg              = msg
        level            = level
        tol              = tol
        quit             = quit
      receiving
        assertion_failed = assertion_failed.

  endmethod.

  method assert_equals.
    check clsname = en_clsname-new or clsname = en_clsname-old.

    call method (clsname)=>assert_equals
      exporting
        exp                  = exp
        act                  = act
        msg                  = msg
        level                = level
        tol                  = tol
        quit                 = quit
        ignore_hash_sequence = ignore_hash_sequence
      receiving
        assertion_failed     = assertion_failed.
  endmethod.

  method fail.
    check clsname = en_clsname-new or clsname = en_clsname-old.

    call method (clsname)=>fail
      exporting
        msg    = msg
        level  = level
        quit   = quit
        detail = detail.

  endmethod.
endclass.
