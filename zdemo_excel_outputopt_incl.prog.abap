*&---------------------------------------------------------------------*
*&  Include           ZDEMO_EXCEL_OUTPUTOPT_INCL
*&---------------------------------------------------------------------*
CLASS lcl_output DEFINITION CREATE PRIVATE.
  PUBLIC SECTION.
    CLASS-METHODS:
      output         IMPORTING cl_excel            TYPE REF TO zcl_excel
                               iv_writerclass_name TYPE clike OPTIONAL,
      f4_path        RETURNING VALUE(selected_folder) TYPE string,
      parametertexts.

  PRIVATE SECTION.
    METHODS:
      download_frontend,
      download_backend,
      display_online,
      send_email.

    DATA: xdata     TYPE xstring,             " Will be used for sending as email
          t_rawdata TYPE solix_tab,           " Will be used for downloading or open directly
          bytecount TYPE i.                   " Will be used for downloading or open directly
ENDCLASS.                    "lcl_output DEFINITION


SELECTION-SCREEN BEGIN OF BLOCK bl1 WITH FRAME TITLE txt_bl1.
PARAMETERS: rb_down RADIOBUTTON GROUP rb1 USER-COMMAND space.

PARAMETERS: rb_back RADIOBUTTON GROUP rb1.

PARAMETERS: rb_show RADIOBUTTON GROUP rb1 DEFAULT 'X' .

PARAMETERS: rb_send RADIOBUTTON GROUP rb1.

PARAMETERS: p_path  TYPE string LOWER CASE MODIF ID pat.
PARAMETERS: p_email TYPE string LOWER CASE MODIF ID ema OBLIGATORY DEFAULT 'insert_your_emailadress@here'.
PARAMETERS: p_backfn TYPE text40 NO-DISPLAY.
SELECTION-SCREEN END OF BLOCK bl1.


AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.

    IF rb_down IS INITIAL AND screen-group1 = 'PAT'.
      screen-input = 0.
      screen-invisible = 1.
    ENDIF.

    IF rb_send IS INITIAL AND screen-group1 = 'EMA'.
      screen-input = 0.
      screen-invisible = 1.
    ENDIF.

    MODIFY SCREEN.

  ENDLOOP.

INITIALIZATION.
  IF sy-batch IS INITIAL.
    cl_gui_frontend_services=>get_sapgui_workdir( CHANGING sapworkdir = p_path ).
    cl_gui_cfw=>flush( ).
  ENDIF.
  lcl_output=>parametertexts( ).  " If started in language w/o textelements translated set defaults
  sy-title = gc_save_file_name.
  txt_bl1 = 'Output options'(bl1).
  p_backfn = gc_save_file_name.  " Use as default if nothing else is supplied by submit

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path.
  p_path = lcl_output=>f4_path( ).


*----------------------------------------------------------------------*
*       CLASS lcl_output IMPLEMENTATION
*----------------------------------------------------------------------*
CLASS lcl_output IMPLEMENTATION.
  METHOD output.

    DATA: cl_output TYPE REF TO lcl_output,
          cl_writer TYPE REF TO zif_excel_writer.

    IF iv_writerclass_name IS INITIAL.
      CREATE OBJECT cl_output.
      CREATE OBJECT cl_writer TYPE zcl_excel_writer_2007.
    ELSE.
      CREATE OBJECT cl_output.
      CREATE OBJECT cl_writer TYPE (iv_writerclass_name).
    ENDIF.
    cl_output->xdata = cl_writer->write_file( cl_excel ).

* After 6.40 via cl_bcs_convert
    cl_output->t_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = cl_output->xdata ).
    cl_output->bytecount = xstrlen( cl_output->xdata ).

* before 6.40
*  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
*    EXPORTING
*      buffer        = cl_output->xdata
*    IMPORTING
*      output_length = cl_output->bytecount
*    TABLES
*      binary_tab    = cl_output->t_rawdata.

    CASE 'X'.
      WHEN rb_down.
        IF sy-batch IS INITIAL.
          cl_output->download_frontend( ).
        ELSE.
          MESSAGE e001(00) WITH 'Frontenddownload impossible in background processing'.
        ENDIF.

      WHEN rb_back.
        cl_output->download_backend( ).

      WHEN rb_show.
        IF sy-batch IS INITIAL.
          cl_output->display_online( ).
        ELSE.
          MESSAGE e001(00) WITH 'Online display absurd in background processing'.
        ENDIF.

      WHEN rb_send.
        cl_output->send_email( ).

    ENDCASE.
  ENDMETHOD.                    "output

  METHOD f4_path.
    DATA: new_path      TYPE string,
          repid         TYPE syrepid,
          dynnr         TYPE sydynnr,
          lt_dynpfields TYPE TABLE OF dynpread,
          ls_dynpfields LIKE LINE OF lt_dynpfields.

* Get current value
    dynnr = sy-dynnr.
    repid = sy-repid.
    ls_dynpfields-fieldname = 'P_PATH'.
    APPEND ls_dynpfields TO lt_dynpfields.

    CALL FUNCTION 'DYNP_VALUES_READ'
      EXPORTING
        dyname               = repid
        dynumb               = dynnr
      TABLES
        dynpfields           = lt_dynpfields
      EXCEPTIONS
        invalid_abapworkarea = 1
        invalid_dynprofield  = 2
        invalid_dynproname   = 3
        invalid_dynpronummer = 4
        invalid_request      = 5
        no_fielddescription  = 6
        invalid_parameter    = 7
        undefind_error       = 8
        double_conversion    = 9
        stepl_not_found      = 10
        OTHERS               = 11.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      EXIT.
    ENDIF.

    READ TABLE lt_dynpfields INTO ls_dynpfields INDEX 1.

    new_path = ls_dynpfields-fieldvalue.
    selected_folder = new_path.

    cl_gui_frontend_services=>directory_browse(
      EXPORTING
        window_title         = 'Select path to download EXCEL-file'
        initial_folder       = new_path
      CHANGING
        selected_folder      = new_path
      EXCEPTIONS
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        OTHERS               = 4
           ).
    cl_gui_cfw=>flush( ).
    CHECK new_path IS NOT INITIAL.
    selected_folder = new_path.

  ENDMETHOD.                                                "f4_path

  METHOD parametertexts.
* If started in language w/o textelements translated set defaults
* Furthermore I don't have to change the selectiontexts of all demoreports.
    DEFINE default_parametertext.
      if %_&1_%_app_%-text = '&1' or
         %_&1_%_app_%-text is initial.
        %_&1_%_app_%-text = &2.
      endif.
    END-OF-DEFINITION.

    default_parametertext:  rb_down  'Save to frontend',
                            rb_back  'Save to backend',
                            rb_show  'Direct display',
                            rb_send  'Send via email',

                            p_path   'Frontend-path to download to',
                            p_email  'Email to send xlsx to'.

  ENDMETHOD.                    "parametertexts

  METHOD: download_frontend.
    DATA: filename TYPE string.
* I don't like p_path here - but for this include it's ok
    filename = p_path.
* Add trailing "\" or "/"
    IF filename CA '/'.
      REPLACE REGEX '([^/])\s*$' IN filename WITH '$1/' .
    ELSE.
      REPLACE REGEX '([^\\])\s*$' IN filename WITH '$1\\'.
    ENDIF.

    CONCATENATE filename gc_save_file_name INTO filename.
* Get trailing blank
    cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = bytecount
                                                      filename     = filename
                                                      filetype     = 'BIN'
                                             CHANGING data_tab     = t_rawdata ).
  ENDMETHOD.                    "download_frontend

  METHOD download_backend.
    DATA: bytes_remain TYPE i.
    FIELD-SYMBOLS: <rawdata> LIKE LINE OF t_rawdata.

    OPEN DATASET p_backfn FOR OUTPUT IN BINARY MODE.
    CHECK sy-subrc = 0.

    bytes_remain = bytecount.

    LOOP AT t_rawdata ASSIGNING <rawdata>.

      AT LAST.
        CHECK bytes_remain >= 0.
        TRANSFER <rawdata> TO p_backfn LENGTH bytes_remain.
        EXIT.
      ENDAT.

      TRANSFER <rawdata> TO p_backfn.
      SUBTRACT 255 FROM bytes_remain.  " Solix has length 255

    ENDLOOP.

    CLOSE DATASET p_backfn.




    IF sy-repid <> sy-cprog AND sy-cprog IS NOT INITIAL.  " no need to display anything if download was selected and report was called for demo purposes
      LEAVE PROGRAM.
    ELSE.
      MESSAGE 'Data transferred to default backend directory' TYPE 'S'.
    ENDIF.
  ENDMETHOD.                    "download_backend

  METHOD display_online.
    DATA:error       TYPE REF TO i_oi_error,
         t_errors    TYPE STANDARD TABLE OF REF TO i_oi_error WITH NON-UNIQUE DEFAULT KEY,
         cl_control  TYPE REF TO i_oi_container_control, "OIContainerCtrl
         cl_document TYPE REF TO i_oi_document_proxy.   "Office Dokument

    c_oi_container_control_creator=>get_container_control( IMPORTING control = cl_control
                                                                     error   = error ).
    APPEND error TO t_errors.

    cl_control->init_control( EXPORTING  inplace_enabled     = 'X'
                                         no_flush            = 'X'
                                         r3_application_name = 'Demo Document Container'
                                         parent              = cl_gui_container=>screen0
                              IMPORTING  error               = error
                              EXCEPTIONS OTHERS              = 2 ).
    APPEND error TO t_errors.

    cl_control->get_document_proxy( EXPORTING document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    IMPORTING document_proxy = cl_document
                                              error          = error ).
    APPEND error TO t_errors.
* Errorhandling should be inserted here

    cl_document->open_document_from_table( EXPORTING document_size    = bytecount
                                                     document_table   = t_rawdata
                                                     open_inplace     = 'X' ).

    WRITE: '.'.  " To create an output.  That way screen0 will exist
  ENDMETHOD.                    "display_online

  METHOD send_email.
* Needed to send emails
    DATA: bcs_exception        TYPE REF TO cx_bcs,
          errortext            TYPE string,
          cl_send_request      TYPE REF TO cl_bcs,
          cl_document          TYPE REF TO cl_document_bcs,
          cl_recipient         TYPE REF TO if_recipient_bcs,
          cl_sender            TYPE REF TO cl_cam_address_bcs,
          t_attachment_header  TYPE soli_tab,
          wa_attachment_header LIKE LINE OF t_attachment_header,
          attachment_subject   TYPE sood-objdes,

          sood_bytecount       TYPE sood-objlen,
          mail_title           TYPE so_obj_des,
          t_mailtext           TYPE soli_tab,
          wa_mailtext          LIKE LINE OF t_mailtext,
          send_to              TYPE adr6-smtp_addr,
          sent                 TYPE os_boolean.


    mail_title     = 'Mail title'.
    wa_mailtext    = 'Mailtext'.
    APPEND wa_mailtext TO t_mailtext.

    TRY.
* Create send request
        cl_send_request = cl_bcs=>create_persistent( ).
* Create new document with mailtitle and mailtextg
        cl_document = cl_document_bcs=>create_document( i_type    = 'RAW' "#EC NOTEXT
                                                        i_text    = t_mailtext
                                                        i_subject = mail_title ).
* Add attachment to document
* since the new excelfiles have an 4-character extension .xlsx but the attachment-type only holds 3 charactes .xls,
* we have to specify the real filename via attachment header
* Use attachment_type xls to have SAP display attachment with the excel-icon
        attachment_subject  = gc_save_file_name.
        CONCATENATE '&SO_FILENAME=' attachment_subject INTO wa_attachment_header.
        APPEND wa_attachment_header TO t_attachment_header.
* Attachment
        sood_bytecount = bytecount.  " next method expects sood_bytecount instead of any positive integer *sigh*
        cl_document->add_attachment(  i_attachment_type    = 'XLS' "#EC NOTEXT
                                      i_attachment_subject = attachment_subject
                                      i_attachment_size    = sood_bytecount
                                      i_att_content_hex    = t_rawdata
                                      i_attachment_header  = t_attachment_header ).

* add document to send request
        cl_send_request->set_document( cl_document ).

* set sender in case if no own email is availabe
*        cl_sender  = cl_cam_address_bcs=>create_internet_address( 'sender@sender.sender' ).
*        cl_send_request->set_sender( cl_sender ).

* add recipient(s) - here only 1 will be needed
        send_to = p_email.
        IF send_to IS INITIAL.
          send_to = 'no_email@no_email.no_email'.  " Place into SOST in any case for demonstration purposes
        ENDIF.
        cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
        cl_send_request->add_recipient( cl_recipient ).

* Und abschicken
        sent = cl_send_request->send( i_with_error_screen = 'X' ).

        COMMIT WORK.

        IF sent IS INITIAL.
          MESSAGE i500(sbcoms) WITH p_email.
        ELSE.
          MESSAGE s022(so).
          MESSAGE 'Document ready to be sent - Check SOST or SCOT' TYPE 'I'.
        ENDIF.

      CATCH cx_bcs INTO bcs_exception.
        errortext = bcs_exception->if_message~get_text( ).
        MESSAGE errortext TYPE 'I'.

    ENDTRY.
  ENDMETHOD.                    "send_email


ENDCLASS.                    "lcl_output IMPLEMENTATION
