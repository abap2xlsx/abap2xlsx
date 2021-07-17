Note these instructions will get ABAP2XLSX working on an older 620 system. In this case I am doing it on a 4.70 Enterprise ext 2 system.
 
The guide draws from the existing installation guides and includes additional steps specific to 620.
 
This is based on the Daily build of ABAP2XLSX version 4, revision 103, dated 19th of January 2011.
 
1. Download abap2xlsx_Daily.nugg.zip and extract the .nugg file from it.
2. Download abap2xlsx6.20patch.nugg.zip and extract the .nugg file from it.
3. Use ZSAPLINK to install NUGG_ABAB2XLS_UTILS620.nugg. Ensure overwrite Originals is ticked. You should get green lights against the 3 classes and 2 interfaces that it imports.
4. Open abap2xlsx_Daily.nugg in text editor and do search and replace according to following table:
 
 
Search	Replace with	# of occurrences
CL_OBJECT_COLLECTION	ZCL_OBJECT_COLLECTION	61
IF_OBJECT_COLLECTION	ZIF_OBJECT_COLLECTION	68
CL_ABAP_ZIP	ZCL_ABAP_ZIP	3
 
 
5. Save that new copy of abap2xlsx_Daily.nugg
6. Use SAPLINK to install your modified abap2xlsx_Daily.nugg. Again you should get a green light on everything it imports.
7. Use transaction SE80 (Object Navigator) to activate the newly imported entries.
 
Select the Inactive Objects from the dropdown. Activate in the following order :
 
Activate all domains
Activate all data elements
Activate all Database Tables / Structures except
      ZEXCEL_S_FIELDCATALOG
      ZEXCEL_S_WORKSHEET_COLUMNDIME
      ZEXCEL_S_WORKSHEET_ROWDIMENSIO
Note you may get a warning on activation. Continue on.
Activate all Table Types except
      ZEXCEL_T_FIELDCATALOG Table binding field catalog
      ZEXCEL_T_WORKSHEET_COLUMNDIME Collection of column dimensions
      ZEXCEL_T_WORKSHEET_ROWDIMENSIO Collection of row dimensions
Activate all interfaces
Activate all classes (see step 8 below for the errors you will encounter)
Activate remaining Database Tables /  Structures (if any error occurs open the structure and double click on  the class object, SAP needs to refresh its buffer)
      ZEXCEL_S_FIELDCATALOG
      ZEXCEL_S_WORKSHEET_COLUMNDIME
      ZEXCEL_S_WORKSHEET_ROWDIMENSIO
Activate remaining Table Types (if any error  occurs open the structure and double click on the class object, SAP  needs to refresh its buffer)
      ZEXCEL_T_FIELDCATALOG
      ZEXCEL_T_WORKSHEET_COLUMNDIME
      ZEXCEL_T_WORKSHEET_ROWDIMENSIO
 
8. You will find during activation on a 620 system that ZCL_EXCEL_HYPERLINKS class is a particular problem. It is easiest to just deactivate and exclude this (remove any reference to it from any other class with the abap2xlsx group). Note that with these changes made, any hyperlink functionality in your programs will not actually make use of the hyperlink functionality.
 
Hyperlinks Cleanup :
 
The main issue is hyperlinks->get_iterator is unknown. Quick fix is to ignore any hyperlink functionality.

Class : ZCL_EXCEL_WORKSHEET

Method : GET_HYPERLINKS_ITERATOR

Comment out  : **eo_iterator = hyperlinks->get_iterator( ).**


Class : ZCL_EXCEL_WORKSHEET

Method : GET_HYPERLINKS_SIZE

Comment out  : **ep_size = hyperlinks->size( ).**
 
 
Class : ZCL_EXCEL_WRITER_2007

Method : CREATE_XL_SHEET_RELS

Comment out  :
**lo_iterator = io_worksheet->get_hyperlinks_iterator( ).**
**WHILE lo_iterator->ZIF_OBJECT_COLLECTION_iterator~has_next( ) EQ abap_true.**
**lo_link ?= lo_iterator->ZIF_OBJECT_COLLECTION_iterator~get_next( ).**
**ADD 1 TO lv_relation_id.**
 
**lv_value = lv_relation_id.**
**CONDENSE lv_value.**
**CONCATENATE 'rId' lv_value INTO lv_value.**

**lo_element = lo_document->create_simple_element( name   = lc_xml_node_relationship**
                                                 **parent = lo_document ).**
**lo_element->set_attribute_ns( name  = lc_xml_attr_id**
                              **value = lv_value ).**
**lo_element->set_attribute_ns( name  = lc_xml_attr_type**
                              **value = lc_xml_node_rid_link_tp ).**
 
**lv_value = lo_link->get_url( ).**
**lo_element->set_attribute_ns( name  = lc_xml_attr_target**
                              **value = lv_value ).**
**lo_element->set_attribute_ns( name  = lc_xml_attr_target_mode**
                              **value = lc_xml_val_external ).**
**lo_element_root->append_child( new_child = lo_element ).**
**ENDWHILE.**
 
And to help with any transport issues later on :
 
Class : ZCL_EXCEL_HYPERLINKS

Method : ADD

Comment out : **data_validations->add( ip_data_validation ).**
 
8a (Alternative to allow Hyperlinks - Tested with ABAP2XLSX version 6) - This is unconfirmed but it seems to be working for us.
 
Class : ZCL_EXCEL_WORKSHEET

Method : GET_HYPERLINKS_ITERATOR

Change to : **eo_iterator = hyperlinks->zif_object_collection~get_iterator( ).**


Class : ZCL_EXCEL_WORKSHEET

Method : GET_HYPERLINKS_SIZE

Change to: **ep_size = hyperlinks->zif_object_collection~size( ).**
         
The Class ZCL_EXCEL_WRITER_2007 Method CREATE_XL_SHEET_RELS code can remain uncommented. 
              
 
9. 620 is picky about parameter names being omitted if there is more than one possible parameter. Minor changes to the code are required in the following area :
 
Class : ZCL_EXCEL_WORKSHEET

Method : CALCULATE_COLUMN_WIDTHS

Find : **column_dimension->set_width( <auto_size>-width ).**

Replace with : **call method column_dimension->set_width exporting ip_width = <auto_size>-width.**
 
10. 620 doesn't support CONCATENATE...RESPECTING BLANKS or REPLACE...REGEX... so this needs to be changed to a more traditional CONCATENATE.
 
Class : ZCL_EXCEL_WRITER_2007

Method :  CREATE_DOCPROPS_CORE

Find : 
**CONCATENATE lv_date lv_time INTO lv_value RESPECTING BLANKS.**
**REPLACE ALL OCCURRENCES OF REGEX '([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})' IN lv_value WITH '--T::Z'.**

Replace with :
**CONCATENATE lv_date+0(4) '-' lv_date+4(2) '-' lv_date+6(2) 'T' lv_time+0(2) ':' lv_time+2(2) ':' lv_time+4(2) 'Z' INTO lv_value.**

There are two occurances to replace.              
              
 
 
11. You should now have no problems activating everything except for one of the Demo reports. Three of the four can be fixed as follows :
 
ZDEMO_EXCEL20
Fix not attempted on this one.
 
Thanks to Regan MacDonald