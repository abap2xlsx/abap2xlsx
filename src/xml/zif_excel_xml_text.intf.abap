INTERFACE zif_excel_xml_text
  PUBLIC.

  INTERFACES zif_excel_xml_character_data.

  ALIASES append_child
    FOR zif_excel_xml_character_data~append_child.
  ALIASES clone
    FOR zif_excel_xml_character_data~clone.
  ALIASES create_iterator
    FOR zif_excel_xml_character_data~create_iterator.
  ALIASES get_attributes
    FOR zif_excel_xml_character_data~get_attributes.
  ALIASES get_children
    FOR zif_excel_xml_character_data~get_children.
  ALIASES get_first_child
    FOR zif_excel_xml_character_data~get_first_child.
  ALIASES get_name
    FOR zif_excel_xml_character_data~get_name.
  ALIASES get_next
    FOR zif_excel_xml_character_data~get_next.
  ALIASES get_value
    FOR zif_excel_xml_character_data~get_value.
  ALIASES set_value
    FOR zif_excel_xml_character_data~set_value.

ENDINTERFACE.
