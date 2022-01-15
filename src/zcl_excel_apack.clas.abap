CLASS zcl_excel_apack DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    INTERFACES if_apack_manifest.

    METHODS: constructor.

    ALIASES descriptor FOR if_apack_manifest~descriptor.

  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_excel_apack IMPLEMENTATION.

  METHOD constructor.

    descriptor-group_id        = 'github.com/abap2xlsx/abap2xlsx'.
    descriptor-artifact_id     = 'abap2xlsx'.
    descriptor-version         = zcl_excel=>version.
    descriptor-repository_type = 'abapGit'.
    descriptor-git_url         = 'https://github.com/abap2xlsx/abap2xlsx.git'.

  ENDMETHOD.

ENDCLASS.
