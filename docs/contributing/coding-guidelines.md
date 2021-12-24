# Naming convention of variables, etc.

In abap2xlsx, over time, alas, the ABAP code came to mix different [naming standards](https://github.com/abap2xlsx/abap2xlsx/issues/773). Naming standards may vary from one class to another, but one naming standards is usually correctly applied in each class.

It's not possible to impose one naming standards by fixing the existing names, because clients may have developed custom programs which may already refer to them.
When it's about adding a new variable, parameter or type in an existing object or method, it's embarrassing to impose one naming standards because that could be inconsistent with the variables, parameters or types which already exist at this place.

Here is the chosen compromise so that to keep the ABAP code the most consistent possible everywhere:
- If you fix a method, reuse its naming standards so that its code looks consistent. If this method has several naming standards, choose the most used one.
- If you create an ABAP object or a method implementation, choose the names according to the below rules.
- If you create an attribute, a method definition or a type in a class or interface, choose the naming standards used in its section.
NB: that means, if you need to create and maintain several objects to fix one issue or to make one pull request, you may end with using different naming standards.

Rules for creating an ABAP object or a method implementation (mix of Hungarian notation (prefixes) and [Clean Code](https://github.com/SAP/styleguides/blob/main/clean-abap/CleanABAP.md#avoid-encodings-esp-hungarian-notation-and-prefixes)):
- Method names
  - clean code
- Local class names
  - LCL_
  - LCX_ for the only exception class
- Types
  - elementary: TV_
  - structure: TS_
  - table: TT_
- Instance and class Attributes (not constants)
  - clean code
- constant and global attributes
  - C_
- Local variables
  - elementary: LV_
  - object: LO_
  - structure: LS_
  - table: LT_
  - data reference: LR_
- Local constants
  - LC_
- Field symbols
  - elementary: <LV_
  - structure: <LS_
  - table: <LT_
- IMPORTING parameters
  - elementary: IV_
  - object: IO_
  - structure: IS_
  - table: IT_
- EXPORTING parameters
  - elementary: EV_
  - object: EO_
  - structure: ES_
  - table: ET_
- RETURNING parameters
  - elementary: RV_
  - object: RO_
  - structure: RS_
  - table: RT_
  - data reference: RR_
  - NB: don't use the general suffix "RESULT", instead use RO_WORKSHEET if method name is GET_ACTIVE_WORKSHEET
- CHANGING parameters
  - elementary: CV_
  - structure: CS_
  - table: CT_
  - object: CO_

Here is a list of other naming standards, that you should use if they occur in existing method implementations or in the definition part of existing classes and interfaces, and if they are majority:
- Types
  - Miscellaneous namings like T_, TY_, LTY, MTY_ were often used
- Field symbols
  - <FS_ and <F_ were often used
- elementary parameters
  - P was often used instead of V, for instance IP_ or EP_
- RETURNING parameters
  - E was often used instead of R
- CHANGING parameters
  - X was often used instead of C
