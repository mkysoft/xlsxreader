class ZCL_XLSX2ABAP definition
  public
  create public .

public section.

  class-methods TOITAB
    importing
      !IV_FILE type XSTRING
    returning
      value(R_ITAB) type ref to DATA .
protected section.
private section.
ENDCLASS.



CLASS ZCL_XLSX2ABAP IMPLEMENTATION.


  method TOITAB.
  endmethod.
ENDCLASS.
