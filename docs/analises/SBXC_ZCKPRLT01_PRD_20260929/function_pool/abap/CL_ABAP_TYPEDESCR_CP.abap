class-pool .
*"* class pool for class CL_ABAP_TYPEDESCR

*"* local type definitions
include CL_ABAP_TYPEDESCR=============ccdef.

*"* class CL_ABAP_TYPEDESCR definition
*"* public declarations
  include CL_ABAP_TYPEDESCR=============cu.
*"* protected declarations
  include CL_ABAP_TYPEDESCR=============co.
*"* private declarations
  include CL_ABAP_TYPEDESCR=============ci.
endclass. "CL_ABAP_TYPEDESCR definition

*"* macro definitions
include CL_ABAP_TYPEDESCR=============ccmac.
*"* local class implementation
include CL_ABAP_TYPEDESCR=============ccimp.

class CL_ABAP_TYPEDESCR implementation.
*"* method's implementations
  include methods.
endclass. "CL_ABAP_TYPEDESCR implementation

