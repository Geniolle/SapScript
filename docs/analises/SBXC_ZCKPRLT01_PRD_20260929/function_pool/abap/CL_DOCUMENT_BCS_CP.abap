class-pool .
*"* class pool for class CL_DOCUMENT_BCS

*"* local type definitions
include CL_DOCUMENT_BCS===============ccdef.

*"* class CL_DOCUMENT_BCS definition
*"* public declarations
  include CL_DOCUMENT_BCS===============cu.
*"* protected declarations
  include CL_DOCUMENT_BCS===============co.
*"* private declarations
  include CL_DOCUMENT_BCS===============ci.
endclass. "CL_DOCUMENT_BCS definition

*"* macro definitions
include CL_DOCUMENT_BCS===============ccmac.
*"* local class implementation
include CL_DOCUMENT_BCS===============ccimp.

class CL_DOCUMENT_BCS implementation.
*"* method's implementations
  include if_os_state_macros.
  include methods.
endclass. "CL_DOCUMENT_BCS implementation

