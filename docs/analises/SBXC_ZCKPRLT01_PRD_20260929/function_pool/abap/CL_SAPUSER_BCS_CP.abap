class-pool .
*"* class pool for class CL_SAPUSER_BCS

*"* local type definitions
include CL_SAPUSER_BCS================ccdef.

*"* class CL_SAPUSER_BCS definition
*"* public declarations
  include CL_SAPUSER_BCS================cu.
*"* protected declarations
  include CL_SAPUSER_BCS================co.
*"* private declarations
  include CL_SAPUSER_BCS================ci.
endclass. "CL_SAPUSER_BCS definition

*"* macro definitions
include CL_SAPUSER_BCS================ccmac.
*"* local class implementation
include CL_SAPUSER_BCS================ccimp.

class CL_SAPUSER_BCS implementation.
*"* method's implementations
  include if_os_state_macros.
  include methods.
endclass. "CL_SAPUSER_BCS implementation

