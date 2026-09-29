class-pool .
*"* class pool for class CL_DISTRIBUTIONLIST_BCS

*"* local type definitions
include CL_DISTRIBUTIONLIST_BCS=======ccdef.

*"* class CL_DISTRIBUTIONLIST_BCS definition
*"* public declarations
  include CL_DISTRIBUTIONLIST_BCS=======cu.
*"* protected declarations
  include CL_DISTRIBUTIONLIST_BCS=======co.
*"* private declarations
  include CL_DISTRIBUTIONLIST_BCS=======ci.
endclass. "CL_DISTRIBUTIONLIST_BCS definition

*"* macro definitions
include CL_DISTRIBUTIONLIST_BCS=======ccmac.
*"* local class implementation
include CL_DISTRIBUTIONLIST_BCS=======ccimp.

class CL_DISTRIBUTIONLIST_BCS implementation.
*"* method's implementations
  include if_os_state_macros.
  include methods.
endclass. "CL_DISTRIBUTIONLIST_BCS implementation

