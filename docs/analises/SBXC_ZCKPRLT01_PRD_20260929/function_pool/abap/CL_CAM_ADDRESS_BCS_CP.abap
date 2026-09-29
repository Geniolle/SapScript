class-pool .
*"* class pool for class CL_CAM_ADDRESS_BCS

*"* local type definitions
include CL_CAM_ADDRESS_BCS============ccdef.

*"* class CL_CAM_ADDRESS_BCS definition
*"* public declarations
  include CL_CAM_ADDRESS_BCS============cu.
*"* protected declarations
  include CL_CAM_ADDRESS_BCS============co.
*"* private declarations
  include CL_CAM_ADDRESS_BCS============ci.
endclass. "CL_CAM_ADDRESS_BCS definition

*"* macro definitions
include CL_CAM_ADDRESS_BCS============ccmac.
*"* local class implementation
include CL_CAM_ADDRESS_BCS============ccimp.

class CL_CAM_ADDRESS_BCS implementation.
*"* method's implementations
  include if_os_state_macros.
  include methods.
endclass. "CL_CAM_ADDRESS_BCS implementation

