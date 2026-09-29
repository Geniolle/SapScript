class-pool .
*"* class pool for class ZCL_BIM

*"* local type definitions
include ZCL_BIM=======================ccdef.

*"* class ZCL_BIM definition
*"* public declarations
  include ZCL_BIM=======================cu.
*"* protected declarations
  include ZCL_BIM=======================co.
*"* private declarations
  include ZCL_BIM=======================ci.
endclass. "ZCL_BIM definition

*"* macro definitions
include ZCL_BIM=======================ccmac.
*"* local class implementation
include ZCL_BIM=======================ccimp.

*"* test class
include ZCL_BIM=======================ccau.

class ZCL_BIM implementation.
*"* method's implementations
  include methods.
endclass. "ZCL_BIM implementation

