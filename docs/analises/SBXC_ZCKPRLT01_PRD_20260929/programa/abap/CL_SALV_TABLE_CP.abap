class-pool .
*"* class pool for class CL_SALV_TABLE

*"* local type definitions
include CL_SALV_TABLE=================ccdef.

*"* class CL_SALV_TABLE definition
*"* public declarations
  include CL_SALV_TABLE=================cu.
*"* protected declarations
  include CL_SALV_TABLE=================co.
*"* private declarations
  include CL_SALV_TABLE=================ci.
endclass. "CL_SALV_TABLE definition

*"* macro definitions
include CL_SALV_TABLE=================ccmac.
*"* local class implementation
include CL_SALV_TABLE=================ccimp.

*"* test class
include CL_SALV_TABLE=================ccau.

class CL_SALV_TABLE implementation.
*"* method's implementations
  include methods.
endclass. "CL_SALV_TABLE implementation

