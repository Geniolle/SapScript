class-pool .
*"* class pool for class CL_SALV_COLUMN

*"* local type definitions
include CL_SALV_COLUMN================ccdef.

*"* class CL_SALV_COLUMN definition
*"* public declarations
  include CL_SALV_COLUMN================cu.
*"* protected declarations
  include CL_SALV_COLUMN================co.
*"* private declarations
  include CL_SALV_COLUMN================ci.
endclass. "CL_SALV_COLUMN definition

*"* macro definitions
include CL_SALV_COLUMN================ccmac.
*"* local class implementation
include CL_SALV_COLUMN================ccimp.

*"* test class
include CL_SALV_COLUMN================ccau.

class CL_SALV_COLUMN implementation.
*"* method's implementations
  include methods.
endclass. "CL_SALV_COLUMN implementation

