class-pool .
*"* class pool for class CL_ALV_TABLE_CREATE

*"* local classes
include CL_ALV_TABLE_CREATE===========cl.

*"* class CL_ALV_TABLE_CREATE definition
*"* public declarations
  include CL_ALV_TABLE_CREATE===========cu.
*"* protected declarations
  include CL_ALV_TABLE_CREATE===========co.
*"* private declarations
  include CL_ALV_TABLE_CREATE===========ci.
endclass. "CL_ALV_TABLE_CREATE definition

*"* test class
include CL_ALV_TABLE_CREATE===========ccau.

class CL_ALV_TABLE_CREATE implementation.
*"* method's implementations
  include methods.
endclass. "CL_ALV_TABLE_CREATE implementation

