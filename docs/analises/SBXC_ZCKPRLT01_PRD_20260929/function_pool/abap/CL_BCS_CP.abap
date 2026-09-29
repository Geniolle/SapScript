class-pool .
*"* class pool for class CL_BCS

*"* local type definitions
include CL_BCS========================ccdef.

*"* class CL_BCS definition
*"* public declarations
  include CL_BCS========================cu.
*"* protected declarations
  include CL_BCS========================co.
*"* private declarations
  include CL_BCS========================ci.
endclass. "CL_BCS definition

*"* macro definitions
include CL_BCS========================ccmac.
*"* local class implementation
include CL_BCS========================ccimp.

class CL_BCS implementation.
*"* method's implementations
  include methods.
endclass. "CL_BCS implementation

