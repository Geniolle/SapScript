class-pool .
*"* class pool for class CX_BCS

*"* local type definitions
include CX_BCS========================ccdef.

*"* class CX_BCS definition
*"* public declarations
  include CX_BCS========================cu.
*"* protected declarations
  include CX_BCS========================co.
*"* private declarations
  include CX_BCS========================ci.
endclass. "CX_BCS definition

*"* macro definitions
include CX_BCS========================ccmac.
*"* local class implementation
include CX_BCS========================ccimp.

class CX_BCS implementation.
*"* method's implementations
  include methods.
endclass. "CX_BCS implementation

