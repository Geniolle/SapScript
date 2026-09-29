class-pool .
*"* class pool for class /SBXC/CO_ICNDDOCUMENT_SERVICE

*"* local type definitions
include /SBXC/CO_ICNDDOCUMENT_SERVICE=ccdef.

*"* class /SBXC/CO_ICNDDOCUMENT_SERVICE definition
*"* public declarations
  include /SBXC/CO_ICNDDOCUMENT_SERVICE=cu.
*"* protected declarations
  include /SBXC/CO_ICNDDOCUMENT_SERVICE=co.
*"* private declarations
  include /SBXC/CO_ICNDDOCUMENT_SERVICE=ci.
endclass. "/SBXC/CO_ICNDDOCUMENT_SERVICE definition

*"* macro definitions
include /SBXC/CO_ICNDDOCUMENT_SERVICE=ccmac.
*"* local class implementation
include /SBXC/CO_ICNDDOCUMENT_SERVICE=ccimp.

class /SBXC/CO_ICNDDOCUMENT_SERVICE implementation.
*"* method's implementations
  include methods.
endclass. "/SBXC/CO_ICNDDOCUMENT_SERVICE implementation

