class-pool .
*"* class pool for class CL_ALV_CHANGED_DATA_PROTOCOL

*"* local classes
include CL_ALV_CHANGED_DATA_PROTOCOL==cl.

*"* class CL_ALV_CHANGED_DATA_PROTOCOL definition
*"* public declarations
  include CL_ALV_CHANGED_DATA_PROTOCOL==cu.
*"* protected declarations
  include CL_ALV_CHANGED_DATA_PROTOCOL==co.
*"* private declarations
  include CL_ALV_CHANGED_DATA_PROTOCOL==ci.
endclass. "CL_ALV_CHANGED_DATA_PROTOCOL definition

class CL_ALV_CHANGED_DATA_PROTOCOL implementation.
*"* method's implementations
  include methods.
endclass. "CL_ALV_CHANGED_DATA_PROTOCOL implementation

