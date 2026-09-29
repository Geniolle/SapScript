*&---------------------------------------------------------------------*
*&  Include           ZCKPRLT01_O01                                    *
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  STATUS_0100  OUTPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE status_0100 OUTPUT.
  SET PF-STATUS 'STATUS'.
  SET TITLEBAR 'CKP'.

ENDMODULE.                 " STATUS_0100  OUTPUT
*&---------------------------------------------------------------------*
*&      Module  PBO  OUTPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE pbo OUTPUT.
  DATA it_exfcode  type TABLE OF rsmpe-func ##NEEDED.
  DATA wa_exfcode  type  rsmpe-func ##NEEDED.

  refresh it_exfcode.

  IF disp_doc_active = 'X'.
    MOVE 'DOC_ON' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
  ELSE.
    MOVE 'DOC_OFF' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
  ENDIF.
  IF disp_fil_active = 'X'.
    MOVE 'TOOG_ON' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
  ELSE.
    MOVE 'TOOG_OFF' TO wa_exfcode.
    APPEND wa_exfcode TO it_exfcode.
  ENDIF.
  SET PF-STATUS 'MAIN100' EXCLUDING it_exfcode.
ENDMODULE.                 " PBO  OUTPUT
