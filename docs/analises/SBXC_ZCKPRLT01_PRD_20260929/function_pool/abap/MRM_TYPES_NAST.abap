*----------------------------------------------------------------------*
*   INCLUDE MRM_TYPES_NAST                                             *
*----------------------------------------------------------------------*

* Objektkey Nachrichten für die Logistik RePrü
* (in der herkömmlichen RePrü ist die Positionsnr BUZEI 3stellig)
TYPES: BEGIN OF TYP_OBJKY,
         BUKRS(4)  TYPE C,             "bei Log.RePrü immer '$$$$'   4
         BELNR     LIKE RBKP-BELNR,    "Rechnungsbelegnummer        10
         GJAHR     LIKE RBKP-GJAHR,    "Rechnungsgeschäftsjahr       4
         BUZEI     LIKE RSEG-BUZEI,    "Rechnungsposition            6
         FILLER(6) TYPE C,             "nicht benutzt                6
       END OF TYP_OBJKY.               "                            30
