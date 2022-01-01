*&---------------------------------------------------------------------*
*&  Include           ZMV_EXCL_STR_SEL
*&---------------------------------------------------------------------*

SELECTION-SCREEN BEGIN OF BLOCK B WITH FRAME TITLE TEXT-001.
SELECT-OPTIONS :S_MATNR FOR MARA-MATNR.
PARAMETERS :P_WERKS TYPE MARC-WERKS MATCHCODE OBJECT ZMV_SRCHELP.
SELECTION-SCREEN END OF BLOCK B .

AT SELECTION-SCREEN OUTPUT.
