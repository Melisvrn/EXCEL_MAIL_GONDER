*----------------------------------------------------------------------*
***INCLUDE ZMV_EXCL_PRG_USER_COMMAND_0I01.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0100  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE USER_COMMAND_0100 INPUT.

  CALL METHOD ALV_0100->CHECK_CHANGED_DATA
    IMPORTING
      E_VALID   =   LV_VALID
    CHANGING
      C_REFRESH = LV_REFRESH  .

  CASE SY-UCOMM.
    WHEN '&BACK' OR '&CANCEL' OR '&EXIT' .
      LEAVE TO SCREEN 0.
    WHEN '&EXCEL'.
      PERFORM  EXCEL_FORMAT_DOWNLOAD.
    WHEN '&E-MAIL'.
      PERFORM EXCEL_MAIL_GONDER.
    WHEN OTHERS.
  ENDCASE.

*  CALL METHOD ALV_0100->REFRESH_TABLE_DISPLAY.
*  PERFORM REGISTER.


ENDMODULE.
