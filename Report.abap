
*&---------------------------------------------------------------------*
*& Report ZMV_EXCL_PRG
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zmv_excl_prg.

INCLUDE zmvexcl_prg_top.
INCLUDE zmv_excl_prg_sel.
INCLUDE zmv_excl__prg_f01.
INCLUDE zmv_excl_prg_status_pbo.
INCLUDE zmv_excl_prg_user_command_pai.


START-OF-SELECTION.

  PERFORM get_data.
  CALL SCREEN 0100.
