*&---------------------------------------------------------------------*
*& PoolMóds.   /SBXC/ZCKPRLT01                                         *
*&                                                                     *
*----------------------------------------------------------------------*
* Programa Cockpit para processamento de facturas vindas de
* sistemas externos
* Author: SBX Consulting
* Data: Outubro 2011
*----------------------------------------------------------------------*

INCLUDE /sbxc/zckprlt01_top                       .  " global Data
INCLUDE /sbxc/zckprlt01_lcl                       .  " classes
INCLUDE /sbxc/zckprlt01_f01                       .  " FORM-Routines
INCLUDE /sbxc/zckprlt01_o01.
INCLUDE /sbxc/zckprlt01_i01.


INITIALIZATION.

v_uname = sy-uname.
* retirada a opção de ecran de seleção dinamica
*
*  PERFORM f_init_push.


*AT SELECTION-SCREEN.
*
*  IF sscrfields-ucomm EQ 'BOTAO1'.
*    CLEAR sscrfields-ucomm.
*
*    PERFORM valida_processos.
*    PERFORM gera_ecra_de_selecao.
*
*  ENDIF.

START-OF-SELECTION.


*1) Definição de estruturas de cabeçalho e linhas
* Valida autorizações para os botões
  PERFORM valida_processos.

*2) Carrega parametrização
  PERFORM get_configuration.

  PERFORM get_data.

*3)
  PERFORM inicializa_control.

*4)
  PERFORM build_alvs_init .


  CALL SCREEN 100.
