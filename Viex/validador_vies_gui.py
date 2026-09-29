"""
Validador VIES - Aplicação GUI Windows
Interface gráfica para validação individual e em massa de VATs.
"""

import PySimpleGUI as sg
import threading
import traceback
import time
from pathlib import Path
from typing import Optional, Callable
from vies_service import (
    validar_vies,
    validar_vies_com_retry,
    ViesResult,
    ViesResultStatus,
)
from vies_excel import ViesExcel


# ============================================================
# CONFIGURAÇÃO VISUAL
# ============================================================

sg.theme("LightBlue2")
sg.set_options(font=("Segoe UI", 10))

WINDOW_TITLE = "Validador VIES"
WINDOW_SIZE = (700, 600)
WINDOW_SIZE_MASSA = (800, 700)


# ============================================================
# TELAS
# ============================================================

class TelaPrincipal:
    """Tela do menu principal."""

    def criar_layout(self):
        return [
            [sg.Text("Validador VIES", font=("Segoe UI", 20, "bold"), justification="center")],
            [sg.Text("Validação de VAT (NIF) europeus", justification="center", text_color="gray")],
            [sg.Text("")],  # espaçamento
            [sg.Button("Pesquisa Individual", size=(30, 2), pad=((100, 100), 10))],
            [sg.Button("Pesquisa em Massa", size=(30, 2), pad=((100, 100), 10))],
            [sg.Text("")],  # espaçamento
            [sg.Button("Sair", size=(30, 2), pad=((100, 100), 10))],
        ]

    def processar(self, values, window):
        return None  # voltará para aqui após submenus


class TelaPesquisaIndividual:
    """Tela de pesquisa individual."""

    def criar_layout(self):
        return [
            [sg.Text("Pesquisa Individual", font=("Segoe UI", 16, "bold"))],
            [sg.Text("")],
            [sg.Text("Informe o VAT (ex: PT504772694, ESA28017895):")],
            [sg.InputText(key="VAT_INPUT", size=(40, 1))],
            [sg.Text("")],
            [sg.Button("Validar", size=(15, 1)), sg.Button("Voltar", size=(15, 1))],
            [sg.Text("")],
            [sg.Multiline(size=(70, 20), key="RESULTADO", disabled=True)],
        ]

    def processar(self, values, window):
        if values.get("VAT_INPUT"):
            vat_input = values["VAT_INPUT"].strip()

            try:
                resultado = validar_vies(vat_input)
                texto_resultado = self._formatar_resultado(resultado)
                window["RESULTADO"].update(texto_resultado)

            except ValueError as exc:
                sg.popup_error(f"Erro de validação:\n{str(exc)}")
                window["VAT_INPUT"].update("")

            except RuntimeError as exc:
                sg.popup_error(f"Erro na consulta VIES:\n{str(exc)}")
                window["VAT_INPUT"].update("")

    def _formatar_resultado(self, resultado: ViesResult) -> str:
        linhas = [
            "=" * 60,
            "RESULTADO DA VALIDAÇÃO",
            "=" * 60,
            "",
            f"VAT:             {resultado.vat}",
            f"Válido no VIES:  {'SIM' if resultado.valid else 'NÃO'}",
            f"País:            {resultado.country_code}",
            f"Número:          {resultado.vat_number}",
        ]

        if resultado.valid:
            linhas.append(
                f"Nome:            {resultado.name or 'NÃO DISPONIBILIZADO'}"
            )

            if resultado.address:
                linhas_morada = resultado.address.splitlines()
                linhas.append(f"Morada:          {linhas_morada[0]}")
                for parte in linhas_morada[1:]:
                    linhas.append(f"                 {parte}")
            else:
                linhas.append("Morada:          NÃO DISPONIBILIZADA")

        else:
            linhas.append("Nome:            -")
            linhas.append("Morada:          -")

        linhas.append(
            f"Data consulta:   {resultado.request_date or '-'}"
        )

        linhas.append("")
        linhas.append("=" * 60)

        return "\n".join(linhas)


class TelaPesquisaMassa:
    """Tela de pesquisa em massa."""

    def criar_layout(self):
        return [
            [sg.Text("Pesquisa em Massa", font=("Segoe UI", 16, "bold"))],
            [sg.Text("Selecione um ficheiro Excel com a sheet 'NIF'")],
            [sg.Text("")],
            [
                sg.InputText(key="FICHEIRO_PATH", disabled=True, size=(50, 1)),
                sg.FileBrowse(
                    "Selecionar ficheiro",
                    target="FICHEIRO_PATH",
                    file_types=(("Excel", "*.xlsx"),),
                    key="FICHEIRO_BROWSER"
                ),
            ],
            [sg.Text("")],
            [
                sg.Button("Iniciar Validação", size=(15, 1)),
                sg.Button("Voltar", size=(15, 1)),
            ],
            [sg.Text("")],
            [
                sg.Multiline(
                    size=(80, 25),
                    key="LOG_MASSA",
                    disabled=True,
                    autoscroll=True
                )
            ],
            [
                sg.ProgressBar(
                    max_value=100,
                    orientation="h",
                    size=(60, 20),
                    key="PROGRESS_BAR"
                ),
                sg.Text("", key="PROGRESS_TEXT", size=(10, 1)),
            ],
        ]

    def processar(self, values, window, thread_func: Optional[Callable] = None):
        ficheiro = values.get("FICHEIRO_PATH", "").strip()

        if not ficheiro and values.get("FICHEIRO_BROWSER"):
            ficheiro = values["FICHEIRO_BROWSER"]
            window["FICHEIRO_PATH"].update(ficheiro)

        if not ficheiro:
            sg.popup_warning("Selecione um ficheiro Excel primeiro.")
            return

        if not Path(ficheiro).exists():
            sg.popup_error(f"Ficheiro não encontrado:\n{ficheiro}")
            return

        # Desativar botões durante processamento
        window["Iniciar Validação"].update(disabled=True)
        window["FICHEIRO_BROWSER"].update(disabled=True)
        window["Voltar"].update(disabled=True)

        # Limpar log e progresso
        window["LOG_MASSA"].update("")
        window["PROGRESS_BAR"].update(0)
        window["PROGRESS_TEXT"].update("0 / 0")

        # Iniciar processamento em thread
        if thread_func:
            t = threading.Thread(
                target=thread_func,
                args=(ficheiro, window),
                daemon=True
            )
            t.start()

    @staticmethod
    def processar_massa_thread(caminho_ficheiro: str, window):
        """Thread worker para processamento em massa com retry automático."""
        try:
            excel = ViesExcel(caminho_ficheiro)
            excel.abrir()

            # Primeira passagem - processar todos os VATs
            registos = excel.obter_registos()

            if not registos:
                window["LOG_MASSA"].print("Nenhum VAT encontrado no ficheiro.")
                excel.fechar()
                return

            # Estatísticas
            total_ficheiro = len(registos)
            total_pendentes = len(registos)
            consultados = 0
            validos = 0
            invalidos = 0
            erros_temp = 0

            # Mostrar lista inicial
            window["LOG_MASSA"].print(f"VATs encontrados: {total_ficheiro}")
            window["LOG_MASSA"].print(f"Pendentes para processar: {total_pendentes}")
            window["LOG_MASSA"].print("=" * 60)
            for idx, registo in enumerate(registos, 1):
                window["LOG_MASSA"].print(f"[ ] {idx:>3}/{total_pendentes}  {registo['vat']}")

            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print("A iniciar validação...")
            window["LOG_MASSA"].print("=" * 60)
            window["LOG_MASSA"].print("")

            # Mapa de resultados para segunda passagem
            registos_com_erro = []

            # PRIMEIRA PASSAGEM
            for idx, registo in enumerate(registos, 1):
                linha_excel = registo["linha"]
                vat = registo["vat"]

                def on_status(status: ViesResultStatus):
                    """Atualizar GUI com status de retry."""
                    if status.is_success and status.result:
                        pass  # Atualizar abaixo
                    elif status.attempt < status.total_attempts:
                        if status.is_temporary_error:
                            window["LOG_MASSA"].print(
                                f"    Tentativa {status.attempt}/{status.total_attempts} - "
                                f"{status.error_message}"
                            )

                try:
                    status = validar_vies_com_retry(vat, on_retry_status=on_status)

                    if status.is_success and status.result:
                        resultado = status.result
                        consultados += 1

                        # Atualizar Excel
                        excel.atualizar_resultado(
                            linha_excel,
                            resultado.valid,
                            resultado.request_date
                        )
                        excel.guardar()

                        # Estatísticas
                        if resultado.valid:
                            validos += 1
                            estado = "VÁLIDO"
                        else:
                            invalidos += 1
                            estado = "NÃO VÁLIDO"

                        msg = f"[✓] {idx:>3}/{total_pendentes}  {vat:<20} {estado}"
                        window["LOG_MASSA"].print(msg)

                    else:
                        # Erro técnico após retries
                        erros_temp += 1
                        excel.atualizar_erro(linha_excel)
                        excel.guardar()

                        registos_com_erro.append(registo)

                        msg = f"[!] {idx:>3}/{total_pendentes}  {vat:<20} ERRO TÉCNICO"
                        window["LOG_MASSA"].print(msg)

                except Exception as exc:
                    # Erro não retentável
                    consultados += 1
                    window["LOG_MASSA"].print(
                        f"[X] {idx:>3}/{total_pendentes}  {vat:<20} "
                        f"ERRO: {str(exc)[:50]}"
                    )

                # Intervalo entre consultas
                time.sleep(0.2)

                # Atualizar barra de progresso
                progress = int((idx / total_pendentes) * 100)
                window["PROGRESS_BAR"].update(progress)
                window["PROGRESS_TEXT"].update(f"{idx} / {total_pendentes}")

            # SEGUNDA PASSAGEM - recuperação de erros
            if registos_com_erro:
                window["LOG_MASSA"].print("")
                window["LOG_MASSA"].print("=" * 60)
                window["LOG_MASSA"].print(
                    f"Primeira passagem concluída. {len(registos_com_erro)} com erro técnico."
                )
                window["LOG_MASSA"].print("A aguardar 10 segundos antes de tentar novamente...")
                window["LOG_MASSA"].print("=" * 60)
                window["LOG_MASSA"].print("")

                # Espera com feedback
                for segundo in range(10, 0, -1):
                    window["LOG_MASSA"].print(f"Retentativas em {segundo}s...")
                    time.sleep(1)

                window["LOG_MASSA"].print("")
                window["LOG_MASSA"].print("A processar VATs com erro técnico...")
                window["LOG_MASSA"].print("=" * 60)
                window["LOG_MASSA"].print("")

                for idx, registo in enumerate(registos_com_erro, 1):
                    linha_excel = registo["linha"]
                    vat = registo["vat"]

                    try:
                        status = validar_vies_com_retry(vat)

                        if status.is_success and status.result:
                            resultado = status.result
                            consultados += 1

                            # Atualizar Excel
                            excel.atualizar_resultado(
                                linha_excel,
                                resultado.valid,
                                resultado.request_date
                            )
                            excel.guardar()

                            # Remover do contador de erros
                            erros_temp -= 1

                            if resultado.valid:
                                validos += 1
                                estado = "VÁLIDO"
                            else:
                                invalidos += 1
                                estado = "NÃO VÁLIDO"

                            msg = f"[✓] {idx:>3}/{len(registos_com_erro)}  {vat:<20} {estado}"
                            window["LOG_MASSA"].print(msg)

                        else:
                            # Ainda com erro
                            msg = f"[!] {idx:>3}/{len(registos_com_erro)}  {vat:<20} AINDA COM ERRO"
                            window["LOG_MASSA"].print(msg)

                    except Exception as exc:
                        msg = f"[X] {idx:>3}/{len(registos_com_erro)}  {vat:<20} " \
                              f"ERRO: {str(exc)[:50]}"
                        window["LOG_MASSA"].print(msg)

                    time.sleep(0.2)

            # Finalização
            excel.fechar()

            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print("=" * 60)
            window["LOG_MASSA"].print("PROCESSAMENTO FINALIZADO")
            window["LOG_MASSA"].print("=" * 60)
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print(f"Total no ficheiro:          {total_ficheiro}")
            window["LOG_MASSA"].print(f"Consultados nesta execução: {consultados}")
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print(f"✓ Válidos:        {validos}")
            window["LOG_MASSA"].print(f"✕ Não válidos:     {invalidos}")
            window["LOG_MASSA"].print(f"⚠ Erros técnicos:  {erros_temp}")
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print(
                f"Consultas concluídas: {consultados - erros_temp}/{consultados}"
            )
            if erros_temp > 0:
                window["LOG_MASSA"].print(f"Pendentes por erro:  {erros_temp}")
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print(f"Excel atualizado: {caminho_ficheiro}")

        except Exception as exc:
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print("ERRO FATAL")
            window["LOG_MASSA"].print("=" * 60)
            window["LOG_MASSA"].print(str(exc))
            window["LOG_MASSA"].print("")
            window["LOG_MASSA"].print(traceback.format_exc())

        finally:
            # Reativar botões
            window["Iniciar Validação"].update(disabled=False)
            window["FICHEIRO_BROWSER"].update(disabled=False)
            window["Voltar"].update(disabled=False)


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================

class ValidadorVIESApp:
    """Aplicação principal."""

    def __init__(self):
        self.tela_atual = "principal"
        self.windows = {}

    def criar_janela(self, tela_nome: str):
        """Cria uma janela baseada na tela selecionada."""

        if tela_nome == "principal":
            tela = TelaPrincipal()
        elif tela_nome == "individual":
            tela = TelaPesquisaIndividual()
        elif tela_nome == "massa":
            tela = TelaPesquisaMassa()
        else:
            return None

        layout = tela.criar_layout()

        size = WINDOW_SIZE_MASSA if tela_nome == "massa" else WINDOW_SIZE

        window = sg.Window(
            WINDOW_TITLE,
            layout,
            size=size,
            finalize=True
        )

        return window, tela

    def executar(self):
        """Loop principal da aplicação."""

        while True:
            window, tela = self.criar_janela(self.tela_atual)

            if window is None:
                break

            while True:
                event, values = window.read()

                if event == sg.WINDOW_CLOSED or event == "Sair":
                    window.close()
                    return

                elif event == "Voltar":
                    window.close()
                    self.tela_atual = "principal"
                    break

                elif self.tela_atual == "principal":
                    if event == "Pesquisa Individual":
                        window.close()
                        self.tela_atual = "individual"
                        break
                    elif event == "Pesquisa em Massa":
                        window.close()
                        self.tela_atual = "massa"
                        break

                elif self.tela_atual == "individual":
                    if event == "Validar":
                        tela.processar(values, window)

                elif self.tela_atual == "massa":
                    if event == "Iniciar Validação":
                        tela.processar(
                            values,
                            window,
                            TelaPesquisaMassa.processar_massa_thread
                        )


def main():
    """Ponto de entrada."""
    app = ValidadorVIESApp()
    app.executar()


if __name__ == "__main__":
    main()
