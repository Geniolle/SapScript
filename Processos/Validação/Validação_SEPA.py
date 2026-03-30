###################################################################################
# BLOCO 1: IMPORTAÇÕES E DEFINIÇÕES INICIAIS
###################################################################################
def executar(ambiente_cockpit=None):
    import os
    import xml.etree.ElementTree as ET
    import re
    import time

    tempo_inicio = time.time()
    caminho_pasta = r"C:\SAP Script\Processos\Validação\SEPA\Ficheiros"

    ###################################################################################
    # BLOCO 2: FUNÇÃO PARA LISTAR ÁRVORE DE FICHEIROS
    ###################################################################################
    def listar_ficheiros_em_arvore(caminho_base, prefixo=""):
        entradas = os.listdir(caminho_base)
        entradas = sorted(entradas)
        for i, nome in enumerate(entradas):
            caminho_completo = os.path.join(caminho_base, nome)
            is_last = (i == len(entradas) - 1)
            branch = "└── " if is_last else "├── "
            print(prefixo + branch + nome)
            if os.path.isdir(caminho_completo):
                sub_prefixo = prefixo + ("    " if is_last else "│   ")
                listar_ficheiros_em_arvore(caminho_completo, sub_prefixo)

    print("\n📂 Estrutura da pasta Ficheiros:")
    if os.path.exists(caminho_pasta):
        listar_ficheiros_em_arvore(caminho_pasta)
    else:
        print(f"❌ Pasta não encontrada: {caminho_pasta}")
        return

    ficheiros = [f for f in os.listdir(caminho_pasta) if f.lower().endswith('.xml')]
    if not ficheiros:
        print("❌ Nenhum ficheiro XML encontrado na pasta.")
        return

    try:
        escolha = int(input("\nIndique o número do ficheiro a validar: "))
        ficheiro_escolhido = ficheiros[escolha - 1]
    except Exception as e:
        print(f"❌ Erro na seleção do ficheiro: {e}")
        return

    caminho_ficheiro = os.path.join(caminho_pasta, ficheiro_escolhido)
    print(f"\n🔍 A validar: {ficheiro_escolhido}")

    ###################################################################################
    # BLOCO 3: EXIBIR ESTRUTURA XML EM LIST TREE
    ###################################################################################
    def exibir_estrutura_xml(caminho):
        try:
            tree = ET.parse(caminho)
            root = tree.getroot()
            print(f"\n📂 Estrutura XML: {root.tag}")

            def percorrer_no(no, prefixo=""):
                filhos = list(no)
                is_last = lambda i: i == len(filhos) - 1
                for i, filho in enumerate(filhos):
                    branch = "└── " if is_last(i) else "├── "
                    texto = filho.text.strip() if filho.text and filho.text.strip() else ""
                    print(prefixo + branch + f"{filho.tag} {texto}")
                    if list(filho):
                        sub_prefixo = prefixo + ("    " if is_last(i) else "│   ")
                        percorrer_no(filho, sub_prefixo)

            percorrer_no(root)
        except Exception as e:
            print(f"❌ Erro ao ler o XML: {e}")

    exibir_estrutura_xml(caminho_ficheiro)

    ###################################################################################
    # BLOCO 4: VALIDAÇÃO DE CAMPOS OBRIGATÓRIOS
    ###################################################################################
    limite_nome = 140
    limite_iban = 34
    formato_nif = re.compile(r'^PT\d{9}$')
    formato_data = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    formato_decimal = re.compile(r'^\d+(\.\d{1,2})?$')

    erros = []

    try:
        tree = ET.parse(caminho_ficheiro)
        root = tree.getroot()

        for payment in root.findall('.//{*}PaymentInformation'):
            iban = payment.find('.//{*}IBAN')
            if iban is not None:
                if len(iban.text.strip()) > limite_iban:
                    erros.append(f"IBAN com comprimento inválido: {iban.text.strip()}")
                if not re.match(r'^[A-Z]{2}[0-9A-Z]+$', iban.text.strip()):
                    erros.append(f"Formato IBAN inválido: {iban.text.strip()}")
            else:
                erros.append("IBAN em falta.")

            end_to_end_id = payment.find('.//{*}EndToEndId')
            if end_to_end_id is None or not end_to_end_id.text.strip():
                erros.append("EndToEndId em falta ou vazio.")

            creditor_name = payment.find('.//{*}Nm')
            if creditor_name is None or not creditor_name.text.strip():
                erros.append("Creditor Name em falta.")
            elif len(creditor_name.text.strip()) > limite_nome:
                erros.append(f"Creditor Name ultrapassa {limite_nome} caracteres: {creditor_name.text.strip()}")

            tax_id = payment.find('.//{*}TaxId')
            if tax_id is not None:
                if not formato_nif.match(tax_id.text.strip()):
                    erros.append(f"Formato inválido para TaxId: {tax_id.text.strip()}")
            else:
                erros.append("TaxId em falta.")

            currency = payment.find('.//{*}Currency')
            if currency is not None:
                if not re.match(r'^[A-Z]{3}$', currency.text.strip()):
                    erros.append(f"CurrencyCode inválido: {currency.text.strip()}")

            amount = payment.find('.//{*}InstdAmt')
            if amount is not None:
                if not formato_decimal.match(amount.text.strip()):
                    erros.append(f"Formato inválido para Amount: {amount.text.strip()}")
            else:
                erros.append("InstdAmt em falta.")

            date = payment.find('.//{*}ReqdExctnDt')
            if date is not None:
                if not formato_data.match(date.text.strip()):
                    erros.append(f"Formato inválido para Data: {date.text.strip()}")
            else:
                erros.append("Data em falta.")

    except Exception as e:
        print(f"❌ Erro ao validar o XML: {e}")
        return

    ###################################################################################
    # BLOCO 5: RESULTADOS
    ###################################################################################
    if erros:
        print("\n🚨 Erros encontrados:")
        for erro in erros:
            print(f"- {erro}")
    else:
        print("\n✅ Nenhum erro encontrado. O ficheiro parece estar em conformidade com as regras básicas.")

    fim = time.time()
    m, s = divmod(int(fim - tempo_inicio), 60)
    print(f"\n⏱️ Tempo total: {m:02d}:{s:02d}")
    print("\n🔁 A voltar ao cockpit principal...")
    return True
