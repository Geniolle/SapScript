# Análise de Programas SAP

Este diretório centraliza investigações técnicas, análises de causa raiz, engenharia reversa e fontes de programas SAP (standard e customizados `Z*`).

---

## Estrutura Padrão

Cada análise técnica é organizada em uma pasta dedicada com o seguinte padrão:

```text
Analise_Programas/
└── <NOME_DO_PROCESSO_OU_PROGRAMA>/
    ├── README.md          # Documentação técnica, arquitetura, fluxo e diagnóstico
    ├── abap/              # Códigos-fonte ABAP extraídos do sistema (standard e Z)
    └── scripts/           # Scripts Python/RFC para reprodução, leitura e diagnóstico
```

---

## Catálogo de Análises

| Processo / Programa | Módulo | Objetivo Técnico | Status |
| :--- | :--- | :--- | :--- |
| [`FI_BILL_ISSUE_SPLIT_999_LINHAS`](./FI_BILL_ISSUE_SPLIT_999_LINHAS/README.md) | SD / FI | Particionamento contábil de faturas que excedem 999 linhas (`F5 727`), BAdI `FI_BILL_ISSUE_SPLIT` e classe `ZCLFI_BILL_ISSUE_SPLIT`. | Documentado |
| [`ZBIM_MONITOR_V1`](./ZBIM_MONITOR_V1/README.md) | MM / FI / WF | Monitor de revisão de faturas bloqueadas com workflow de divergências de preço e quantidade (`ZBIM_MONITOR`). | Documentado |

---

## Boas Práticas

1. **Read-Only**: Todos os scripts de teste contidos em `scripts/` devem ser estritamente de leitura (consultas RFC via `pyrfc`, `RFC_READ_TABLE`, `RPY_PROGRAM_READ`).
2. **Segurança de Credenciais**: Nunca colocar senhas ou credenciais nos scripts; usar sempre o carregamento a partir do `.env` raiz do projeto.
3. **Preservação de Fontes**: Os fontes em `abap/` devem refletir fielmente o código ativo no ambiente SAP para fins de auditoria e versionamento no Git.
