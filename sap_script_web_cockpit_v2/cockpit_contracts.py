###################################################################################
# Contratos tipados de entrada e saida do SAP Cockpit (Terminal e Worker Web)
###################################################################################

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping


class CockpitContractError(ValueError):
    """Payload do cockpit fora do contrato."""


@dataclass(frozen=True)
class CockpitRunRequest:
    processo: str
    subprocesso: str
    payload: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any] | None) -> "CockpitRunRequest":
        if payload is not None and not isinstance(payload, Mapping):
            raise CockpitContractError("Payload do cockpit deve ser um objeto.")
        dados = dict(payload or {})
        processo = str(dados.get("processo") or "").strip()
        if not processo:
            raise CockpitContractError("Processo/pasta nao informado no payload.")
        subprocesso = str(dados.get("subprocesso") or dados.get("script") or "").strip()
        return cls(processo=processo, subprocesso=subprocesso, payload=dados)


@dataclass(frozen=True)
class CockpitRunResult:
    status: str
    log: str

    def to_dict(self) -> dict[str, str]:
        return {"status": self.status, "log": self.log}
