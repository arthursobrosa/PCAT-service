from typing import Literal
import os
from .distributor_data import _load_acronyms

def create_all_folders():
    _create_folders("Concessionária")
    _create_folders("Permissionária")


def _create_folders(agent: Literal["Concessionária", "Permissionária"]):
    base_path = os.path.join(os.path.dirname(__file__), "../../")
    base_path = os.path.abspath(base_path)

    distributors_path = os.path.join(base_path, f"{agent}s")
    distributors = _load_acronyms(agent=agent)

    for distributor in distributors:
        distributor_path = os.path.join(distributors_path, distributor)
        os.makedirs(distributor_path, exist_ok=True)

        for tariff_process in ["Ajuste EER ANGRA III", "Liminar abrace", "Reajuste", "Revisão", "Revisão Extraordinária", "Tarifas Iniciais"]:
            tariff_path = os.path.join(distributor_path, tariff_process)
            os.makedirs(tariff_path)