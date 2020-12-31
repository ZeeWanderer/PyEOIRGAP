import xlsxwriter

from util import get_sum_string
from util import UserPreset


def pz_4(preset):
    # норматив рентабельності
    profitability_norm = 0.55

    # ставка податку
    tax_rate_general = 0.20

    full_self_cost = 30.0714609396552

    if preset == UserPreset.max:
        full_self_cost = 31.7978875224138

    print("4.1 Визначення ціни та критичного обсягу виробництва інноваційного виробу")
    price_lower_bound = full_self_cost * (1+profitability_norm) * (1+tax_rate_general)
    print(f"НижняМежаЦіниРеалізації = {full_self_cost} * (1+{profitability_norm}) * (1+{tax_rate_general}) = {price_lower_bound}")


    t = 4