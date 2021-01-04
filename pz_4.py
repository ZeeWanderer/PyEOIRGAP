from util import UserPreset, get_sum_string


def get_named_arr_sum_str(name="РічнийДохід", iStart=1, iEnd=4):
    str_ = ""
    for i in range(iStart, iEnd):
        str_ += f"{name}[{i}] + "
    str_ = str_[0:len(str_)-3]
    return str_


def pz_4(preset):
    print("Початкові дані:")
    # норматив рентабельності
    profitability_norm = 0.55

    discount_rate = 0.2
    years_percentage = [0.70, 0.2, 0.1]
    years = len(years_percentage)

    # ставка податку
    tax_rate_general = 0.20

    full_self_cost = 30.0714609396552
    relative_q_koeff = 1.183787879
    price_koeff = 0.8
    analitical_production_koeff = 0.07
    realisation_amount = 8000

    # Капіталовкладення, Кприведене
    development_cost = 217960.8070208812

    if preset == UserPreset.max:
        full_self_cost = 31.7978875224138
        relative_q_koeff = 1.9756592520954221
        price_koeff = 0.9
        analitical_production_koeff = 0.07
        realisation_amount = 4700

        # Капіталовкладення, Кприведене
        development_cost = 205726.02887125162

    print(f"НормаРентабельності = {profitability_norm}")
    print(f"СтавкаДисконту = {discount_rate}")
    print(f"СтавкаПодатку = {tax_rate_general}")
    print(f"АналітичнийКоефіцієнт = {analitical_production_koeff}")
    print(f"КількістьРеалізації = {realisation_amount}")
    print(f"ТермінРеалізації = {years} роки")
    print(f"РічнийВідсотокДоходу = {years_percentage}")

    print(f"КапіталовкладенняПриведене = {development_cost}, взято з 3 практичної")
    print(f"ПовнаСобівартість = {full_self_cost}, взято з 3 практичної")
    print(f"ВідноснийПоказникЯкості = {relative_q_koeff}, взято з 3 практичної")

    print("4.1 Визначення ціни та критичного обсягу виробництва інноваційного виробу")
    price_lower_bound = full_self_cost * (1 + profitability_norm) * (1 + tax_rate_general)
    print(
        f"НижняМежаЦіниРеалізації = {full_self_cost} * (1+{profitability_norm}) * (1+{tax_rate_general}) = {price_lower_bound}")

    price_upper_bound = price_lower_bound * relative_q_koeff
    print(
        f"ВерхняМежаЦіниРеалізації = {price_lower_bound} * {relative_q_koeff} = {price_upper_bound}")

    selected_price = (price_upper_bound - price_lower_bound) * price_koeff + price_lower_bound
    print(
        f"ДоговірнаЦіна = {selected_price}")

    selected_price_no_taxes = 100/120 * selected_price
    print(
        f"ДоговірнаЦінаБезПДВ = {selected_price_no_taxes}")

    analitical_production = (analitical_production_koeff * full_self_cost * realisation_amount)/(selected_price_no_taxes - (1-analitical_production_koeff)*full_self_cost)
    print(
        f"КритичнийОбсягВиробництва = ({analitical_production_koeff} * {full_self_cost} * {realisation_amount})/"
        f"({selected_price_no_taxes} - {(1 - analitical_production_koeff).__round__(3)} * {full_self_cost})"
        f" = {analitical_production}")

    print("4.2 Оцінювання ефективності інноваційного рішення")
    yearly_income_str = get_named_arr_sum_str()
    yearly_income = [realisation_amount * selected_price_no_taxes * x for x in years_percentage]

    for x, t in zip(years_percentage, range(1, years+1)):
        print(f"РічнийДохід[{t}] = {realisation_amount} * {selected_price_no_taxes} * {x} = {realisation_amount * selected_price_no_taxes * x}")

    for yi, t in zip(yearly_income, range(1, years+1)):
        print(f"ГП[{t}] = РічнийДохід[{t}]/(1+СтавкаДисконту)**{t} = {yi}/(1+{discount_rate})**{t} = {yi/(1+discount_rate)**t}")

    NVP_arr = [yi/(1+discount_rate)**t for yi, t in zip(yearly_income, range(1, years+1))]

    GPprivedene = NVP_arr
    GPprivedene_sum = sum(GPprivedene)
    GPprivedene_sum_str = get_sum_string(GPprivedene)
    print(f"СумаГПприв = {get_named_arr_sum_str(name='ГП')} = {GPprivedene_sum_str} = {GPprivedene_sum}")

    NVP = sum(NVP_arr) - development_cost
    NVP_sum_str = get_sum_string(NVP_arr)
    print(f"NVP = СумаГПприв - Капіталовкладення = ({NVP_sum_str}) - {development_cost} = {NVP}")
    
    profitability_index = GPprivedene_sum/development_cost
    print(f"ІндексДохідності = СумаГПприв / КапіталовкладенняПриведене = {GPprivedene_sum}/{development_cost} = {profitability_index}")

    GPprivedene_mean_average = GPprivedene_sum/years
    print(f"ГПприв_сер = СумаГПприв / ТермінРеалізації = {GPprivedene_sum}/{years} = {GPprivedene_mean_average}")
    payback_period = development_cost/GPprivedene_mean_average
    print(f"ТермінОкупності = КапіталовкладенняПриведене / ГПприв_сер = {development_cost}/{GPprivedene_mean_average} = {payback_period}")

    yearly_sum_str = get_sum_string(yearly_income)
    IRRmin = (sum(yearly_income)/development_cost)**(1/years) - 1
    print(f"IRRmin = (({yearly_income_str}) / КапіталовкладенняПриведене) ** (1/{years}) - 1 = (({yearly_sum_str})/{development_cost})**(1/{years}) - 1 = {IRRmin}")


    t = 4
