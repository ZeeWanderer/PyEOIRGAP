import xlsxwriter

from util import get_sum_string


def compute_am_deductions(usage_time_in_months, balance, years, name):
    val = balance / years / 12 * usage_time_in_months
    print(f"МатеріальнаАммортизація2[\"{name}\"] = {balance} / {years} / 12 * {usage_time_in_months} = {val}")
    return val


def compute_nm_am_deductions(usage_time_in_months, balance, norm, name):
    val = balance * (usage_time_in_months / 12) * norm
    print(f"НеМатеріальнаАммортизація2[\"{name}\"] = {balance} * ({usage_time_in_months} / 12) * {norm} = {val}")
    return val


def compute_dev_use_time(usage_time_in_months, balance, years, name):
    val = balance / years * usage_time_in_months / 12
    print(f"МатеріальнаАммортизація1[\"{name}\"] = {balance} / {years} * {usage_time_in_months} / 12 = {val}")
    return val


def compute_dev_nm_use_time(usage_time_in_months, balance, norm, name):
    val = balance * (usage_time_in_months / 12) * norm
    print(f"НеМатеріальнаАммортизація1[\"{name}\"] = {balance} * ({usage_time_in_months} / 12) * {norm} = {val}")
    return val


def pz_t():
    products = 8000
    realization_cost = 1200
    material_cost = 220
    additional_w_pay = 100
    ammortization = 50
    yearly_cost = 4. * 10 ** 6

    realization_cost_without_pdv = 1200 / 1.2

    dohid = realization_cost_without_pdv * products
    other_costs = yearly_cost - (material_cost + additional_w_pay * 1.22 + ammortization) * products

    podatok = (dohid - yearly_cost) * 0.18

    all_costs = yearly_cost + podatok

    pure_dohid = dohid - all_costs
    t = 10


def pz_3(preset):
    # TODO: Also output initial data
    # TODO: verify calc output is correct
    # MARK: Part 2
    nonmaterial_ammortization_norm = 0.12
    other_costs_coeff = 2.0

    # [[name, time(h), work rank, tariff, grn]...]
    tariff_data = {1: 1.1, 2: 1.1, 3: 1.35, 4: 1.5, 5: 1.7, 6: 2.0, 7: 2.2, 8: 2.4}
    main_costs_data = [["Записування на носій копії ПЗ", 10.0 / 60, 3, 29.2],
                       ["Упаковування носія", 20.0 / 60, 2, 29.2]]
    if preset == "max":
        main_costs_data = [["Записування на носій копії ПЗ", 11.0 / 60, 3, 29.2],
                           ["Упаковування носія", 21.0 / 60, 2, 29.2]]

    main_costs = [price * time * tariff_data[rank] for name, time, rank, price in main_costs_data]
    for name, time, rank, price in main_costs_data:
        print(
            f"ВартістьОперації[\"{name}\"] = {price} * {time} * {tariff_data[rank]} = {price * time * tariff_data[rank]}")

    worker_costs_total = sum(main_costs)
    worker_costs_total_str = get_sum_string(main_costs)
    print(f"ОсновнаЗПРобітників = {worker_costs_total_str} = {worker_costs_total}")

    additional_pay_norm = 0.15

    worker_additional_pay = worker_costs_total * additional_pay_norm
    print(f"ДодатковаЗПРобітників = {worker_costs_total} * {additional_pay_norm} = {worker_additional_pay}")

    worker_pay_accrual = (worker_additional_pay + worker_costs_total) * 0.22
    print(f"НарахуванняНаЗПРобітників = ({worker_additional_pay} + {worker_costs_total}) * 0.22 = {worker_pay_accrual}")

    time_sum = sum(list(zip(*main_costs_data))[1])
    ammortization_deductions_data = [["ПК", 10_000, 2, time_sum / 8 / 20],
                                     ["Принтер", 6_000, 2, time_sum / 8 / 20],
                                     ["Будівля", 100_000, 20, time_sum / 8 / 20]]

    ammortization_deductions_data_nm = [["ПЗ", 7_000, 2, time_sum / 8 / 20]]

    if preset == "max":
        ammortization_deductions_data = [["ПК", 11_000, 2, time_sum / 8 / 20],
                                         ["Принтер", 6_500, 2, time_sum / 8 / 20],
                                         ["Будівля", 110_000, 20, time_sum / 8 / 20]]

        ammortization_deductions_data_nm = [["ПЗ", 7600, 2, time_sum / 8 / 20]]

    print(f"Час використання кожного з активів (місяців): {time_sum} / 8 / 20 = { time_sum / 8 / 20}")
    ammortization_deductions_nm = [compute_nm_am_deductions(time_usage, base_cost, nonmaterial_ammortization_norm, name)
                                   for
                                   name, base_cost, _, time_usage in
                                   ammortization_deductions_data_nm]

    ammortization_deductions = [compute_am_deductions(time_usage, base_cost, expl_time, name) for
                                name, base_cost, expl_time, time_usage in
                                ammortization_deductions_data]
    ammortization_deductions_sum = sum(ammortization_deductions) + sum(ammortization_deductions_nm)
    ammortization_deductions_sum_str = get_sum_string(ammortization_deductions)
    ammortization_deductions_nm_str = get_sum_string(ammortization_deductions_nm)
    print(
        f"Аммортизація2 = ({ammortization_deductions_sum_str}) + ({ammortization_deductions_nm_str}) = {ammortization_deductions_sum}")

    kvpi = 1.
    kv_cost = 0.9
    kpd_data = [0.87, 0.9, 0.9]
    printer_usage_time_minutes = 20.0
    electricity_costs_data = [["ПК", 0.2, time_sum],
                              ["Лампочки", 0.1, time_sum / 2],
                              ["Принтер", 0.050, printer_usage_time_minutes / 60]]
    if preset == "max":
        printer_usage_time_minutes = 21.0
        electricity_costs_data = [["ПК", 0.21, time_sum],
                                  ["Лампочки", 0.1, time_sum / 2],
                                  ["Принтер", 0.050, printer_usage_time_minutes / 60]]

    print(f"Час використання кожного з активів (годин): \n"
          f"ПК - {time_sum}\n"
          f"Лампочки - {time_sum} / 2 = {time_sum / 2}\n"
          f"Принтер - {printer_usage_time_minutes} / 60 = {printer_usage_time_minutes / 60}")

    electricity_costs = [kw_usage * time * kv_cost * kvpi / kpd for [_, kw_usage, time], kpd in
                         zip(electricity_costs_data, kpd_data)]
    for [name, kw_usage, time], kpd in zip(electricity_costs_data, kpd_data):
        print(
            f"ВитратиЕлектроенергії2[\"{name}\"] = {kw_usage} * {time} * {kv_cost} * {kvpi} / {kpd} = {kw_usage * time * kv_cost * kvpi / kpd}")

    electricity_costs_sum = sum(electricity_costs)
    electricity_costs_sum_str = get_sum_string(electricity_costs)
    print(f"СумаВитратЕлектроенергії2 = {electricity_costs_sum_str} = {electricity_costs_sum}")

    other_costs = other_costs_coeff * worker_costs_total
    print(f"ІншіВитрати2 = {other_costs_coeff} * {worker_costs_total} = {other_costs}")

    planned_sale_costs_coeff = .05

    production_costs = worker_costs_total + worker_additional_pay + worker_pay_accrual + ammortization_deductions_sum \
                       + electricity_costs_sum

    production_costs_str = get_sum_string(
        [worker_costs_total, worker_additional_pay, worker_pay_accrual, ammortization_deductions_sum,
         electricity_costs_sum])
    print(f"ВитратиНаВиробництво = {production_costs_str} = {production_costs}")
    sale_costs = production_costs * planned_sale_costs_coeff
    print(f"ВитратиНаЗбут = {production_costs} * {planned_sale_costs_coeff} = {sale_costs}")

    self_cost = production_costs + sale_costs
    print(f"Собівартість = {production_costs} + {sale_costs} = {self_cost}")

    # MARK: Part 1
    transport_coeff = 1.1
    parts_cost_data = [["Бумага для друку", 500, 24.90 / 100],
                       ["Картридж для принтера", 2, 600],
                       ["Диск для здачі копій", 2, 30]]

    if preset == "max":
        parts_cost_data = [["Бумага для друку", 400, 24.90 / 100],
                           ["Картридж для принтера", 2, 650],
                           ["Диск для здачі копій", 2, 35]]

    parts_cost = [n * c * transport_coeff for _, n, c in parts_cost_data]
    for name, n, c in parts_cost_data:
        print(f"ВитратиНаКомплектуючі[\"{name}\"] = {n} * {c} * {transport_coeff} = {n * c * transport_coeff}")

    parts_cost_sum = sum(parts_cost)
    parts_cost_sum_str = get_sum_string(parts_cost)
    print(f"СумаВитратНаКомплектуючі = {parts_cost_sum_str} = {parts_cost_sum}")

    average_days_per_month = 20.0
    develop_time_months = 2.0
    hours_per_day = 8.0
    print(f"ЧасРозробки = 2 місяці")

    develop_time_days = average_days_per_month * develop_time_months
    develop_time_hours = develop_time_days * hours_per_day
    # [[name, monthly, days worked total]]
    developers_costs_data = [["Инженер", 15_000, develop_time_days],
                             ["Керівник проекту", 20_000, develop_time_days]]

    if preset == "max":
        developers_costs_data = [["Инженер", 14_000, develop_time_days],
                                 ["Керівник проекту", 19_000, develop_time_days]]

    # [[daily]]
    developers_daily = [monthly / average_days_per_month for _, monthly, days in developers_costs_data]
    for name, monthly, days in developers_costs_data:
        print(
            f"ЩоденнаЗПРозробника[\"{name}\"] = {monthly} / {average_days_per_month} = {monthly / average_days_per_month}")
    # [[cost]]
    developers_cost = [monthly * days / average_days_per_month for _, monthly, days in developers_costs_data]
    for name, monthly, days in developers_costs_data:
        print(
            f"ЗПРозробника[\"{name}\"] = {monthly} * {days} / {average_days_per_month} = {monthly * days / average_days_per_month}")

    developers_cost_total = sum(developers_cost)
    developers_cost_total_str = get_sum_string(developers_cost)
    print(f"СумаЗПРозробників = {developers_cost_total_str} = {developers_cost_total}")

    developers_additional_pay = developers_cost_total * additional_pay_norm
    print(f"ДодатковаЗПРозробників = {developers_cost_total} * {additional_pay_norm} = {developers_additional_pay}")
    developers_pay_accrual = (developers_additional_pay + worker_costs_total) * 0.22
    print(
        f"НарахуванняНаЗПРозробників = ({developers_additional_pay} + {worker_costs_total}) * 0.22 = {developers_pay_accrual}")

    developer_other_costs = other_costs_coeff * developers_cost_total
    print(f"ІншіВитрати1 = {other_costs_coeff} * {developers_cost_total} = {developer_other_costs}")

    developer_printer_usage_time_minutes = 500 * 1 / 6

    developer_ammortization_deductions_data = [["ПК", 10_000, 2, develop_time_months],
                                               ["ПК", 10_000, 2, develop_time_months],
                                               ["Принтер", 6_000, 2, developer_printer_usage_time_minutes / 60 / 8 / 20],
                                               # 500 pages, 6 pages per minute
                                               ["Будівля", 100_000, 20, develop_time_months],
                                               ["Меблі", 22_000, 4, develop_time_months]]

    developer_ammortization_deductions_data_nm = [["ПЗ", 7_000, 2, develop_time_months]]

    if preset == "max":
        developer_printer_usage_time_minutes = 400 * 1 / 6
        developer_ammortization_deductions_data = [["ПК", 11_000, 2, develop_time_months],
                                                   ["ПК", 11_000, 2, develop_time_months],
                                                   ["Принтер", 6_500, 2, developer_printer_usage_time_minutes / 60 / 8 / 20],
                                                   # 500 pages, 6 pages per minute
                                                   ["Будівля", 110_000, 20, develop_time_months],
                                                   ["Меблі", 21_000, 4, develop_time_months]]

        developer_ammortization_deductions_data_nm = [["ПЗ", 7_600, 2, develop_time_months]]

    print(f"Час використання кожного з активів (місяців): \n"
          f"ПК - {develop_time_months}\n"
          f"Принтер - {developer_printer_usage_time_minutes} / 60 / 8 / 20 = {developer_printer_usage_time_minutes / 60 / 8 / 20}\n"
          f"Будівля - {develop_time_months}\n"
          f"Меблі - {develop_time_months}")

    developer_ammortization_deductions_nm = [
        compute_dev_nm_use_time(time_months, base_cost, nonmaterial_ammortization_norm, name) for
        name, base_cost, _, time_months in
        developer_ammortization_deductions_data_nm]

    developer_ammortization_deductions = [compute_dev_use_time(time_months, base_cost, expl_time, name) for
                                          name, base_cost, expl_time, time_months in
                                          developer_ammortization_deductions_data]
    developer_ammortization_deductions_sum = sum(developer_ammortization_deductions) + sum(
        developer_ammortization_deductions_nm)
    developer_ammortization_deductions_str = get_sum_string(developer_ammortization_deductions)
    developer_ammortization_deductions_nm_str = get_sum_string(developer_ammortization_deductions_nm)
    print(
        f"Аммортизація1 = ({developer_ammortization_deductions_str}) + ({developer_ammortization_deductions_nm_str}) = {developer_ammortization_deductions_sum}")

    developer_kpd_data = [0.87, 0.87, 0.9, 0.9]

    developer_electricity_costs_data = [["ПК", 0.2, develop_time_hours],
                                        ["ПК", 0.2, develop_time_hours],
                                        ["Лампочки", 0.1, develop_time_hours / 2],
                                        ["Принтер", 0.1, developer_printer_usage_time_minutes / 60]]  # 500 pages, 6 pages per minute

    if preset == "max":
        developer_kpd_data = [0.87, 0.87, 0.9, 0.9]
        developer_electricity_costs_data = [["ПК", 0.205, develop_time_hours],
                                            ["ПК", 0.205, develop_time_hours],
                                            ["Лампочки", 0.1, develop_time_hours / 2],
                                            ["Принтер", 0.1, developer_printer_usage_time_minutes / 60]]  # 400 pages, 6 pages per minute

    print(f"Час використання кожного з активів (годин): \n"
          f"ПК - {develop_time_hours}\n"
          f"Лампочки - {develop_time_hours} / 2 = {develop_time_hours / 2}\n"
          f"Принтер - {developer_printer_usage_time_minutes} / 60 = {developer_printer_usage_time_minutes / 60}")

    developer_electricity_costs = [kw_usage * time_hours * kv_cost * kvpi / kpd for [_, kw_usage, time_hours], kpd in
                                   zip(developer_electricity_costs_data, developer_kpd_data)]

    for [name, kw_usage, time_hours], kpd in zip(developer_electricity_costs_data, developer_kpd_data):
        print(
            f"ВитратиЕлектроенергії1[\"{name}\"] = {kw_usage} * {time_hours} * {kv_cost} * {kvpi} / {kpd} = {kw_usage * time_hours * kv_cost * kvpi / kpd}")

    developer_electricity_costs_sum = sum(developer_electricity_costs)
    developer_electricity_costs_str = get_sum_string(developer_electricity_costs)
    print(f"СумаВитратЕлектроенергії1 = {developer_electricity_costs_str} = {developer_electricity_costs_sum}")

    all_developer_costs = parts_cost_sum + developers_cost_total + developer_other_costs \
                          + developer_ammortization_deductions_sum + developer_electricity_costs_sum \
                          + developers_pay_accrual

    all_developer_costs_str = get_sum_string(
        [parts_cost_sum, developers_cost_total, developer_other_costs, developer_ammortization_deductions_sum,
         developer_electricity_costs_sum, developers_pay_accrual])
    print(f"ВитратиНаРозробку = {all_developer_costs_str} = {all_developer_costs}")

    # MARK: PART 1 TABLES
    default_tables_gap = 5

    # MARK: developers_costs_data_table GENERATION
    print("INSERT DEV_COST_DATA_TABLE FILE\n")
    workbook = xlsxwriter.Workbook(f'developers_costs_data_table_{preset}.xlsx')
    worksheet = workbook.add_worksheet()

    format = workbook.add_format()
    format_header = workbook.add_format()

    format.set_border()
    format.set_align('center')
    format.set_align('vcenter')

    format_header.set_border()
    format_header.set_align('center')
    format_header.set_align('vcenter')
    format_header.set_text_wrap()

    developers_costs_data_table_header = ["Найменування посади", "Місячний посадовий оклад, грн.",
                                          "Оплата за робочий день, грн.", "Число днів роботи ",
                                          "Витрати назаробітну плату, грн."]
    row = 0
    for col in range(0, len(developers_costs_data_table_header)):
        worksheet.write(row, col, developers_costs_data_table_header[col], format_header)

    row += 1

    for row_i, [name, monthly, days_worked_total], amount_daily, dev_cost in zip(range(0, len(developers_costs_data)),
                                                                                 developers_costs_data,
                                                                                 developers_daily, developers_cost):
        # print(f"{row_i}, {name}, {monthly}, {days_worked_total}, {amount_daily}, {dev_cost}")
        arr_ = [name, monthly, amount_daily, days_worked_total, dev_cost]
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(developers_costs_data)
    worksheet.merge_range(row, 0, row, 3, "Всього", format)
    worksheet.write(row, 4, developers_cost_total, format)

    row += default_tables_gap

    # TODO: all tables are written in one .xlsx file for now
    # MARK: developer_ammortization_deductions_data_table GENERATION
    print("INSERT DEV_AMORTISATION_DEDUCTIONS_DATA_TABLE FILE\n")
    max_name_col_width = 0
    name_col_padding = 1
    developer_ammortization_deductions_data_table = ["Найменування обладнання", "Балансова вартість, грн",
                                                     "Строк корисного використання, років",
                                                     "Термін використання обладнання, місяців",
                                                     "Амортизаційні відрахування, грн"]

    for col in range(0, len(developer_ammortization_deductions_data_table)):
        worksheet.write(row, col, developer_ammortization_deductions_data_table[col], format_header)

    row += 1

    for row_i, [name, balance_cost, usage_term, usage_time], a_deductions in \
            zip(range(0, len(developer_ammortization_deductions_data)), developer_ammortization_deductions_data,
                developer_ammortization_deductions):
        arr_ = [name, balance_cost, usage_term, usage_time, a_deductions]
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(developer_ammortization_deductions_data)
    for row_i, [nm_name, nm_balance_cost, nm_usage_term, nm_usage_time], nm_a_deductions in \
            zip(range(0, len(developer_ammortization_deductions_data_nm)), developer_ammortization_deductions_data_nm,
                developer_ammortization_deductions_nm):
        arr_ = [nm_name, nm_balance_cost, nm_usage_term, nm_usage_time, nm_a_deductions]
        max_name_col_width = max(max_name_col_width, len(nm_name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(developer_ammortization_deductions_data_nm)
    worksheet.merge_range(row, 0, row, 3, "Всього", format)
    worksheet.write(row, 4, developer_ammortization_deductions_sum, format)

    row += default_tables_gap

    # MARK: developer_electricity_costs_data_table GENERATION
    print("INSERT DEV_ELECTRICITY_COSTS_DATA_TABLE FILE\n")
    developer_electricity_costs_data_table = ["Найменування обладнання", "Встановлена потужність, кВт.",
                                              "Тривалість роботи, год.", "Сума, грн"]
    for col in range(0, len(developer_electricity_costs_data_table)):
        worksheet.write(row, col, developer_electricity_costs_data_table[col], format_header)

    row += 1
    # TODO: add the example of calculation to tables
    for row_i, [name, power, usage_time], kpd, electricity_cost in \
            zip(range(0, len(developer_electricity_costs_data)), developer_electricity_costs_data, developer_kpd_data,
                developer_electricity_costs):
        arr_ = [name, power, usage_time, electricity_cost]
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(developer_electricity_costs_data)
    worksheet.merge_range(row, 0, row, 2, "Всього", format)
    worksheet.write(row, 3, developer_electricity_costs_sum, format)

    row += default_tables_gap

    # MARK: parts_cost_data_table GENERATION
    # WTF: is total needed here?
    print("INSERT PARTS_COST_DATA_TABLE FILE\n")
    parts_cost_data_table_header = ["Найменування комплектуючих", "Кількість, шт.", "Ціна за штуку, грн", "Сума, грн"]
    for col in range(0, len(parts_cost_data_table_header)):
        worksheet.write(row, col, parts_cost_data_table_header[col], format_header)

    row += 1
    table_total = 0
    for row_i, [name, units, price_per_unit] in \
            zip(range(0, len(parts_cost_data)), parts_cost_data):
        cost = units * price_per_unit
        arr_ = [name, units, price_per_unit, cost]
        table_total += cost
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(parts_cost_data)
    worksheet.merge_range(row, 0, row, 2, "Всього", format)
    worksheet.write(row, 3, table_total, format)

    row += default_tables_gap

    # MARK: PART 2 TABLES
    worksheet.write(row, 0, "PART 2 TABLES", format)
    row += default_tables_gap

    # MARK: workers_costs_data_table GENERATION
    print("INSERT WORKERS_COST_DATA_TABLE FILE\n")
    workers_costs_data_table_header = ["Найменування робіт",
                                       "Тривалість операції, год.",
                                       "Розряд роботи", "Тарифний коефіцієнт",
                                       "Погодинна тарифна ставка, грн",
                                       "Величина оплати на робітника грн"]

    for col in range(0, len(workers_costs_data_table_header)):
        worksheet.write(row, col, workers_costs_data_table_header[col], format_header)

    row += 1

    for row_i, [name, time, rank, price], worker_cost in zip(range(0, len(main_costs_data)), main_costs_data,
                                                             main_costs):
        arr_ = [name, time, rank, tariff_data[rank], price, worker_cost]
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(main_costs_data)
    worksheet.merge_range(row, 0, row, 4, "Всього", format)
    worksheet.write(row, 5, worker_costs_total, format)

    row += default_tables_gap

    # MARK: workers_costs_data_table GENERATION
    print("INSERT WORKERS_AMORTISATION_DEDUCTIONS_DATA_TABLE FILE\n")
    workers_ammortization_deductions_data_table = ["Найменування обладнання", "Балансова вартість, грн",
                                                   "Строк корисного використання, років",
                                                   "Термін використання обладнання, місяців",
                                                   "Амортизаційні відрахування, грн"]

    for col in range(0, len(workers_ammortization_deductions_data_table)):
        worksheet.write(row, col, workers_ammortization_deductions_data_table[col], format_header)

    row += 1

    for row_i, [name, balance_cost, usage_term, usage_time], a_deductions in \
            zip(range(0, len(ammortization_deductions_data)), ammortization_deductions_data,
                ammortization_deductions):
        arr_ = [name, balance_cost, usage_term, usage_time, a_deductions]
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(ammortization_deductions_data)
    for row_i, [nm_name, nm_balance_cost, nm_usage_term, nm_usage_time], nm_a_deductions in \
            zip(range(0, len(ammortization_deductions_data_nm)), ammortization_deductions_data_nm,
                ammortization_deductions_nm):
        arr_ = [nm_name, nm_balance_cost, nm_usage_term, nm_usage_time, nm_a_deductions]
        max_name_col_width = max(max_name_col_width, len(nm_name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(ammortization_deductions_data_nm)
    worksheet.merge_range(row, 0, row, 3, "Всього", format)
    worksheet.write(row, 4, ammortization_deductions_sum, format)

    row += default_tables_gap

    # TODO: some bug with table below
    # MARK: workers_electricity_costs_data_table GENERATION
    print("INSERT WORKERS_ELECTRICITY_COSTS_DATA_TABLE FILE\n")
    workers_electricity_costs_data_table = ["Найменування обладнання", "Встановлена потужність, кВт.",
                                            "Тривалість роботи, год.", "Сума, грн"]
    for col in range(0, len(workers_electricity_costs_data_table)):
        worksheet.write(row, col, workers_electricity_costs_data_table[col], format_header)

    row += 1
    # TODO: add the example of calculation to tables
    for row_i, [name, power, usage_time], kpd, worker_electricity_cost in \
            zip(range(0, len(electricity_costs_data)), electricity_costs_data, kpd_data,
                electricity_costs):
        arr_ = [name, power, usage_time, worker_electricity_cost]
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(electricity_costs_data)
    worksheet.merge_range(row, 0, row, 2, "Всього", format)
    worksheet.write(row, 3, electricity_costs_sum, format)

    row += default_tables_gap

    # MARK: overall_costs_table_header GENERATION
    # TODO: verify this is all correct
    print("INSERT OVERALL_COSTS FILE\n")
    overall_costs_table_header = ["Стаття витрат", "Умовне позначення", "Сума, грн", "Примітка"]
    costs_name_col = ["1.Витрати на матеріали на одиницю продукції, грн",
                      "2. Витрати на комплектуючі на одиницю продукції, грн",
                      "3. Витрати на силову електроенергію, грн",
                      "4. Витрати на основну заробітну плату робітників, грн",
                      "5. Витрати на додаткову заробітну плату робітників, грн",
                      "6. Витрати на нарахування на заробітну плату робітників, грн",
                      "7. Амортизаційні відрахування, грн",
                      "8. Інші витрати, грн",
                      "Виробнича собівартість",
                      "9. Витрати на збут, грн",
                      "Повна собівартість одиниці виробу"]
    costs_symbol_col = ["М",
                        "Кв",
                        "Ве",
                        "Зр",
                        "Здод",
                        "Зн",
                        "А",
                        "Взаг",
                        "Sв",
                        "Взб",
                        "Sп"]
    costs_n_col = [0, parts_cost_sum, electricity_costs_sum, worker_costs_total, worker_additional_pay,
                   worker_pay_accrual, ammortization_deductions_sum, other_costs, production_costs, sale_costs,
                   self_cost]
    costs_comment_col = [""] * len(costs_n_col)

    for col in range(0, len(overall_costs_table_header)):
        worksheet.write(row, col, overall_costs_table_header[col], format_header)

    row += 1
    for row_i, val_arr in zip(range(0, len(costs_name_col)),
                              zip(*[costs_name_col, costs_symbol_col, costs_n_col, costs_comment_col])):
        for col_i, val in zip(range(0, len(val_arr)), val_arr):
            max_name_col_width = max(max_name_col_width, len(val) / 2) if isinstance(val, str) else max_name_col_width
            worksheet.write(row + row_i, col_i, val, format_header if isinstance(val, str) else format)

    row += len(overall_costs_table_header)
    row += default_tables_gap

    worksheet.set_column(0, 0, max_name_col_width + name_col_padding)

    workbook.close()
    t = 4
