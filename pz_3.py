import xlsxwriter
from util import get_sum_string, get_mul_sum_string


def compute_am_deductions(usage_time_in_months, balance, years):
    val = balance / years / 12 * usage_time_in_months
    return val


def compute_nm_am_deductions(usage_time_in_months, balance, norm):
    g = balance * (usage_time_in_months / 12) * norm
    return g


def compute_dev_use_time(usage_time_in_months, balance, years):
    val = balance / years * usage_time_in_months / 12
    return val


def compute_dev_nm_use_time(usage_time_in_months, balance, norm):
    g = balance * (usage_time_in_months / 12) * norm
    return g


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
    # MARK: Part 2
    nonmaterial_ammortization_norm = 0.12
    other_costs_coeff = 2.0

    # [[name, time(h), work rank, tariff, grn]...]
    tariff_data = {1: 1.1, 2: 1.1, 3: 1.35, 4: 1.5, 5: 1.7, 6: 2.0, 7: 2.2, 8: 2.4}
    main_costs_data = [["Записування на носій копії ПЗ", 10.0 / 60, 3, 29.2],
                       ["Упаковування носія", 20.0 / 60, 2, 29.2]]
    main_costs = [price * time * tariff_data[rank] for name, time, rank, price in main_costs_data]

    main_costs_total = sum(main_costs)

    additional_pay_norm = 0.15

    additional_pay = main_costs_total * additional_pay_norm

    worker_pay_accrual = (additional_pay + main_costs_total) * 0.22

    time_sum = sum(list(zip(*main_costs_data))[1])
    ammortization_deductions_data = [["ПК", 10000, 2, time_sum / 8 / 20],
                                     ["Принтер", 6000, 2, time_sum / 8 / 20],
                                     ["Будівля", 100000, 20, time_sum / 8 / 20]]

    ammortization_deductions_data_nm = [["ПЗ", 7000, 2, time_sum / 8 / 20]]

    ammortization_deductions_nm = [compute_nm_am_deductions(time_usage, base_cost, nonmaterial_ammortization_norm) for
                                   name, base_cost, _, time_usage in
                                   ammortization_deductions_data_nm]

    ammortization_deductions = [compute_am_deductions(time_usage, base_cost, expl_time) for
                                name, base_cost, expl_time, time_usage in
                                ammortization_deductions_data]
    ammortization_deductions_sum = sum(ammortization_deductions) + sum(ammortization_deductions_nm)

    kvpi = 1.
    kv_cost = 0.9
    kpd_data = [0.87, 0.9, 0.9]
    electricity_costs_data = [["ПК", 0.2, time_sum],
                              ["Лампочки", 0.1, time_sum],
                              ["Принтер", 0.050, time_sum]]

    electricity_costs = [kw_usage * time * kv_cost * kvpi / kpd for [_, kw_usage, time], kpd in
                         zip(electricity_costs_data, kpd_data)]

    electricity_costs_sum = sum(electricity_costs)

    other_costs = other_costs_coeff * main_costs_total

    planned_sale_costs = .05

    production_costs = main_costs_total + additional_pay + worker_pay_accrual \
                       + ammortization_deductions_sum \
                       + electricity_costs_sum

    self_cost = production_costs * (1 + planned_sale_costs)

    # MARK: Part 1
    transport_coeff = 1.1
    parts_cost_data = [["Бумага для друку", 500, 24.90 / 100],
                       ["Картридж для принтера", 2, 600],
                       ["Диск для здачі копій", 2, 30]]
    parts_cost = [n * c * transport_coeff for _, n, c in parts_cost_data]
    parts_cost_sum = sum(parts_cost)

    average_days_per_month = 20.0
    develop_time_months = 2.0
    hours_per_day = 8.0

    develop_time_days = average_days_per_month * develop_time_months
    develop_time_hours = develop_time_days * hours_per_day
    # [[name, monthly, days worked total]]
    developers_costs_data = [["Инженер", 15000, develop_time_days],
                             ["Керівник проекту", 20000, develop_time_days]]
    # [[daily]]
    developers_daily = [monthly / average_days_per_month for _, monthly, days in developers_costs_data]
    # [[cost]]
    developers_cost = [monthly * days / average_days_per_month for _, monthly, days in developers_costs_data]

    developers_cost_total = sum(developers_cost)

    developers_additional_pay = developers_cost_total * additional_pay_norm
    developers_pay_accrual = (developers_additional_pay + main_costs_total) * 0.22

    developer_ammortization_deductions_data = [["ПК", 10_000, 2, develop_time_months],
                                               ["ПК", 10_000, 2, develop_time_months],
                                               ["Принтер", 6_000, 2, (500 * 1 / 6) / 60 / 8 / 20],
                                               # 500 pages, 6 pages per minute
                                               ["Будівля", 100_000, 20, develop_time_months],
                                               ["Меблі", 22_000, 4, develop_time_months]]

    developer_ammortization_deductions_data_nm = [["ПЗ", 7_000, 2, develop_time_months]]

    developer_ammortization_deductions_nm = [
        compute_dev_nm_use_time(time_months, base_cost, nonmaterial_ammortization_norm) for
        name, base_cost, _, time_months in
        developer_ammortization_deductions_data_nm]

    developer_ammortization_deductions = [compute_dev_use_time(time_months, base_cost, expl_time) for
                                          name, base_cost, expl_time, time_months in
                                          developer_ammortization_deductions_data]
    developer_ammortization_deductions_sum = sum(developer_ammortization_deductions) + sum(
        developer_ammortization_deductions_nm)

    developer_kpd_data = [0.87, 0.87, 0.9, 0.9]
    developer_electricity_costs_data = [["ПК", 0.2, develop_time_hours],
                                        ["ПК", 0.2, develop_time_hours],
                                        ["Лампочки", 0.1, develop_time_hours / 2],
                                        ["Принтер", 0.1, (500 * 1 / 6) / 60]]  # 500 pages, 6 pages per minute

    developer_electricity_costs = [kw_usage * time_hours * kv_cost * kvpi / kpd for [_, kw_usage, time_hours], kpd in
                                   zip(developer_electricity_costs_data, developer_kpd_data)]

    developer_electricity_costs_sum = sum(developer_electricity_costs)

    developer_other_costs = other_costs_coeff * developers_cost_total

    all_developer_costs = parts_cost_sum + developers_cost_total + developer_other_costs \
                          + developer_ammortization_deductions_sum \
                          + developer_electricity_costs_sum + developers_pay_accrual

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
        print(f"{row_i}, {name}, {monthly}, {days_worked_total}, {amount_daily}, {dev_cost}")
        arr_ = [name, monthly, days_worked_total, amount_daily, dev_cost]
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
    for row_i, [name, power, usage_time], kpd, electricity_costs in \
            zip(range(0, len(developer_electricity_costs_data)), developer_electricity_costs_data, developer_kpd_data,
                developer_electricity_costs):
        arr_ = [name, power, usage_time, electricity_costs]
        max_name_col_width = max(max_name_col_width, len(name))
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)

    row += len(developer_electricity_costs_data)
    worksheet.merge_range(row, 0, row, 2, "Всього", format)
    worksheet.write(row, 3, developer_ammortization_deductions_sum, format)

    row += default_tables_gap

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

    worksheet.set_column(0, 0, max_name_col_width + name_col_padding)

    workbook.close()
    t = 4
