import xlsxwriter
from util import get_sum_string, get_mul_sum_string


def pz_2(preset):
    parameters = ["Час компіляції (Fibonacci) [μs]", "Час виконання (Fibonacci, 100000) [μs]",
                  "Обсяг програми (Fibonacci) [byte]", "Підтримка архітектур [шт]",
                  "Кількість базових типів даних [шт]"]
    expert_weight_opinins = [[0.17, 0.3, 0.1, 0.25, 0.18],  # Nikiforova
                             [0.16, 0.31, 0.09, 0.26, 0.18],
                             [0.18, 0.29, 0.11, 0.25, 0.17],
                             [0.2, 0.3, 0.1, 0.23, 0.17],
                             [0.21, 0.31, 0.07, 0.25, 0.16],
                             [0.15, 0.31, 0.12, 0.24, 0.18],
                             [0.14, 0.29, 0.12, 0.28, 0.17],
                             [0.16, 0.36, 0.09, 0.23, 0.16],
                             [0.16, 0.26, 0.14, 0.26, 0.18],
                             [0.15, 0.28, 0.1, 0.23, 0.24],
                             [0.15, 0.32, 0.13, 0.23, 0.17]]
    own_opinion = [0.60, 0.90, 0.75, 0.68, 0.30]  # TODO

    # [(cmp, own, better),...]
    comparison = [((8898 / 3).__round__(0), 8898, -1), ((94 * 4.2).__round__(0), 94, -1),
                  ((108 * 2.1).__round__(0), 108, -1),
                  (8, 8, 1), (1, 1, 1)]

    if preset == "sergey":
        parameters = ["", "", "", "", "", ""]
        expert_weight_opinins = [[0.05, 0.3, 0.15, 0.25, 0.1, 0.15],
                                 [0.05, 0.2, 0.2, 0.3, 0.1, 0.15],
                                 [0.06, 0.25, 0.19, 0.3, 0.1, 0.1],
                                 [0.05, 0.2, 0.2, 0.35, 0.1, 0.1],
                                 [0.06, 0.25, 0.2, 0.3, 0.1, 0.09],
                                 [0.05, 0.25, 0.2, 0.25, 0.1, 0.15],
                                 [0.04, 0.25, 0.25, 0.25, 0.1, 0.11],
                                 [0.05, 0.2, 0.2, 0.3, 0.1, 0.15],
                                 [0.05, 0.25, 0.2, 0.3, 0.1, 0.1],
                                 [0.08, 0.2, 0.15, 0.35, 0.1, 0.12],
                                 [0.07, 0.18, 0.2, 0.25, 0.15, 0.15]]
        own_opinion = [0.7, 0.68, 0.7, 0.75, 0.8, 0.8]
        comparison = [(1, 3, 1), (50, 60, -1), (15, 15, 1), (1, 1, 1),
                      (3, 5, 1), (3, 4, 1)]

    # verify expert opinions
    for expert in expert_weight_opinins:
        val = sum(expert)
        if val != 1:
            print(f"Error in expert opinions {expert} sum is not 1, {val}")
            quit()
    # compute normalized experet_opinions

    expert_weight_opinins_sum = [sum(x) for x in zip(*expert_weight_opinins)]
    print("Розрахуємо суму усіх оцінок для кожного параметра p:")
    index = 1
    for x in zip(*expert_weight_opinins):
        sum_str = get_sum_string(list(x))
        print(f"\tp{index}.sum = {sum_str} = {sum(x)}")
        index += 1
    print()

    print("Розрахуємо середню суму усіх оцінок для кожного параметра p (нормалізуємо):")
    # avgX
    expert_weight_opinins_normalized = [sum(x) / len(expert_weight_opinins) for x in zip(*expert_weight_opinins)]
    index = 1
    for x in zip(*expert_weight_opinins):
        print(f"\tp{index}.nsum = {sum(x)}/{len(expert_weight_opinins)} = {sum(x) / len(expert_weight_opinins)}")
        index += 1
    print()

    # write expert opinions excel file
    print("INSERT EXPERT_OPINIONS FILE\n")
    workbook = xlsxwriter.Workbook(f'expert_opinions_{preset}.xlsx')
    worksheet = workbook.add_worksheet()

    format = workbook.add_format()

    format.set_border()
    format.set_align('center')
    format.set_align('vcenter')

    worksheet.merge_range(0, 0, 1, 0, "Експерти", format)
    worksheet.merge_range(0, 1, 0, len(parameters), "Показники", format)
    worksheet.merge_range(0, len(parameters) + 1, 1, len(parameters) + 1, "Загально", format)

    for expert in range(1, len(expert_weight_opinins) + 1):
        worksheet.write(expert + 1, 0, expert, format)
    for param in range(1, len(parameters) + 1):
        worksheet.write(1, param, param, format)
    for row, expert in zip(range(2, len(expert_weight_opinins) + 2), expert_weight_opinins):
        for col, val in zip(range(1, len(parameters) + 2), expert):
            worksheet.write(row, col, val, format)
    for expert in range(1, len(expert_weight_opinins) + 1):
        worksheet.write(expert + 1, len(parameters) + 1, 1, format)

    worksheet.write(len(expert_weight_opinins) + 2, 0, "Середнє занчення", format)
    worksheet.set_column(len(expert_weight_opinins) + 2, 0, len("Середнє занчення"))
    for col, val in zip(range(1, len(parameters) + 2), expert_weight_opinins_normalized):
        worksheet.write(len(expert_weight_opinins) + 2, col, val, format)
        worksheet.set_column(len(expert_weight_opinins) + 2, 0, len(f"{val}"))

    workbook.close()

    # verify
    # k = sum(expert_weight_opinins_normalized)
    # if sum(expert_weight_opinins_normalized) != 1:
    #     print("verification failed")

    # rank expert opinions
    expert_w_ranks = []
    for expert in expert_weight_opinins:
        rank = 1
        ranks = []
        ranks_map = dict()
        expert_tmp = expert.copy()
        for _ in expert:
            val = max(expert_tmp)
            t = list(ranks_map.keys())
            if val not in list(ranks_map.keys()):
                ranks_map[val] = rank
                rank += 1
            expert_tmp.remove(val)

        for key in expert:
            ranks.append(ranks_map[key])
        expert_w_ranks.append(ranks)

    # sum of ranks for each parameter fro all experts
    parameter_rank_sum = [sum(x) for x in zip(*expert_w_ranks)]
    parameter_rank_sum_normalized = [sum(x) / len(expert_w_ranks) for x in zip(*expert_w_ranks)]
    print("Розрахуємо середню суму рангів R.sum:")
    print(f"\t N = {len(expert_weight_opinins)}")
    print(f"\t M = {len(parameters)}")
    mean_rank__sum = len(expert_weight_opinins) * (len(parameters) + 1) / 2
    print(f"\t R.sum = N*(M+1)/2 = {mean_rank__sum}")
    print()
    print("Рангування, delt та delt^2 наведено у наступній таблиці:")

    delt = []
    for rsum in parameter_rank_sum:
        val = rsum - mean_rank__sum
        delt.append(val)
    delt2 = [x ** 2 for x in delt]
    delt2_sum = sum(delt2)

    # write DELTA file
    print("INSERT DELTA FILE\n")
    workbook = xlsxwriter.Workbook(f'delt_{preset}.xlsx')
    worksheet = workbook.add_worksheet()

    format = workbook.add_format()
    format_f = workbook.add_format()

    format.set_border()
    format.set_align('center')
    format.set_align('vcenter')

    format_f.set_border()
    format_f.set_align('center')
    format_f.set_align('vcenter')
    format_f.set_bg_color('gray')

    worksheet.merge_range(0, 0, 1, 0, "Показник", format)
    worksheet.merge_range(0, 1, 1, 1, "Сумарний ранг", format)
    worksheet.merge_range(0, 2, 1, 2, "delt", format)
    worksheet.merge_range(0, 3, 1, 3, "delt^2", format)

    worksheet.merge_range(0, 4, 0, 4 + len(expert_weight_opinins) - 1, "Ранги за кількістю експертів", format)

    worksheet.set_column(0, 0, len("Показник"))
    worksheet.set_column(0, 1, len("Сумарний ранг") + 1)

    row_o = 2
    col_o = 0
    for param in range(0, len(parameters)):
        worksheet.write(row_o + param, col_o, param + 1, format)

    row_o = 2
    col_o = 1
    for rank_idx, rankval in zip(range(0, len(parameter_rank_sum)), parameter_rank_sum):
        worksheet.write(row_o + rank_idx, col_o, rankval, format)

    row_o = 2
    col_o = 2
    for delt_idx, deltval in zip(range(0, len(delt)), delt):
        worksheet.write(row_o + delt_idx, col_o, deltval, format)

    row_o = 2
    col_o = 3
    for delt2_idx, delt2val in zip(range(0, len(delt2)), delt2):
        worksheet.write(row_o + delt2_idx, col_o, delt2val, format)

    row_o = 1
    col_o = 4
    for expert_idx in range(0, len(expert_weight_opinins)):
        worksheet.write(row_o, col_o + expert_idx, expert_idx + 1, format_f)

    row_o = 2
    col_o = 4
    for expert_idx, expert in zip(range(0, len(expert_w_ranks)), expert_w_ranks):
        for param_idx, rankval in zip(range(0, len(expert)), expert):
            worksheet.write(row_o + param_idx, col_o + expert_idx, rankval, format)

    worksheet.write(2 + len(parameters), 3, delt2_sum, format)

    workbook.close()

    print("Розрахуємо коефіцієнт конкординації CCKoef:")
    concordination_koeff = (12 * delt2_sum) / (
            (len(expert_weight_opinins) ** 2) * (len(parameters) ** 3 - len(parameters)))
    print(
        f"\t CCKoef = (12 * {delt2_sum}) / ({len(expert_weight_opinins)}^2 * ({len(parameters)}^2 - {len(parameters)}) = {concordination_koeff}")
    if not 0.4 <= concordination_koeff <= 1:
        print("Коефіцієнт конкординації виходить за проміжок [0.4, 1.0], потрібно змінити неправильних експертів.")
        quit()
    else:
        print(f"Коефіцієнт конкординації 1.0 > {concordination_koeff} > 0.4, отже думки експертів є узгодженими.")
    print()

    print("Розрахуємо значення sigma:")
    # (Xn-avgX)^2
    print("Розрахуємо квадрат різниці ваги та середнього значення ваги по кожному"
          " параметру для кожного експерта (Xn-avgX)^2:")
    qdiff_opinion = []
    for expert, expert_idx in zip(expert_weight_opinins, range(0, len(expert_weight_opinins))):
        qdiff = []
        for xn, avgX, param_idx in zip(expert, expert_weight_opinins_normalized, range(0, len(expert_weight_opinins))):
            val = (xn - avgX) ** 2
            print(
                f"qdiff[{expert_idx}][{param_idx}] = (x[{expert_idx}][{param_idx}] - avgX)^2 = ({xn} - {avgX})^2 = {val}")
            qdiff.append(val)
        qdiff_opinion.append(qdiff)
    print()
    # sum((Xn-avgX)^2)

    print("Розрахуємо sigma^2 як середнё арифметичне занчень кожного параметра для усіх експертів:")
    sigma2 = [sum(x) / len(qdiff_opinion) for x in zip(*qdiff_opinion)]
    for x, idx in zip(zip(*qdiff_opinion), range(0, len(parameters))):
        sum_str = get_sum_string(list(x))
        print(f"sigma^2[{idx}] = ({sum_str})/{len(qdiff_opinion)} = {sum(x) / len(qdiff_opinion)}")
    print()

    print("Розрахуємо sigma:")
    sigma = [x ** (1 / 2) for x in sigma2]
    for x, idx in zip(sigma2, range(0, len(sigma2))):
        print(f"sigma[{idx}] = {x}^(1/2) = {x ** (1 / 2)}")
    print()

    print("Розрахуємо sigma%:")
    sigma_percent = [x * 100 / y for x, y in zip(sigma, expert_weight_opinins_normalized)]
    for x, y, idx in zip(sigma, expert_weight_opinins_normalized, range(0, len(sigma))):
        print(f"sigma%[{idx}] = {x} * 100 / {y} = {x * 100 / y}")
    print()

    # write SIGMA file
    print("INSERT SIGMA FILE\n")
    workbook = xlsxwriter.Workbook(f'sigma_{preset}.xlsx')
    worksheet = workbook.add_worksheet()

    format = workbook.add_format()

    format.set_border()
    format.set_align('center')
    format.set_align('vcenter')

    worksheet.write(0, 0, "Показник:", format)
    worksheet.write(1, 0, "Σ Xn", format)
    worksheet.write(2, 0, "avg X", format)
    worksheet.merge_range(3, 0, 3 + len(expert_weight_opinins) - 1, 0, "(Xn-avgX)^2", format)
    worksheet.write((3 + len(expert_weight_opinins) - 1) + 1, 0, "sigma^2", format)
    worksheet.write((3 + len(expert_weight_opinins) - 1) + 2, 0, "sigma", format)
    worksheet.write((3 + len(expert_weight_opinins) - 1) + 3, 0, "sigma%", format)
    worksheet.set_column(0, 0, len("Показник:") + 2)

    row_o = 0
    col_o = 1
    for param_idx in range(0, len(parameters)):
        worksheet.write(row_o, col_o + param_idx, param_idx + 1, format)

    row_o = 1
    col_o = 1
    for val_idx, val in zip(range(0, len(expert_weight_opinins_sum)), expert_weight_opinins_sum):
        worksheet.write(row_o, col_o + val_idx, val, format)

    row_o = 2
    col_o = 1
    for val_idx, val in zip(range(0, len(expert_weight_opinins_normalized)), expert_weight_opinins_normalized):
        worksheet.write(row_o, col_o + val_idx, val, format)

    row_o = 3
    col_o = 1
    for val_idx, valx in zip(range(0, len(qdiff_opinion)), qdiff_opinion):
        for val_idy, valy in zip(range(0, len(valx)), valx):
            worksheet.write(row_o + val_idx, col_o + val_idy, valy, format)

    row_o = (3 + len(expert_weight_opinins) - 1) + 1
    col_o = 1
    for val_idx, val in zip(range(0, len(sigma2)), sigma2):
        worksheet.write(row_o, col_o + val_idx, val, format)

    row_o = (3 + len(expert_weight_opinins) - 1) + 2
    col_o = 1
    for val_idx, val in zip(range(0, len(sigma)), sigma):
        worksheet.write(row_o, col_o + val_idx, val, format)

    row_o = (3 + len(expert_weight_opinins) - 1) + 3
    col_o = 1
    for val_idx, val in zip(range(0, len(sigma_percent)), sigma_percent):
        worksheet.write(row_o, col_o + val_idx, val, format)
        worksheet.set_column(row_o, col_o + val_idx, len(f"{val}"))
    workbook.close()

    quality_koeff = sum([x * y for x, y in zip(own_opinion, expert_weight_opinins_normalized)])
    relative_weight = [(y / x) ** z for x, y, z in comparison]

    print("Розрахуємо абсолютний коефіцієнт якості QKoef:")
    sim_mul = get_mul_sum_string(own_opinion, expert_weight_opinins_normalized)
    print(f"\tQKoef = {sim_mul} = {quality_koeff}")
    print()

    # write CMP file
    print("INSERT CMP FILE\n")
    workbook = xlsxwriter.Workbook(f'cmp_{preset}.xlsx')
    worksheet = workbook.add_worksheet()

    format = workbook.add_format()
    format_t = workbook.add_format()

    format.set_border()
    format.set_align('center')
    format.set_align('vcenter')

    format_t.set_border()
    format_t.set_align('center')
    format_t.set_align('vcenter')
    format_t.set_text_wrap()

    worksheet.merge_range(0, 0, 1, 0, "Показники", format_t)

    worksheet.merge_range(0, 1, 0, 2, "Варіанти", format)
    worksheet.write(1, 1, "Базовий", format)
    worksheet.write(1, 2, "Новий", format)

    worksheet.merge_range(0, 3, 1, 3, "Відносний показник якості", format_t)
    worksheet.merge_range(0, 4, 1, 4, "Коефіцієнт вагомості параметра", format_t)

    worksheet.set_column(0, 0, len("Показники") + 2)
    worksheet.set_column(3, 3, len("Відносний показник якості") / 2)
    worksheet.set_column(4, 4, len("Коефіцієнт вагомості параметра") / 2)

    row_o = 2
    col_o = 0
    for param_idx in range(0, len(parameters)):
        worksheet.write(row_o + param_idx, col_o, param_idx + 1, format)

    row_o = 2
    col_o = 1
    for val_idx, (base, new, st) in zip(range(0, len(comparison)), comparison):
        worksheet.write(row_o + val_idx, col_o, base, format)
        worksheet.write(row_o + val_idx, col_o + 1, new, format)

    row_o = 2
    col_o = 3
    for val_idx, val in zip(range(0, len(relative_weight)), relative_weight):
        worksheet.write(row_o + val_idx, col_o, val, format)

    row_o = 2
    col_o = 4
    for val_idx, val in zip(range(0, len(expert_weight_opinins_normalized)), expert_weight_opinins_normalized):
        worksheet.write(row_o + val_idx, col_o, val, format)

    workbook.close()

    relative_quality_koeff = sum([x * y for x, y in zip(relative_weight, expert_weight_opinins_normalized)])
    print("Розрахуємо відносний коефіцієнт якості RQKoef:")
    sim_mul = get_mul_sum_string(relative_weight, expert_weight_opinins_normalized)
    print(f"\tRQKoef = {sim_mul} = {relative_quality_koeff}")
    print()
