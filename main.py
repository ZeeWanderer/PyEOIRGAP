import argparse
import pprint
import json
import xlsxwriter

preset = None

prettyp = pprint.PrettyPrinter(indent=4)


# pz_1 functions


def get_mul_sum_string(first: list, second: list):
    str_ = f"{' + '.join(f'({x.__round__(7)} * {y.__round__(7)})' for x, y in zip(first, second))}"
    return str_


def get_sum_string(list_: list):
    str_ = f"{' + '.join(str(x.__round__(10)) for x in list_)}"
    return str_


def compute_expet_opinions(expert_names, expert_opinions_given, arg_expnames):
    expert_opinions = []

    for list_ in expert_opinions_given:
        sublist = []
        for oplist in list_:
            suboplist = []
            assert (len(expert_names) == len(oplist))
            for ename, item in zip(expert_names, oplist):
                if ename in arg_expnames:
                    suboplist.append(item)
            sublist.append(suboplist)
        expert_opinions.append(sublist)
    return expert_opinions


def compute_category_weights(expert_names, category_weights_given, arg_expnames):
    category_weights = []

    for oplist in category_weights_given:
        suboplist = []
        assert (len(expert_names) == len(oplist))
        for ename, item in zip(expert_names, oplist):
            if ename in arg_expnames:
                suboplist.append(item)
        category_weights.append(suboplist)
    return category_weights


def compute_category_amean(category_annotation, expert_opinions):
    category_amean = []

    for category_name, subcategory_list, subcategory_oplist in zip(category_annotation.keys(),
                                                                   category_annotation.values(), expert_opinions):
        print(f"Рахуємо середнє арифметичне оцінок для категорії {category_name}:")
        category_amean_value = 0
        subcategory_amean = []
        for subcategory_name, oplist in zip(subcategory_list, subcategory_oplist):
            subcategory_amean_value = sum(oplist) / len(oplist)
            subcategory_amean.append(subcategory_amean_value)
            print(
                f"\t- {subcategory_name} - середнє арифметичне: SM = {subcategory_amean_value.__round__(5)} = ({get_sum_string(oplist)})/{len(oplist)}")
            category_amean_value += subcategory_amean_value
        category_amean.append(category_amean_value)
        print(
            f"\t-> {category_name} - середнє арифметичне: CM = {category_amean_value.__round__(5)} = ({get_sum_string(subcategory_amean)})")
    return category_amean


def compute_category_novelty_type(category_annotation, category_amean):
    category_novelty_type = []

    for category_name, subcategory_list, amean in zip(category_annotation.keys(), category_annotation.values(),
                                                      category_amean):
        print(f"Рахуємо I для категорії {category_name}:")
        N = len(subcategory_list) * 5
        retval = amean / N

        print(f"\t- значення Biomp = {amean.__round__(5)}")
        print(f"\t- значення BiMAX = {N}")
        print(
            f"\t-> I = Biomp/BiMAX = {amean.__round__(5)}/{N} = {retval.__round__(5)}{f' // novelty type not present' if retval < 0 else ''}")
        category_novelty_type.append(retval)
    return category_novelty_type


def compute_category_weights_amean(category_annotation, category_weights):
    category_weights_amean = []

    for category_name, weights_list in zip(category_annotation.keys(), category_weights):
        print(f"Рахуємо Середню Вагу для категорії {category_name}:")
        mweight = sum(weights_list) / len(weights_list)
        print(f"\t-> CMW = {mweight.__round__(10)} = ({get_sum_string(weights_list)})/{len(weights_list)}:")
        category_weights_amean.append(mweight)
    return category_weights_amean


def conclude_novelty_type(category_annotation, category_novelty_type, novelty_type_conclusion):
    for category_name, novelty_type, novelty_conclusion in zip(category_annotation.keys(), category_novelty_type,
                                                               novelty_type_conclusion):
        print(f"Висновок по I для категорії {category_name}:")
        rounded_val = novelty_type.__round__(3)
        print(f"\t- значення I = {novelty_type.__round__(7)} ≃ {rounded_val}")
        analyze_success = False
        for range_, msg in novelty_conclusion.items():
            left, right = range_
            if rounded_val > left and rounded_val <= right:
                print(f"\t-> I знаходиться у проміжку ({left}, {right}], висновок : \"{msg}\"")
                analyze_success = True
                break
        if not analyze_success:
            print(f"\t-> ERROR")


def print_annotation(category_annotation: dict, expert_opinions: list, category_weights: list, expert_names: list):
    for (category, sublist), opinions_list, weights_list in zip(category_annotation.items(), expert_opinions,
                                                                category_weights):
        print(f"{category} - максимальна кількість балів: {len(sublist) * 5}")
        for subcat, opinions in zip(sublist, opinions_list):
            opinion_str = "".join([f"{name}: {value}, " for name, value in zip(expert_names, opinions)])
            opinion_str = opinion_str[0:len(opinion_str) - 2]
            weight_str = "".join([f"{name}: {value}, " for name, value in zip(expert_names, weights_list)])
            weight_str = weight_str[0:len(weight_str) - 2]
            print(f"\t- {subcat} - Оцінки Експертів: {opinion_str}; Ваги Надані експертами: {weight_str}")


def pz_1():
    global preset
    print("Практична Робота #1:")
    print("Вхідні Дані:")
    arg_expnames = "АБВГД"

    if preset == 'max':
        arg_expnames = "ГДЖКМ"
    if preset == 'sergey':
        arg_expnames = "ВГДИМ"

    category_annotation = {
        "Споживча новизна": ["1. Зміна поведінкових звичок споживача",
                             "2. Ступінь задоволення потреб і запитів",
                             "3. Спосіб задоволення потреби",
                             "4. Формування нової потреби",
                             "5. Формування нового споживача"],
        "Товарна новизна": [  # "1. Параметричні зміни показників продукції",
            "1.1. Якісні",
            "1.2. Технічні",
            "1.3. Економічні",
            "1.4. Сервісні",
            "2. Якість продукції по відношенню до конкурентів",
            "3. Функціональні зміни"],
        "Виробнича новизна": ["1. Рівень унікальності товару для підприємства",
                              "2. Рівень унікальності для галузі",
                              "3. Рівень унікальності товару для країни",
                              "4. Зміна виробничої системи",
                              "5. Відносно існуючого асортименту"],
        "Прогресивна новизна": ["1. Зміна технологій виготовлення",
                                "2. Рівень застосування нових компонентів і матеріалів",
                                "3. зміна технологічного принципу дії виробу",
                                "4. Зміна конструктивного виконання",
                                "5. Рівень застосування інновацій"],
        "Ринкова новизна": ["1. Новий виріб на новому ринку",
                            "2. Новий виріб на відомому ринку",
                            "3. Модернізований виріб",
                            "4. Нова модель"],
        "Екологічна новизна": ["1. Рівень екологічної чистоти технології виробництва",
                               "2. Рівень впровадження мало- та безвідходних технологій",
                               "3. Рівень екологічно небезпечних режимів експлуатації продукції",
                               "4. Рівень забруднення навколишнього середовища"],
        "Соціальна новизна": ["1. Використання нового товару приводить до покращення стану здоров'я нації",
                              "2. Використання нового товару приводить до зростання доходів населення",
                              "3. Виробництво нового товару приводить до збільшення(зменшення) кількості робочих місць на підприємстві",
                              "4. Виробництво нового товару приводить до підвищення кваліфікації персоналу"],
        "Маркетингова новизна": ["1. Нові методи маркетингових досліджень",
                                 "2. Вживання нових стратегій сегментації ринку",
                                 "3. Вибір нової маркетингової стратегії обхвату і розвитку цільового сегмента",
                                 "4. Побудова нових каналів збуту"]
    }
    expert_names = ["А", "Б", "В", "Г", "Д", "Е", "Ж", "И", "К", "Л", "М", "Н", "П", "Р"]
    expert_opinions_given = [
        [
            [5, 4, 4, 0, 2, 1, 3, 3, 3, 5, 3, 2, -1, -1],
            [5, 5, -1, 1, 2, 0, 1, 5, 4, 2, 3, 0, 0, 2],
            [1, 1, 1, 4, 1, 4, 3, 4, 5, 3, -1, 5, 4, 4],
            [5, 3, 3, 0, 1, 1, 2, 2, 0, 3, 2, 5, 4, 5],
            [1, 1, 4, 0, 0, 4, 1, 0, 4, 5, -1, 0, 1, -1]
        ],
        [
            [4, 4, -1, 4, 0, 2, 5, 0, 2, 3, 4, 2, 1, 2],
            [5, 0, 0, 1, 1, 2, 2, 2, 0, 4, 0, 3, -1, 4],
            [-1, -1, 1, 5, 5, -1, 1, 1, 3, 5, 4, 1, 1, 1],
            [3, 5, 4, -1, 0, 4, -1, 2, 2, 2, 2, 5, -1, 1],
            [4, 5, -1, -1, 1, 4, 0, -1, 0, 5, 1, 0, 0, 1],
            [0, 2, 0, 1, 2, 5, 1, 1, 1, 3, 0, -1, 4, 4]
        ],
        [
            [1, 0, 1, -1, 0, 5, 1, 5, 5, 0, 2, 1, 2, 2],
            [3, 0, 5, 4, 2, 2, 0, 5, 0, 4, 0, 5, -1, 1],
            [4, 4, 3, 1, 1, 0, 3, 5, 0, 4, 3, 2, 2, 0],
            [2, 0, 4, 0, 4, 1, 4, 2, -1, -1, 1, 2, 5, 1],
            [-1, 3, 0, 3, 1, 5, 5, -1, 5, 2, 0, -1, 3, 1]
        ],
        [
            [2, 0, 4, 4, 1, 2, 5, 2, -1, 0, 5, 0, 5, 5],
            [-1, 2, 1, 0, 4, 1, 0, 2, 2, 5, 2, 5, 2, 3],
            [1, 2, 5, -1, -1, 1, 0, 2, 1, 2, -1, 1, 2, 4],
            [3, -1, 1, 4, 1, -1, 1, 2, 1, 3, 1, 4, -1, 3],
            [0, 2, 0, 0, 3, 0, 4, 2, 4, 4, -1, 4, 0, 2],
        ],
        [
            [1, -1, 5, 4, 5, 0, 1, 3, 0, 0, 1, -1, 2, 0],
            [4, 0, 5, 2, 3, 2, 1, 2, -1, 1, 4, -1, 2, 5],
            [2, 3, -1, 2, 5, 3, 4, 5, 4, 5, 3, 5, 5, 3],
            [0, 3, 5, 2, 3, 0, 3, -1, 1, 5, 4, 4, 5, 4]
        ],
        [
            [4, 5, 5, 4, 0, 0, -1, 2, 1, 4, 5, 0, 1, 5],
            [4, 1, 0, 4, 1, -1, 3, 0, 4, 0, 1, -1, 5, -1],
            [4, 4, 0, 0, 1, 1, 2, -1, 2, 0, 4, 1, 3, -1],
            [3, 4, 1, -1, -1, 4, -1, 0, 3, 5, 3, 5, -1, 3]
        ],
        [
            [-1, -1, 3, -1, 5, 2, 1, 5, 3, -1, 1, -1, 4, 5],
            [1, 1, 0, 0, -1, 0, 2, 4, 0, 5, 2, 4, 1, 2],
            [1, 3, 2, 0, 3, 2, 2, -1, 1, -1, 1, 4, 5, 4],
            [-1, -1, 3, 1, 1, 0, 0, 2, 1, -1, 4, 1, 2, 4]
        ],
        [
            [-1, 3, 3, 3, 0, 5, 5, 3, 5, 2, -1, 1, 4, -1],
            [3, 2, 4, 2, 3, 4, 4, 0, 0, 2, 0, 4, -1, 5],
            [0, 1, 3, 3, 0, 4, 0, 1, 3, 5, 4, 4, 3, 4],
            [1, 4, 4, 4, 2, 1, 5, 2, 3, 0, 5, 5, 5, 0]
        ]
    ]

    category_weights_given = [
        [0.25, 0.296, 0.224, 0.31, 0.223, 0.274, 0.215, 0.245, 0.255, 0.248, 0.301, 0.244, 0.241, 0.243],
        [0.214, 0.251, 0.243, 0.221, 0.241, 0.204, 0.248, 0.208, 0.223, 0.204, 0.201, 0.211, 0.214, 0.224],
        [0.036, 0.024, 0.035, 0.041, 0.029, 0.031, 0.021, 0.033, 0.024, 0.034, 0.041, 0.032, 0.033, 0.05],
        [0.179, 0.169, 0.153, 0.234, 0.181, 0.164, 0.154, 0.174, 0.181, 0.178, 0.224, 0.117, 0.184, 0.153],
        [0.107, 0.098, 0.074, 0.052, 0.115, 0.101, 0.079, 0.101, 0.107, 0.104, 0.051, 0.101, 0.111, 0.045],
        [0.035, 0.015, 0.05, 0.021, 0.03, 0.031, 0.048, 0.041, 0.031, 0.031, 0.03, 0.031, 0.039, 0.074],
        [0.036, 0.032, 0.045, 0.009, 0.024, 0.044, 0.041, 0.031, 0.05, 0.034, 0.01, 0.031, 0.035, 0.035],
        [0.143, 0.115, 0.176, 0.112, 0.157, 0.151, 0.194, 0.194, 0.167, 0.129, 0.142, 0.173, 0.143, 0.176]
    ]

    # with open("annotation.json", "w") as json_file:
    #     json.dump(category_annotation, json_file)
    #
    # with open("expert_names.json", "w") as json_file:
    #     json.dump(expert_names, json_file)
    #
    # with open("expert_opinions.json", "w") as json_file:
    #     json.dump(expert_opinions_given, json_file)
    #
    # with open("category_weights.json", "w") as json_file:
    #     json.dump(category_weights_given, json_file)

    novelty_type_conclusion = [
        {  # Споживча
            (0.0, 0.25): "Товари нової сфери використання",
            (0.25, 0.50): "Товари, що змінюють спосіб задоволення існуючих потреб",
            (0.50, 0.75): "Товари, що більш ефективно задовольняють існуючі потреби споживачів",
            (0.75, 1.0): "Товари, що приводять до формування нової потреби споживачів"
        },
        {  # Товарна
            (0.0, 0.25): "Товар-дублікат",
            (0.25, 0.50): "Оновлений товар",
            (0.50, 0.75): "Наступне покоління товару споживачі",
            (0.75, 1.0): "Абсолютно новий товар"
        },
        {  # Виробнича
            (0.0, 0.25): "Копія продукту конкурента, що вписується в асортимент продуктів аналогічної номенклатури",
            (0.25, 0.50): "Копія продукту конкурента, що не вписується в асортимент продуктів аналогічної номенклатури",
            (0.50, 0.75): "По-справжньому новий товар в межах існуючого асортименту",
            (0.75, 1.0): "По-справжньому новий товар без прив'язки до існуючого асортименту"
        },
        {  # Прогресивна
            (0.0, 0.25): "Недосконалий продукт",
            (0.25, 0.50): "Технічно вдосконалений продукт",
            (0.50, 0.75): "Заново створений продукт",
            (0.75, 1.0): "Піонерний товар"
        },
        {  # Ринкова
            (0.0, 0.25): "Товар з маркетинговими ноу-хау ",
            (0.25, 0.50): "Товар з новою картою",
            (0.50, 0.75): "Змінений товар",
            (0.75, 1.0): "Товар ринкової новизни"
        },
        {  # Екологічна
            (0.0, 0.25): "Екологічно небезпечні нові товари",
            (0.25, 0.50): "Екологічно прийняті нові товари",
            (0.50, 0.75): "Екологічно нейтральні нові товари",
            (0.75, 1.0): "Екологічно спрямовані нові товари"
        },
        {  # Соціальна
            (0.0, 0.25): "Соціально неповноцінні нові товари",
            (0.25, 0.50): "Нові товари, що приносять задоволення",
            (0.50, 0.75): "Соціально корисні нові товари",
            (0.75, 1.0): "Соціально бажані нові товари"
        },
        {  # Маркетингова
            (0.0, 0.25): "Неприйнятний товар",
            (0.25, 0.50): "Товар з маркетинговим оновленням",
            (0.50, 0.75): "Товарне оновлення",
            (0.75, 1.0): "Товар маркетингової новизни"
        }
    ]

    integral_novelty_conclusions = {
        (0.00, 0.19): ("Помилкова", "Малоістотна модифікація", "Новий товар"),
        (0.19, 0.39): ("Незначна", "Кардинальна зміна параметрів", "Новий товар"),
        (0.39, 0.59): ("Достатня", "Принципова технологічна модифікація товару", "Інноваційний товар"),
        (0.59, 0.79): ("Значуща", "Принципова зміна споживчих властивостей товару", "Інноваційний товар"),
        (0.79, 0.99): ("Висока", "Товар, який не має аналогів", "Інноваційний товар"),
        (0.99, 1.00): ("Найвища", "Абсолютно новий товар", "Інноваційний товар")
    }

    category_weights = compute_category_weights(expert_names, category_weights_given, arg_expnames)

    expert_opinions = compute_expet_opinions(expert_names, expert_opinions_given, arg_expnames)

    print(f"Імена експертів: {'-'.join(x for x in arg_expnames)}\n")
    print("Список категорій та підкатегорій з відповідними ім значеннями:")
    print_annotation(category_annotation, expert_opinions, category_weights, [x for x in arg_expnames])
    print()

    # print(f"Оцінки експертів відповыдно до вказаних категорій:")
    # print(expert_opinions)
    # print()
    # print(f"Ваги категорій надані єкспертами:")
    # print(category_weights)

    print()
    print("#1 Розрахунок:")

    print("#1.1 Розрахуємо середнє арифметичне оцінок експертів для кожної категорії:")

    category_amean = compute_category_amean(category_annotation, expert_opinions)

    print()
    print("#1.2 Розрахуємо тип новизни I для кожної категорії:")

    category_novelty_type = compute_category_novelty_type(category_annotation, category_amean)

    print()
    print("#1.3 Розрахуємо середнє арифметичне ваг наданих експертами для кожної категорії:")

    category_weights_amean = compute_category_weights_amean(category_annotation, category_weights)

    print()
    print("#1.4 Розрахуємо Інтегральну Новизну Nint:")

    integral_novelty = sum([x * y for x, y in zip(category_weights_amean, category_novelty_type)])

    print(
        f"\t-> Nint = sum(Wi*Ii) = ({get_mul_sum_string(category_weights_amean, category_novelty_type)}) = {integral_novelty.__round__(10)}")

    # for in zip(category_annotation.keys(), category_weights)

    print()
    print("#2 Висновки:")
    print("#2.1 Висновки по Типу Новизни I:")

    conclude_novelty_type(category_annotation, category_novelty_type, novelty_type_conclusion)

    print()
    print("#2.2 Висновки по Інтегральній Новизні Nint:")

    rounded_val = integral_novelty.__round__(3)
    print(f"\t- значення Nint = {integral_novelty.__round__(7)} ≃ {rounded_val}")
    for range_, msg in integral_novelty_conclusions.items():
        left, right = range_
        level, char, type_ = msg
        if left < rounded_val <= right:
            print(
                f"\t-> Nint знаходиться у проміжку ({left}, {right}], висновок : рівень - \"{level}\", характеристика - \"{char}\", тип - \"{type_}\"")
            break


def pz_2():
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


def compute_am_deductions(usage_time_in_months, balance, years):
    val = balance / years / 12 * usage_time_in_months
    return val


def compute_nm_am_deductions(usage_time_in_months, balance, norm):
    g = balance * (usage_time_in_months / 12) * norm
    return g


def compute_dev_use_time(usage_time_in_months, balance, years):
    val = balance / years / 12 * usage_time_in_months
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
    yearly_cost = 4.*10**6

    realization_cost_without_pdv = 1200 / 1.2

    dohid = realization_cost_without_pdv*products
    other_costs = yearly_cost - (material_cost+additional_w_pay*1.22+ammortization)*products


    podatok = (dohid - yearly_cost)*0.18

    all_costs = yearly_cost + podatok

    pure_dohid = dohid - all_costs
    t = 10


def pz_3():
    # Part 2
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
    # Part 1
    transport_coeff = 1.1
    parts_cost_data = [["Бумага для друку", 200, 24.90 / 100],
                       ["Картридж для принтера", 2, 350],
                       ["Носії для здачі копій", 2, 15]]
    parts_cost = [n * c * transport_coeff for _, n, c in parts_cost_data]
    parts_cost_sum = sum(parts_cost)

    average_days_per_month = 20
    develop_time = average_days_per_month * 2
    develop_time_months = develop_time / 40
    develop_time_hours = develop_time * 8
    # [[name, monthly, days worked total]]
    developers_costs_data = [["Инженер", 15000, develop_time],
                             ["Керівник проекту", 20000, develop_time]]
    # [[daily]]
    developers_daily = [monthly / average_days_per_month for _, monthly, days in developers_costs_data]
    # [[cost]]
    developers_cost = [monthly * days / average_days_per_month for _, monthly, days in developers_costs_data]

    developers_cost_total = sum(developers_cost)

    developers_additional_pay = developers_cost_total * additional_pay_norm
    developers_pay_accrual = (developers_additional_pay + main_costs_total) * 0.22

    developer_ammortization_deductions_data = [["ПК", 10000, 2],
                                               ["ПК", 10000, 2],
                                               ["Принтер", 6000, 2],
                                               ["Будівля", 100000, 20]]

    developer_ammortization_deductions_data_nm = [["ПЗ", 7000, 2]]

    developer_ammortization_deductions_nm = [compute_dev_nm_use_time(develop_time_months, base_cost, nonmaterial_ammortization_norm) for
                                          name, base_cost, _ in
                                          developer_ammortization_deductions_data_nm]

    developer_ammortization_deductions = [compute_dev_use_time(develop_time_months, base_cost, expl_time) for
                                          name, base_cost, expl_time in
                                          developer_ammortization_deductions_data]
    developer_ammortization_deductions_sum = sum(developer_ammortization_deductions) + sum(developer_ammortization_deductions_nm)

    developer_kpd_data = [0.87, 0.9, ]
    developer_electricity_costs_data = [["ПК", 0.2],
                                        ["ПК", 0.2],
                                        ["Лампочки", 0.1],
                                        ["Принтер", 0.1]]

    developer_electricity_costs = [kw_usage * develop_time_hours * kv_cost * kvpi / kpd for [_, kw_usage], kpd in
                                   zip(developer_electricity_costs_data, developer_kpd_data)]

    developer_electricity_costs_sum = sum(developer_electricity_costs)

    developer_other_costs = other_costs_coeff * developers_cost_total

    all_developer_costs = parts_cost_sum + developers_cost_total + developer_other_costs\
                + developer_ammortization_deductions_sum\
                + developer_electricity_costs_sum + developers_pay_accrual

    # developers_costs_data_table
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


    developers_costs_data_table_header = ["Найменування посади", "Місячний посадовий оклад, грн.", "Оплата за робочий день, грн.", "Число днів роботи ", "Витрати назаробітну плату, грн."]
    row = 0
    for col in range(0, len(developers_costs_data_table_header)):
        worksheet.write(row, col, developers_costs_data_table_header[col], format_header)

    row += 1

    for row_i, [name, monthly, days_worked_total], amount_daily, dev_cost in zip(range(0, len(developers_costs_data)), developers_costs_data, developers_daily, developers_cost):
        print(f"{row_i}, {name}, {monthly}, {days_worked_total}, {amount_daily}, {dev_cost}")
        arr_ = [name, monthly, days_worked_total, amount_daily, dev_cost]
        for col_i, val in zip(range(0, len(arr_)), arr_):
            worksheet.write(row_i + row, col_i, val, format)
    # TODO: add developers_cost_total row
    row += len(developers_costs_data)

    workbook.close()
    t = 4


def main():
    global preset

    switcher = {"pz_1": pz_1,
                "pz_2": pz_2,
                "pz_3": pz_3,
                # "pz_4": pz_4,
                }

    parser = argparse.ArgumentParser("main.py")
    parser.add_argument('pz', choices=switcher.keys())
    parser.add_argument('preset', choices=['max', 'sergey'])

    args = parser.parse_args()

    preset = args.preset

    switcher.get(args.pz)()


if __name__ == "__main__":
    main()
