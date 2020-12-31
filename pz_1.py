from util import get_sum_string, get_mul_sum_string, UserPreset


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


def pz_1(preset: UserPreset):
    print("Практична Робота #1:")
    print("Вхідні Дані:")
    arg_expnames = "АБВГД"

    if preset == UserPreset.max:
        arg_expnames = "ГДЖКМ"
    if preset == UserPreset.sergey:
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
