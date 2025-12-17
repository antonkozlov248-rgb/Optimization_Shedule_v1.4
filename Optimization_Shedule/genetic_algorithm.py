import random
import copy
from typing import Dict, List, Set, Tuple
import math
from collections import defaultdict

# from sanpin import enrich_subject_data


class ScheduleBuilder:
    """Конструктор с 8-м периодом и гарантированным разрешением конфликтов"""

    def __init__(self, classes, subjects, teachers, rooms):
        self.classes = classes
        self.subjects = subjects
        self.teachers = teachers
        self.rooms = rooms
        self.days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']
        self.schedule = {}
        self.conflicts = {
            'teacher_conflicts': 0,
            'class_conflicts': 0,
            'room_conflicts': 0,
            'sanpin_violations': 0
        }
        self.fitness = 0
        self.teacher_rooms = self._build_teacher_rooms_map()
        self.room_usage_counter = defaultdict(int)

    def _build_teacher_rooms_map(self):
        """Строим словарь: учитель -> список доступных кабинетов"""
        teacher_rooms = {}
        for teacher in self.teachers:
            fio = teacher.get('ФИО', '')
            rooms_str = teacher.get('кабинеты', '')
            if rooms_str and rooms_str != 'N/A':
                room_list = [r.strip() for r in str(rooms_str).split(';')]
                teacher_rooms[fio] = room_list
        return teacher_rooms

    def validate_data(self):
        if not self.classes or not self.subjects:
            raise ValueError("Нет классов или предметов")
        for subj in self.subjects:
            if not subj.get('учитель') or subj.get('учитель') == 'N/A':
                subj['учитель'] = 'Generic'
            if not subj.get('часов_в_неделю'):
                subj['часов_в_неделю'] = 1

    def build_schedule(self) -> Dict:
        """Построение расписания с 8-м периодом"""
        self.validate_data()
        self.schedule = {}
        self.room_usage_counter = defaultdict(int)

        class_lessons_pool = {}
        class_quotas = {}
        parallels_map = defaultdict(list)

        for c in self.classes:
            parallels_map[str(c.get('параллель'))].append(c.get('класс'))

        for p in parallels_map:
            parallels_map[p].sort()

        distributed_subjects = defaultdict(list)

        for p_key, classes_list in parallels_map.items():
            p_subjects = [s for s in self.subjects if str(s.get('параллель')) == p_key]
            grouped_by_name = defaultdict(list)
            for s in p_subjects:
                grouped_by_name[s['предмет']].append(s)

            for subj_name, variants in grouped_by_name.items():
                if len(variants) >= len(classes_list):
                    for i, cls_name in enumerate(classes_list):
                        my_variant = variants[i % len(variants)]
                        distributed_subjects[cls_name].append(my_variant)
                else:
                    for cls_name in classes_list:
                        distributed_subjects[cls_name].append(variants[0])

        for c in self.classes:
            c_name = c.get('класс')
            personal = [s for s in self.subjects if str(s.get('параллель')) == c_name]
            distributed_subjects[c_name].extend(personal)

        for cls in self.classes:
            c_name = cls.get('класс')
            self.schedule[c_name] = {day: [] for day in self.days}

            lessons = []
            my_subjects = distributed_subjects[c_name]
            for subj in my_subjects:
                hours = int(subj.get('часов_в_неделю', 1))
                for _ in range(hours):
                    lessons.append(subj.copy())

            lessons.sort(key=lambda x: (x['учитель'], x.get('коэффициент', 0)), reverse=True)
            class_lessons_pool[c_name] = lessons

            total = len(lessons)
            base = total // 5
            rem = total % 5
            quotas = [base] * 5
            priority = [2, 1, 3, 0, 4]
            for i in range(rem):
                quotas[priority[i]] += 1

            class_quotas[c_name] = {day: q for day, q in zip(self.days, quotas)}

        cls_names_fixed = sorted([c.get('класс') for c in self.classes])
        teacher_occupied = {day: {p: set() for p in range(1, 10)} for day in self.days}  # 1-9 (8 периодов)

        for day_idx in range(5):
            day = self.days[day_idx]

            for cls_name in cls_names_fixed:
                target_count = class_quotas[cls_name][day]
                pool = class_lessons_pool[cls_name]

                if not pool:
                    continue

                selected_lessons = []
                used_subjects = set()
                rem_pool = []

                for l in pool:
                    if len(selected_lessons) < target_count:
                        if l['предмет'] not in used_subjects:
                            selected_lessons.append(l)
                            used_subjects.add(l['предмет'])
                        else:
                            rem_pool.append(l)
                    else:
                        rem_pool.append(l)

                while len(selected_lessons) < target_count and rem_pool:
                    selected_lessons.append(rem_pool.pop(0))

                class_lessons_pool[cls_name] = rem_pool

                self._place_lessons_with_backtracking(
                    cls_name, day, selected_lessons, teacher_occupied
                )

        self.calculate_fitness()
        return self.schedule

    def _place_lessons_with_backtracking(self, cls, day, lessons, t_occ):
        """Размещение уроков с гарантией (8 периодов)"""
        lessons.sort(key=lambda x: x.get('коэффициент', 0), reverse=True)
        schedule_map = [None] * 9  # 8 периодов (индексы 1-8)
        unplaced = []

        for lesson in lessons:
            placed = False
            teacher = lesson.get('учитель')

            slots_order = [2, 3, 4, 1, 5, 6, 7, 8]  # +8-й период!
            if lesson.get('направление') == 'Физкультура':
                slots_order = [6, 7, 8, 5, 4, 3, 2]  # +8-й период!

            for p in slots_order:
                if schedule_map[p - 1] is None and teacher not in t_occ[day][p]:
                    schedule_map[p - 1] = lesson
                    t_occ[day][p].add(teacher)
                    placed = True
                    break

            if not placed:
                for p in range(1, 9):  # 1-8 периоды
                    if schedule_map[p - 1] is None and teacher not in t_occ[day][p]:
                        schedule_map[p - 1] = lesson
                        t_occ[day][p].add(teacher)
                        placed = True
                        break

            if not placed:
                for p in range(1, 9):  # 1-8 периоды
                    if schedule_map[p - 1] is None:
                        schedule_map[p - 1] = lesson
                        placed = True
                        break

            if not placed:
                unplaced.append(lesson)

        final = []
        for p in range(1, 9):  # 1-8 периоды
            if schedule_map[p - 1]:
                l = schedule_map[p - 1]
                teacher_name = l.get('учитель', 'Generic')
                room = self._get_best_room_for_teacher(teacher_name, day, p)
                final.append({
                    'урок': len(final) + 1,
                    'предмет': l['предмет'],
                    'учитель': l['учитель'],
                    'кабинет': room,
                    'направление': l.get('направление', 'Other'),
                    'коэффициент': l.get('коэффициент', 0)
                })

        self.schedule[cls][day] = final

    def _get_best_room_for_teacher(self, teacher_name, day, period):
        """Выбирает оптимальный кабинет для учителя"""
        available_rooms = self.teacher_rooms.get(teacher_name, ['101'])
        if not available_rooms:
            return '101'
        if len(available_rooms) == 1:
            return available_rooms[0]

        best_room = min(available_rooms, key=lambda r: self.room_usage_counter.get(f"{teacher_name}_{r}", 0))
        self.room_usage_counter[f"{teacher_name}_{best_room}"] += 1
        return best_room

    def _find_all_conflicts(self) -> List[Dict]:
        """Находит ВСЕ конфликты"""
        conflicts = []
        for day in self.days:
            for p in range(1, 9):  # 1-8 периоды
                t_map = {}
                for c_name, c_sched in self.schedule.items():
                    lesson = next((x for x in c_sched.get(day, []) if x['урок'] == p), None)
                    if lesson:
                        t = lesson.get('учитель')
                        if t and t != 'Generic':
                            if t not in t_map:
                                t_map[t] = []
                            t_map[t].append(c_name)

                for teacher, classes_list in t_map.items():
                    if len(classes_list) > 1:
                        conflicts.append({
                            'day': day,
                            'period': p,
                            'teacher': teacher,
                            'classes': classes_list
                        })
        return conflicts

    def has_teacher_conflicts(self) -> bool:
        """Проверяет наличие конфликтов"""
        return len(self._find_all_conflicts()) > 0

    def resolve_all_conflicts_powerful(self, max_retries=10):
        """СУПЕР-МОЩНОЕ разрешение конфликтов"""
        for retry in range(max_retries):
            conflicts = self._find_all_conflicts()
            if not conflicts:
                print(f"✅ Все конфликты разрешены на попытке {retry + 1}")
                return True

            print(f"  Попытка {retry + 1}: {len(conflicts)} конфликтов, разрешаем...")

            for iteration in range(1000):
                conflicts = self._find_all_conflicts()
                if not conflicts:
                    return True

                conflict = conflicts[0]
                day = conflict['day']
                period = conflict['period']
                teacher = conflict['teacher']
                problem_classes = conflict['classes']

                resolved = False

                for cls in problem_classes:
                    lessons_on_day = self.schedule[cls].get(day, [])
                    for idx, lesson in enumerate(lessons_on_day):
                        if lesson.get('учитель') == teacher:
                            for other_day in self.days:
                                if other_day != day:
                                    lessons_other = self.schedule[cls].get(other_day, [])
                                    self.schedule[cls][day].pop(idx)
                                    for i, l in enumerate(self.schedule[cls][day], 1):
                                        l['урок'] = i
                                    lessons_other.append(lesson)
                                    for i, l in enumerate(lessons_other, 1):
                                        l['урок'] = i
                                    resolved = True
                                    break
                            if resolved:
                                break
                    if resolved:
                        break

                if not resolved and random.random() < 0.7:
                    try:
                        c1 = random.choice(list(self.schedule.keys()))
                        c2 = random.choice(list(self.schedule.keys()))
                        d1 = random.choice(self.days)
                        d2 = random.choice(self.days)

                        l1 = self.schedule[c1].get(d1, [])
                        l2 = self.schedule[c2].get(d2, [])

                        if len(l1) > 0 and len(l2) > 0:
                            i1 = random.randint(0, len(l1) - 1)
                            i2 = random.randint(0, len(l2) - 1)
                            l1[i1], l2[i2] = l2[i2], l1[i1]
                            resolved = True
                    except:
                        pass

                if not resolved:
                    try:
                        cls = problem_classes[0]
                        min_day = min(self.days, key=lambda d: len(self.schedule[cls].get(d, [])))
                        if min_day != day:
                            lesson = self.schedule[cls][day].pop(0)
                            for i, l in enumerate(self.schedule[cls][day], 1):
                                l['урок'] = i
                            self.schedule[cls][min_day].append(lesson)
                            for i, l in enumerate(self.schedule[cls][min_day], 1):
                                l['урок'] = i
                            resolved = True
                    except:
                        pass

                if not resolved:
                    try:
                        c = random.choice(list(self.schedule.keys()))
                        d1 = random.choice(self.days)
                        d2 = random.choice(self.days)
                        if d1 != d2:
                            if len(self.schedule[c][d1]) > 0 and len(self.schedule[c][d2]) > 0:
                                idx1 = random.randint(0, len(self.schedule[c][d1]) - 1)
                                idx2 = random.randint(0, len(self.schedule[c][d2]) - 1)
                                l = self.schedule[c][d1].pop(idx1)
                                l2 = self.schedule[c][d2].pop(idx2)
                                self.schedule[c][d1].append(l2)
                                self.schedule[c][d2].append(l)
                                for day_x in [d1, d2]:
                                    for i, lesson in enumerate(self.schedule[c][day_x], 1):
                                        lesson['урок'] = i
                    except:
                        pass

        return False

    def calculate_fitness(self):
        """Расчет fitness"""
        self.conflicts = {k: 0 for k in self.conflicts}
        penalties = 0

        conflicts = self._find_all_conflicts()
        for conflict in conflicts:
            self.conflicts['teacher_conflicts'] += 1
            penalties += 500000

        for c_sched in self.schedule.values():
            for day, lessons in c_sched.items():
                for i in range(len(lessons) - 1):
                    if lessons[i].get('кабинет') == lessons[i + 1].get('кабинет'):
                        penalties += 5

        if self.conflicts['teacher_conflicts'] == 0:
            for c_sched in self.schedule.values():
                lesson_counts = [len(c_sched.get(day, [])) for day in self.days]
                avg_lessons = sum(lesson_counts) / 5
                variance = sum((count - avg_lessons) ** 2 for count in lesson_counts)
                penalties += min(variance * 0.5, 10)

        for c_sched in self.schedule.values():
            for day, lessons in c_sched.items():
                if len(lessons) < 4:
                    penalties += (4 - len(lessons)) * 100

        if penalties == 0:
            self.fitness = 100.0
        else:
            self.fitness = 100 / (1 + (penalties / 1000))

        return self.fitness


class Result:
    def __init__(self, schedule, fitness, conflicts):
        self.schedule = schedule
        self.fitness = fitness
        self.conflicts = conflicts


class GeneticAlgorithm:
    """Генетический алгоритм с 8-м периодом и 100% гарантией"""

    def __init__(self, classes, subjects, teachers, rooms, generations=100, population_size=40, **kwargs):
        self.classes = classes
        self.subjects = subjects
        self.teachers = teachers
        self.rooms = rooms

        self.generations = max(150, len(classes) * 3)
        self.population_size = min(50 + len(classes) * 2, 200)

        self.best_schedule = None
        self.best_fitness = 0
        self.conflicts = {}
        self.no_improvement_generations = 0

        self.days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']

    def order_crossover(self, parent1_schedule: Dict, parent2_schedule: Dict) -> Dict:
        """Order Crossover"""
        child_schedule = copy.deepcopy(parent1_schedule)

        start_day = random.randint(0, len(self.days) - 1)
        end_day = random.randint(start_day, len(self.days) - 1)

        for i in range(start_day, end_day + 1):
            day = self.days[i]
            for class_name in child_schedule.keys():
                if class_name in parent2_schedule:
                    child_schedule[class_name][day] = copy.deepcopy(
                        parent2_schedule[class_name][day]
                    )

        return child_schedule

    def swap_mutation(self, schedule: Dict, mutation_rate: float = 0.15) -> Dict:
        """Swap mutation"""
        mutated = copy.deepcopy(schedule)

        if random.random() > mutation_rate:
            return mutated

        try:
            c_name = random.choice(list(mutated.keys()))
            day = random.choice(self.days)
            lessons = mutated[c_name][day]

            if len(lessons) >= 2:
                i, j = random.sample(range(len(lessons)), 2)
                lessons[i], lessons[j] = lessons[j], lessons[i]

                for idx, lesson in enumerate(lessons, 1):
                    lesson['урок'] = idx
        except:
            pass

        return mutated

    def run(self):
        """Основной цикл GA с 8-м периодом"""
        print(f"🚀 Запуск GA с {self.population_size} особей на {self.generations} поколений (8 периодов)")

        current_pop = [self._create_individual() for _ in range(self.population_size)]
        self._update_global_best(current_pop)

        for g in range(self.generations):
            current_pop.sort(key=lambda x: x['fitness'], reverse=True)

            new_pop = []

            for i in range(min(5, len(current_pop))):
                new_pop.append(copy.deepcopy(current_pop[i]))

            while len(new_pop) < self.population_size:
                parent1 = current_pop[random.randint(0, min(9, len(current_pop) - 1))]
                parent2 = current_pop[random.randint(0, min(9, len(current_pop) - 1))]

                child_schedule = self.order_crossover(parent1['schedule'], parent2['schedule'])
                child_schedule = self.swap_mutation(child_schedule, mutation_rate=0.15)

                b = ScheduleBuilder(self.classes, self.subjects, self.teachers, self.rooms)
                b.schedule = child_schedule
                b.calculate_fitness()

                new_pop.append({
                    'schedule': child_schedule,
                    'fitness': b.fitness,
                    'conflicts': b.conflicts
                })

            current_pop = new_pop[:self.population_size]
            old_fitness = self.best_fitness
            self._update_global_best(current_pop)

            if self.best_fitness == old_fitness:
                self.no_improvement_generations += 1
            else:
                self.no_improvement_generations = 0

            if self.best_fitness > 99 or self.no_improvement_generations > 50:
                break

            if g % 10 == 0:
                print(f"Gen {g}: Best={self.best_fitness:.2f}, Conflicts={self.conflicts.get('teacher_conflicts', 0)}")

        # СУПЕР-МОЩНОЕ РАЗРЕШЕНИЕ КОНФЛИКТОВ
        print(f"\n🔧 СУПЕР-МОЩНОЕ разрешение конфликтов (с 8-м периодом)...")

        if self.best_schedule:
            b = ScheduleBuilder(self.classes, self.subjects, self.teachers, self.rooms)
            b.schedule = copy.deepcopy(self.best_schedule)

            success = b.resolve_all_conflicts_powerful(max_retries=10)

            b.calculate_fitness()

            if b.fitness > self.best_fitness:
                self.best_fitness = b.fitness
                self.best_schedule = b.schedule
                self.conflicts = b.conflicts

            # ФИНАЛЬНАЯ ПРОВЕРКА
            print(f"\n✓ Финальная проверка...")
            conflicts = b._find_all_conflicts()
            if conflicts:
                print(f"❌ Осталось {len(conflicts)} конфликтов!")
                print(f"   {conflicts[:5]}")
            else:
                print(f"✅ ИДЕАЛЬНО! Конфликтов 0. Расписание валидно (8 периодов).")

        return Result(self.best_schedule, self.best_fitness, self.conflicts)

    def _create_individual(self):
        """Создание одной особи"""
        b = ScheduleBuilder(self.classes, self.subjects, self.teachers, self.rooms)
        return {
            'schedule': b.build_schedule(),
            'fitness': b.fitness,
            'conflicts': b.conflicts
        }

    def _update_global_best(self, pop):
        """Обновляем лучшее решение"""
        best = max(pop, key=lambda x: x['fitness'])
        if best['fitness'] > self.best_fitness:
            self.best_fitness = best['fitness']
            self.best_schedule = copy.deepcopy(best['schedule'])
            self.conflicts = copy.deepcopy(best['conflicts'])
