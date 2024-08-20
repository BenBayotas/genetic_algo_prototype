import random
import pandas as pd


# Define classes and data structures
class Subject:
    def __init__(self, code, name, time_slots, days, room_avail, instructor_avail, num_students):
        self.code = code
        self.name = name
        self.time_slots = time_slots
        self.days = days
        self.room_avail = room_avail
        self.instructor_avail = instructor_avail
        self.num_students = num_students


class Schedule:
    def __init__(self, subjects):
        self.subjects = subjects  # list of subjects
        self.schedule = {}

    def initialize(self):
        # Randomly assign each subject to a time slot and room
        for subject in self.subjects:
            available_times = subject.time_slots
            available_days = subject.days
            available_rooms = subject.room_avail

            time_slot = random.choice(available_times)
            day = random.choice(available_days)
            room = random.choice(available_rooms)

            self.schedule[subject.code] = (day, time_slot, room)

    def calculate_fitness(self):
        conflicts = 0

        # Check for conflicts in schedule
        for key1, val1 in self.schedule.items():
            for key2, val2 in self.schedule.items():
                if key1 != key2:
                    if val1[0] == val2[0] and val1[1] == val2[1]:  # same day and time slot
                        if val1[2] == val2[2]:  # same room
                            conflicts += 1

        return 1 / (1 + conflicts)  # Lower conflicts means higher fitness

    def mutate(self):
        # Randomly alter a subject's time slot or room
        subject_code = random.choice(list(self.schedule.keys()))
        subject = next((s for s in self.subjects if s.code == subject_code), None)

        available_times = subject.time_slots
        available_days = subject.days
        available_rooms = subject.room_avail

        if random.random() > 0.5:
            self.schedule[subject_code] = (
            random.choice(available_days), self.schedule[subject_code][1], random.choice(available_rooms))
        else:
            self.schedule[subject_code] = (
            self.schedule[subject_code][0], random.choice(available_times), random.choice(available_rooms))


def crossover(parent1, parent2):
    # Single point crossover
    crossover_point = random.randint(0, len(parent1.schedule) - 1)

    child1 = Schedule(parent1.subjects)
    child2 = Schedule(parent2.subjects)

    for i, (key, val) in enumerate(parent1.schedule.items()):
        if i < crossover_point:
            child1.schedule[key] = val
            child2.schedule[key] = parent2.schedule[key]
        else:
            child1.schedule[key] = parent2.schedule[key]
            child2.schedule[key] = val

    return child1, child2


def genetic_algorithm(subjects, population_size=100, generations=1000):
    population = [Schedule(subjects) for _ in range(population_size)]

    for schedule in population:
        schedule.initialize()

    for generation in range(generations):
        population = sorted(population, key=lambda x: x.calculate_fitness(), reverse=True)

        next_generation = population[:population_size // 2]

        while len(next_generation) < population_size:
            parent1 = random.choice(population[:population_size // 2])
            parent2 = random.choice(population[:population_size // 2])

            child1, child2 = crossover(parent1, parent2)

            if random.random() < 0.1:  # Mutation probability
                child1.mutate()
                child2.mutate()

            next_generation.extend([child1, child2])

        population = next_generation

    return max(population, key=lambda x: x.calculate_fitness())


def export_to_excel(schedule, filename="schedule.xlsx"):
    time_slots = [f"{hour:02d}:{minute:02d}" for hour in range(7, 21) for minute in (0, 30)]
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

    df = pd.DataFrame(index=time_slots, columns=days)

    for subject_code, (day, time_slot, room) in schedule.schedule.items():
        df.at[time_slot, day] = f"{subject_code} ({room})"

    df.to_excel(filename)
    print(f"Schedule saved to {filename}")


# Example usage
subjects = [
    Subject("CS101", "Intro to CS", ["08:00", "10:00", "14:00"], ["Monday", "Wednesday", "Friday"],
            ["CB 223", "CB 224"], ["Instructor A", "Instructor B"], 30),
    Subject("CS102", "Data Structures", ["09:00", "11:00", "13:00"], ["Tuesday", "Thursday", 'Saturday'], ["CB 226", "CB 227"],
            ["Instructor B", "Instructor C"], 25),
    # Add more subjects as needed
]

best_schedule = genetic_algorithm(subjects)
export_to_excel(best_schedule)
