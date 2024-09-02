import random
import pandas as pd


# Define classes and data structures
class Subject:
    def __init__(self, code, name, time_slots, days, room_avail, instructor_avail, num_students):
        self.code = code
        self.name = name
        self.time_slots = time_slots  # Time slots in 1.5-hour blocks
        self.days = days  # List of days subject is held
        self.room_avail = room_avail  # Available rooms
        self.instructor_avail = instructor_avail  # Available instructors
        self.num_students = num_students  # Number of students


class Schedule:
    def __init__(self, subjects):
        self.subjects = subjects  # list of subjects
        self.schedule = {}

    def initialize(self):
        # Randomly assign each subject to a time slot and room for each day it is held
        for subject in self.subjects:
            subject_schedule = []
            for day in subject.days:
                time_slot = random.choice(subject.time_slots)
                room = random.choice(subject.room_avail)
                subject_schedule.append((day, time_slot, room))
            self.schedule[subject.code] = subject_schedule

    def calculate_fitness(self):
        conflicts = 0

        # Check for conflicts in schedule
        time_room_assignments = {}

        for subject_code, subject_schedule in self.schedule.items():
            for day, time_slot, room in subject_schedule:
                key = (day, time_slot, room)
                if key in time_room_assignments:
                    conflicts += 1
                else:
                    time_room_assignments[key] = subject_code

        return 1 / (1 + conflicts)  # Lower conflicts mean higher fitness

    def mutate(self):
        # Randomly alter a subject's time slot or room for one of the days it is held
        subject_code = random.choice(list(self.schedule.keys()))
        subject = next((s for s in self.subjects if s.code == subject_code), None)

        day_index = random.randint(0, len(subject.days) - 1)
        new_time_slot = random.choice(subject.time_slots)
        new_room = random.choice(subject.room_avail)

        self.schedule[subject_code][day_index] = (subject.days[day_index], new_time_slot, new_room)


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
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

    df = pd.DataFrame(index=time_slots, columns=days)

    for subject_code, subject_schedule in schedule.schedule.items():
        for day, time_slot, room in subject_schedule:
            df.at[time_slot, day] = f"{subject_code} ({room})"

    df.to_excel(filename)
    print(f"Schedule saved to {filename}")


# Example usage
subjects = [
    Subject("CS101", "Intro to CS", ["08:00", "10:00", "14:00"], ["Monday", "Thursday"], ["Room 101", "Room 102"],
            ["Instructor A", "Instructor B"], 30),
    Subject("CS102", "Data Structures", ["09:00", "11:00", "13:00"], ["Tuesday", "Friday"], ["Room 103", "Room 104"],
            ["Instructor B", "Instructor C"], 25),
    Subject("CS103", "Algorithms", ["07:30", "09:00", "15:00"], ["Wednesday"], ["Room 105", "Room 106"],
            ["Instructor A", "Instructor D"], 20),
    # Add more subjects as needed
]

best_schedule = genetic_algorithm(subjects)
export_to_excel(best_schedule)
