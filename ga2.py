import random
import pandas as pd

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
        self.subjects = subjects
        self.schedule = {}

    def initialize(self):
        for subject in self.subjects:
            available_times = subject.time_slots
            available_rooms = subject.room_avail
            time_slot = random.choice(available_times)
            room = random.choice(available_rooms)

            for day in subject.days:
                self.schedule[(subject.code, day)] = (time_slot, room)

    def calculate_fitness(self):
        conflicts = 0
        for key1, val1 in self.schedule.items():
            for key2, val2 in self.schedule.items():
                if key1 != key2 and key1[1] == key2[1]:  # Same day
                    if val1[0] == val2[0] and val1[1] == val2[1]:  # Same time and room
                        conflicts += 1
        return 1 / (1 + conflicts)

    def mutate(self):
        subject_code, day = random.choice(list(self.schedule.keys()))
        subject = next((s for s in self.subjects if s.code == subject_code), None)
        available_times = subject.time_slots
        available_rooms = subject.room_avail
        time_slot = random.choice(available_times)
        room = random.choice(available_rooms)

        for d in subject.days:
            self.schedule[(subject_code, d)] = (time_slot, room)

def crossover(parent1, parent2):
    crossover_point = random.randint(0, len(parent1.schedule) - 1)
    child1 = Schedule(parent1.subjects)
    child2 = Schedule(parent2.subjects)

    for i, key in enumerate(parent1.schedule.keys()):
        if i < crossover_point:
            child1.schedule[key] = parent1.schedule[key]
            child2.schedule[key] = parent2.schedule[key]
        else:
            child1.schedule[key] = parent2.schedule[key]
            child2.schedule[key] = parent1.schedule[key]

    return child1, child2

def genetic_algorithm(subjects, population_size=100, generations=1000):
    population = [Schedule(subjects) for _ in range(population_size)]
    for schedule in population:
        schedule.initialize()

    for generation in range(generations):
        population = sorted(population, key=lambda x: x.calculate_fitness(), reverse=True)
        next_generation = population[:population_size//2]

        while len(next_generation) < population_size:
            parent1 = random.choice(population[:population_size//2])
            parent2 = random.choice(population[:population_size//2])
            child1, child2 = crossover(parent1, parent2)

            if random.random() < 0.1:
                child1.mutate()
                child2.mutate()

            next_generation.extend([child1, child2])

        population = next_generation

    return max(population, key=lambda x: x.calculate_fitness())

def export_to_excel(schedule, filename="schedule.xlsx"):
    time_slots = [f"{hour:02d}:{minute:02d}" for hour in range(7, 21) for minute in (0, 30)]
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

    df = pd.DataFrame(index=time_slots, columns=days)

    for (subject_code, day), (time_slot, room) in schedule.schedule.items():
        df.at[time_slot, day] = f"{subject_code} \n ({room})"

    df.to_excel(filename)
    print(f"Schedule saved to {filename}")

# Example usage
subjects = [
    Subject("CS112", "Intro to CS", ["08:00", "10:00", "14:00"], ["Monday", "Thursday"], ["CB 221", "CB 223", "CB 224", "CB 225", "CB 226", "CB 227", "CB 228"], ["Instructor A", "Instructor B"], 30),
    Subject("CS113", "Data Structures", ["09:00", "11:00", "13:00"], ["Tuesday", "Friday"], ["CB 221", "CB 223", "CB 224", "CB 225", "CB 226", "CB 227", "CB 228"], ["Instructor B", "Instructor C"], 25),
    Subject("CS114", "Programming 1", ["7:00", "9:00", "11:00"], ["Monday", "Thursday"],["CB 221", "CB 223", "CB 224", "CB 225", "CB 226", "CB 227", "CB 228"], ["Instructor C", "Instructor A"], 40),
    # Add more subjects as needed
]

best_schedule = genetic_algorithm(subjects)
export_to_excel(best_schedule)
