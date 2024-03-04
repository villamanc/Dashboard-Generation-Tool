import random


def list_nums(a, b):
    while True:
        numbers = []
        for val in range(0,b):
            num = random.randint(0,a)
            numbers.append(num)
        if sum(numbers) == a:
            return numbers