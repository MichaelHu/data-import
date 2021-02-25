import math
import random

def random_minus_equation(max_num):
    max_num -= 1
    left = math.ceil(random.random() * max_num)
    right = math.ceil(random.random() * left)
    print("{:2d} - {:2d} =".format(left, right))

def random_plus_equation(max_num):
    max_num -= 1
    left = math.ceil(random.random() * max_num)
    right = math.ceil(random.random() * (max_num - left + 1))
    print("{:2d} + {:2d} =".format(left, right))

count = 24
max_num = 30
while count > 0:
    if random.random() > 0.5:
        random_minus_equation(max_num)
    else:
        random_plus_equation(max_num)
    count -= 1
    
