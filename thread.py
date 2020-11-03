# # -*- coding: utf-8 -*-
#
# """
# Three threads print A B C in order.
# """
#
#
# from threading import Thread, Condition
#
# condition = Condition()
# current = "A"
#
#
# class ThreadA(Thread):
#     def run(self):
#         global current
#         for _ in range(10):
#             with condition:
#                 while current != "A":
#                     condition.wait()
#                 print("A")
#                 current = "B"
#                 condition.notify_all()
#
#
# class ThreadB(Thread):
#     def run(self):
#         global current
#         for _ in range(10):
#             with condition:
#                 while current != "B":
#                     condition.wait()
#                 print("B")
#                 current = "C"
#                 condition.notify_all()
#
#
# class ThreadC(Thread):
#     def run(self):
#         global current
#         for _ in range(10):
#             with condition:
#                 while current != "C":
#                     condition.wait()
#                 print("C")
#                 current = "A"
#                 condition.notify_all()
#
#
# if __name__ == '__main__':
#     a = ThreadA()
#     b = ThreadB()
#     c = ThreadC()
#
#     a.start()
#     b.start()
#     c.start()
#
#     a.join()
#     b.join()
#     c.join()


import asyncio
import time
from random import randint


# @asyncio.coroutine
async def start_state():
    print("Start State called \n")
    input_value = randint(0, 1)
    time.sleep(1)
    if input_value == 0:
        result = await state2(input_value)
    else:
        result = await state1(input_value)
    print("Resume of the Transition : \nStart State calling " + result)


# @asyncio.coroutine
async def state1(transition_value):
    output_value = str("State 1 with transition value = %s \n" % transition_value)
    input_value = randint(0, 1)
    time.sleep(1)
    print("...Evaluating...")
    if input_value == 0:
        result = await state3(input_value)
    else:
        result = await state2(input_value)
    result = "State 1 calling " + result
    return output_value + str(result)


# @asyncio.coroutine
async def state2(transition_value):
    output_value = str("State 2 with transition value = %s \n" % transition_value)
    input_value = randint(0, 1)
    time.sleep(1)
    print("...Evaluating...")
    if input_value == 0:
        result = await state1(input_value)
    else:
        result = await state3(input_value)
    result = "State 2 calling " + result
    return output_value + str(result)


# @asyncio.coroutine
async def state3(transition_value):
    output_value = str("State 3 with transition value = %s \n" % transition_value)
    input_value = randint(0, 1)
    time.sleep(1)
    print("...Evaluating...")
    if input_value == 0:
        result = await state1(input_value)
    else:
        result = await EndState(input_value)
    result = "State 3 calling " + result
    return output_value + str(result)


# @asyncio.coroutine
async def EndState(transition_value):
    output_value = str("End State with transition value = %s \n" % transition_value)
    print("...Stop Computation...")
    return output_value


if __name__ == "__main__":
    print("Finite State Machine simulation with Asyncio Coroutine")
    loop = asyncio.get_event_loop()
    loop.run_until_complete(start_state())
