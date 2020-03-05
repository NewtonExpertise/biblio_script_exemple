#coding:utf-8
import threading

def process_one():

    i=0
    while i<3:
        print('0000000')
        i+=1

def process_two():

    i=0
    while i<3:
        print('xxxxxxxx')
        i+=1

th1 = threading.Thread(target=process_one)
th2 = threading.Thread(target=process_two)

th1.start()
th2.start()

th1.join()
th2.join()