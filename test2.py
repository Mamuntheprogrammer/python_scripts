from itertools import combinations,islice
import itertools
import time

start = time.process_time()
comb = list(combinations([2, 1, 3,6,7,8,5,9,1, 3,6,7,8,5,9], 4) )
print(time.process_time() - start)
print(comb)




# print("the length of combination is " + str(len(comb)))

a,b=map(int,input("enter a= lower limit and b= upper limit ").split())


start = time.process_time()
comb1 = list(itertools.islice(combinations([2, 1, 3,6,7,8,5,9,1, 3,6,7,8,5,9], 4), a, b))
print(time.process_time() - start)
print(comb1)






