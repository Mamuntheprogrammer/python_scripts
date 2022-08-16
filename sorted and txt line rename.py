import os,sys,glob
import time
import natsort

os.chdir(input("Enter Location address : "))
lst = natsort.natsorted(sorted(glob.glob("*.mp4")))

for x in lst:
	print(x)


with open('index.txt') as f:
    ll = [line for line in f.read().splitlines()]
print(ll)

for x in range(len(lst)):
	name,e = lst[x].split('.')
	os.rename(lst[x],name+'--'+ll[x]+'.'+e)
print("*** RENAME COMPLETED ***")