#Mult 3 fizz
#Mult 5 buzz
#Mult 3 & 5 fizz buzz

n = input()
n = int(n) + 1
t = []
f = []
tf = []
tn = 0
fn = 0
while tn < n:
	tn += 3
	t.append(tn)
while fn < n:
	fn += 5
	f.append(fn)
for x in t:
	if x not in tf:
		for y in f:
			if x == y and x not in tf:
				tf.append(x)
				break
for a in range(n):
	if a in tf:
		print("Fizz Buzz")
	elif a in t:
		print("Fizz")
	elif a in f:
		print("Buzz")
	else:
		print(a)
	


#____________

n = input()
n = list(range(1,int(n)+1))
for i in n:
	if i % 3 == 0 and i % 5 == 0:
		print("FizzBuzz")
	elif i % 3 == 0:
		print("Fizz")
	elif i % 5 == 0:
		print("Buzz")
	else:
		print(i)



#____________

n = input()
n = list(range(1,int(n)+1))
t = 3
f = 5
tw = "Fizz"
fw = "Buzz"
for i in n:
	op = ""
	if i % t == 0: op = op + tw
	if i % f == 0: op = op + fw
	if op == "": op = str(i)
	print(op)



#_________

n = input()
v = [3,5]
o = ["Fizz","Buzz"]
a = lambda x,y: y % x == 0
for i in list(range(1,int(n)+1)):
	p = ""
	c = 0
	for j in v:
		if a(j,i) == True: p = p + o[c]
		c += 1
	if p == "": p = str(i)
	print(p)
