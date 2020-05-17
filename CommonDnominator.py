n = 9*9*9*9*9*9*9
v = [7,11,13,17,23]
a = lambda x,y: y % x == 0
for i in list(range(1,int(n)+1)):
	p = 0
	c = 0
	for j in v:
		if a(j,i) == True: p += 1
		c += 1
	if p == c: 
		print(i)
		break
