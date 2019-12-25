def increment(num):
    if num[-1] == 'Z':
        temp = increment(num[:-1])+'A' if len(num) != 1 else 'AA'
        return temp
    else:
        ascii_val = ord(num[-1])
        return num[:-1]+chr(ascii_val+1)

test= 'ABCD'
temp = 'A'
for i in range(1000):
    print(temp)
    temp = increment(temp)
    
