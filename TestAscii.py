def toBase (num, base):
    new_base = []
    while num > 0:
        new_base.append(num%base)
        num = int(num/base)
    return new_base

def toDec(num, base):
    strNum = str(num)
    strNum = strNum[::-1]
    decimal = []
    final = 0
    
    for i in range(len(strNum)):
        decimal.append(int(strNum[i])*base**i)
        
    for j in decimal:
        final += j 
    return final
    
def numToAscii (num):
    strNum = str(num)
    letter = []
    for i in strNum:
        letter.append(chr(int(i)+64))
    return letter

def asciiToNum (string):
    number = []
    for i in string:
        number.append(ord(i)-64)
    return number
    
    
toDec(13, 26)
print(asciiToNum('AZ'))
print(toDec(126, 26))
print(toBase(734,26))









