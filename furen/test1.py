t=[' A ','a']
import re
for i in t:
    print(re.findall(r'\w',i)[0].upper())

#第一次修改
#第二次修改