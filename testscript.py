import os
total = 0
data = tuple(os.walk('E:/ERP and SCM'))
for (dir, subdir, files) in data:
    total = total + len(files)
print(total)
print(len('E:\demo\There is a folder inside\There is another folder inside\There is another folder inside\There is another folder inside\There is another folder inside\There is another folder inside\There is another folder inside\There is another folder insi'))