import os

if __name__ == "__main__":
    result = next(os.walk('E:/',topdown=True))
    print(result[1])

    for i in result[1]:
        print(i)
        total = 0
        for (root, subdir, files) in os.walk('E:/'+i):
            total = total + len(files)
        print(total)



