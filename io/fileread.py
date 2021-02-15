f = open("./tmp-file/python-code.txt", "r")

print(f.tell())

for line in f:
    print(line, end='$ ')

print(f.tell())
f.close()
