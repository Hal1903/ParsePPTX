from tika import parser
from textblob import TextBlob

parsed = parser.from_file('order.pptx')

def create(file):#*args):
    filenum=0
    parsed = file#args[filenum]
    count = 0
    f = open("test.txt", "a", encoding='utf8')
    for x in range(len(parsed["content"])):
        if parsed["content"][x] == "　":
            pass #f.write('s')
        elif parsed["content"][x] =="\n":
            f.write("\n")
        elif parsed["content"][x]==" ":
            f.write(" ")
            pass #f.write("//")
        else:
            f.write(parsed["content"][x])
            #print(x, end="")
    filenum += 1
    f.close()

def modify(file):
    f = open(file, "r", encoding="utf8")
    f2 = open("test1.txt", "a", encoding="utf8")
    removed = open("removed.txt", "a", encoding="utf8")
    count = 0
    for line in f:
        if line.isspace():
            pass
        else:
            line = line.rstrip()
            if not (line[-1]=="." or line[-1]=="。" or line[-1]=="?" or line[-1]=="？" or line[-1]=="!" or line[-1]=="！" or line[-1]==")" or line[-1]=="）"):
                line = line + "."
                f2.write(line+"\n")
            else:
                f2.write(line+"\n")
        count += 1
    f.close()
    f2.close()
    removed.close()

s = "こんちゃっす"
print(TextBlob(s).detect_language())
