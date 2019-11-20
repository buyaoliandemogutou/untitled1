
def write_file(L):
    try:
        f = open("d:/input.txt","w")
        for x in L:
            f.write(x)
            f.write('\n')
        f.close()
    except IOError:
        print("write error;")

if __name__ == "__main__":
    l={"qw","we"}
    write_file(l)
