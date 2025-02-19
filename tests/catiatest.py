def test():
    att = "let lst(list) set lst=cm\\iBodys let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1} cm\\iMass=Vol*cm\\iDensity"
    print(att)


if __name__ == "__main__":
    test()
