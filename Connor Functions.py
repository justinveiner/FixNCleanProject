def compare(dict1,dict2):
    # Initializes variables storing the shortest list in the array and its index
    shortest = 10000
    key=0
    dLen = len(dict1)
    # The loop checks for the shortest list in the array and records its index
    for i in range(dLen):
        dLen = len(dict1[i]) + len(dict2[i])
        if (dLen <= shortest):
            shortest = dLen
            key = dict1[i]

    return key



# TEST CODE FOR compare()
a={0:[1,1,1,1,1,1], 1:[2,1,1,1]}
b={0:[1,1], 1:[1,1]}
z=compare(a,b)
print(z)