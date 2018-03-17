def surplus(dict1,dict2,client):
    # Initializes variables storing the shortest list in the array and its index
    # All three dictionaries must have the exact same keys
    shortest = 10000
    # The loop checks for the shortest list in the array and records its index
    for i in dict1.keys():
        tLen = len(dict1[i]) + len(dict2[i]) - len(client[i])
        if (tLen <= shortest):
            shortest = tLen
            key = i
    return key


def eliminate(volunteer, vDict):
    # This function removes a volunteer from the dictionary
    for key in vDict.keys():
        vDict[key].remove(volunteer)

    return vDict

# TEST CODE FOR surplus()
# a={"apple":[1,1,1,1,1,1], "pie":[2,1,1,1]}
# b={"apple":[1,1], "pie":[1,1]}
# z = compare(a,b)
# z = eliminate(1,a)
# print(z)
