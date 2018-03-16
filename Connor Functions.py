def compare(dict1,dict2):
    # Initializes variables storing the shortest list in the array and its index
    shortest = 10000
    key = 0
    dLen = len(dict1)
    # The loop checks for the shortest list in the array and records its index
    for i in dict1.keys():
        dLen = len(dict1[i]) + len(dict2[i])
        if (dLen <= shortest):
            shortest = dLen
            key = i
    return key

def eliminate(volunteer,dict):
    # This function removes a volunteer from the dictionary
    for key in dict.keys():
        dict[key].remove(volunteer)

    return dict


# TEST CODE FOR compare()
#a={"apple":[1,1,1,1,1,1], "pie":[2,1,1,1]}
#b={"apple":[1,1], "pie":[1,1]}
#z = compare(a,b)
#z = eliminate(1,a)
#print(z)
