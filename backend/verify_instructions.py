

def verify_instructions(instructions):
    for num in range(0,len(instructions)): #Iterates through all rows
        #confirms the difference between all widths is more than 6
        for i in range(len(instructions[num][1])):
            for j in range(i + 1, len(instructions[num][1])):
                if abs(instructions[num][1][i] - instructions[num][1][j]) < 6:
                    return False
    return True