from backend.add_one_by_one import add_one_by_one


def generate_starting_values(row_heights):
    int_vals = []
    cumsum = add_one_by_one(row_heights)
    for num in range(0,len(cumsum)):
        int_vaL = cumsum[num]-row_heights[num]+1
        int_vals.append([int_vaL,cumsum[num]])
    return int_vals 
