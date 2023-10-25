# show cordinate for df pandas
def cordinate(df, col_start="A", col_end="F", row_start=0, row_end=100):
    # create empy list of car
    col_char = []
    # 'A' is represent by value 65
    # py convert A to 65 and Z to 90
    for char_code in range(ord('A'), ord('Z') + 1):
        # chr change value on character for example 65 to A
        character = chr(char_code)
        # add character to list
        col_char.append(character)
    # start col
    conventer_start = col_char.index(col_start)
    # end col
    conventer_end = col_char.index(col_end)
    # this will be input, what we want to find
    df_check = (df.iloc[row_start: row_end, conventer_start: conventer_end])
    # only one return can be return
    return df_check
