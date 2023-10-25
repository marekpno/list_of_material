# check name of headers
def head(df):
    # empty list of head
    head_error = []
    # set of heads Grubosc has extra space !!!!!!!!!!!!!!!!!
    headers = ['NumerX', 'Indeks CzesciXX', 'indeks rysunku', 'Ilosc', 'Opis',
               'Masa', 'Materiał', 'Grubość ', 'Dlugosc', 'Szerokość',
               'Gięcie', 'Plazma', 'Laser', 'Standardowa wypalka', 'Tokarka',
               'Piła', 'Spawanie', 'Zakupowe', 'Pokrycie', 'dxf',
               'Indeks materiałowy', 'Revision', 'Typ czesci', 'Komentarz']

    # create name of col
    tf = df.columns
    # change tf to list
    df_list = tf.values.tolist()
    # set first step
    i = 0
    # loop for description in list
    for description in df_list:
        # check that headers are ok
        if df_list[i] == headers[i]:
            pass
        # dd bed headers to list
        else:
            head_error.append(df_list[i])
        # add iteration
        i = i + 1
    # retur head_error
    return head_error


# def check first row with tree of parts


def numer(df):
    # create empty list
    numer_error = []
    # check by cordiate
    nrows, ncol = df.shape
    # import cordinate
    from Cordinate import cordinate
    # col for check
    tf1 = cordinate(df, col_start="A", col_end="B", row_start=0, row_end=nrows)
    # def
    tf1 = tf1.Numer
    # becouse I cant import row below from cordinate
    numer_value = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.']
    # loop for cell in tf1
    for cell in tf1:
        # loop for char in cell
        for char in cell:
            # if ok pass
            if char in numer_value:
                pass
            # if bed add to list of error
            else:
                numer_error.append(cell)
    return numer_error
