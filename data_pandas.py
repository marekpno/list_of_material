# create def
def data_pandas(df):
    # delated spaces
    df["Indeks Czesci"] = df["Indeks Czesci"].str.replace(' ', '')
    # convert value to str
    df.loc[:, ["Numer"]] = df.loc[:, ["Numer"]].applymap(str)
    # count how many dots is in cell and create new col(?)
    df_kropki = df.loc[:, ["Numer"]].applymap(lambda x: x.count('.'))
    # match dot to n
    n = int(df_kropki.max())
    # create col
    for kolumny in range(0, n+1):
        # name of col
        nazwa = n-kolumny
        # add new col
        df.insert(1, nazwa, "")
    # split col by .
    df.loc[:, "Numer":n] = df["Numer"].str.split(".", expand=True)
    # del col numer
    df = df.drop(columns=["Numer"])
    # drop nan
    df = df.dropna(how="all")
    # deleted cell with value like below
    df = df.replace({"Materiał": {"Materiał <nieokreślony>": "X"}})
    # deleted cell with value like below
    df = df.replace({"Materiał": {"Element handlowy": "handlowy"}})
    # deleted cell with value like below
    df = df.replace("Część wieloobiektowa", "")

    # need add 0 becouse can't change to float
    df.loc[:, ["Grubość ", "Dlugosc", "Szerokość"]] = \
        df.loc[:, ["Grubość ", "Dlugosc", "Szerokość"]].replace("", "0")
    # change col to float
    df.loc[:, ["Grubość ", "Dlugosc", "Szerokość"]] =\
        df.loc[:, ["Grubość ", "Dlugosc", "Szerokość"]].astype(float)
    # round value
    df = df.round({"Masa": 1, "Grubość ": 0, "Dlugosc": 0, "Szerokość": 0})
    # add 0 in col
    df.loc[:, "Gięcie":"Zakupowe"] = df.loc[:, "Gięcie":"Zakupowe"].fillna(0)
    # convert to int
    df.loc[:, "Gięcie":"Zakupowe"] =\
        df.loc[:, "Gięcie":"Zakupowe"].astype("int")
    # converto to str
    df.loc[:, "Gięcie": "Zakupowe"] =\
        df.loc[:, "Gięcie":"Zakupowe"].astype("str")
    # replace 0 to empty ("")
    df.loc[:, "Gięcie": "Zakupowe"] =\
        df.loc[:, "Gięcie": "Zakupowe"].replace("0", "")
    # replace 1 to X
    df.loc[:, "Gięcie": "Zakupowe"] =\
        df.loc[:, "Gięcie": "Zakupowe"].replace("1", "X")
    # change spaces to nothing ("")
    df["Indeks Czesci"] = df["Indeks Czesci"].str.replace(' ', '')
    # drop col becouse I don't need
    df = df.drop(columns=['Indeks materiałowy'])

    # save (optional)
    df.to_excel("test_excel.xlsx", index=False)

    print("pandas finished sucessfull optional--> check text_excel.xlsx")
    return df
