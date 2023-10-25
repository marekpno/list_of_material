# input adres for file
def link():
    # input adres
    adres = input("input adres csv file>")
    # return adres
    return adres

# load data to pandas


def adres(adres):
    # import pd
    import pandas as pd
    # loop while
    while True:
        # try to read
        try:
            # condition for cvs (link)
            if (adres[0:5]) == "https":
                # url
                url = adres
                # add begin
                url = 'https://drive.google.com/uc?id=' + url.split('/')[-2]
                # match dataframe and encoding  win 1250 -  polish signs
                df = pd.read_csv(url, encoding='windows-1250')
                # return df
                return df
            # condition for csv from local disc
            else:
                # url
                url = adres
                # match dataframe and encoding  win 1250 -  polish signs
                df = pd.read_csv(adres, encoding='windows-1250')
                # return df
                return df
        # 1 first type of loaded error
        except IOError:
            # info
            print("first type of loaded error")
            # input adres one more time
            adres = input("input adres csv file>")
        # 1 second type of loaded error
        except IndexError:
            # info
            print("second type of loaded error")
            # input adres one more time
            adres = input("input adres csv file>")
        # if cvs was lodaed without error --> break
        else:
            # return df
            return df
            # break loop
            break