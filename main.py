import pandas as pd
import numpy as np
from matplotlib import pyplot as plt

plt.rcParams["figure.figsize"] = (10, 5)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
plt.style.use("fivethirtyeight")

c_df = pd.read_excel(open('customers.xls', 'rb'), sheet_name='customers')
r_df = pd.read_excel(open('residence.xls', 'rb'), sheet_name='locations')


def data_prep():
    # converting to datetime
    c_df["Տրվել է(օր)"] = pd.to_datetime(c_df["Տրվել է(օր)"])

    # empty "Շրջանառություն" --> 0
    c_df["Շրջանառություն"].fillna(0, inplace=True)

    # empty "Տարիքային.խումբ" --> "Not mentioned"
    c_df["Տարիքային.խումբ"].fillna("Չնշված", inplace=True)

    # empty "Համաձայն.է.ստանալ.SMS" --> "n"
    c_df["Համաձայն.է.ստանալ.SMS"].fillna("n", inplace=True)

    # empty "Կուտակում" --> 0
    c_df["Կուտակում"].fillna(0, inplace=True)

    # empty "Վճարում" --> 0
    c_df["Վճարում"].fillna(0, inplace=True)


def gender_distrb():
    gender_gr = c_df["Սեռ"].value_counts()
    labels = ['Կին', 'Տղամարդ']
    explode = [0, 0.1]

    plt.pie(gender_gr, labels=labels, wedgeprops={'edgecolor': 'black'}, shadow=True, autopct='%.2f%%', explode=explode)
    plt.title("Քարտապանները ըստ սեռի")
    plt.tight_layout()
    plt.show()


def given_by_week_days():
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    label = ["Երկ", "Երք", "Չոր", "Հին", "Ուրբ", "Շաբ", "Կիր"]

    gender_group = c_df.groupby(["Սեռ"])
    gender_group_m = gender_group.get_group("Мужской")
    day_names_m = gender_group_m["Տրվել է(օր)"].dt.day_name().value_counts()
    day_names_m = day_names_m.reindex(index=day_order)
    y_m = day_names_m.values

    gender_group_f = gender_group.get_group("Женский")
    day_names_f = gender_group_f["Տրվել է(օր)"].dt.day_name().value_counts()
    day_names_f = day_names_f.reindex(index=day_order)
    y_f = day_names_f.values

    day_names = c_df["Տրվել է(օր)"].dt.day_name().value_counts()
    day_names = day_names.reindex(index=day_order)
    y = day_names.values

    x_indexes = np.arange(7)  # week days == 7

    width = 0.25
    plt.bar(x_indexes + width, y_m, width=width, label="Տղամարդիք")
    plt.bar(x_indexes, y, width=width, color="green", label="Ընդհանուր")
    plt.bar(x_indexes - width, y_f, width=width, color="red", label="Կանայք")

    plt.legend(loc="upper left")
    plt.xticks(x_indexes, label, rotation='vertical')
    plt.title("Տրամադրված բոնուս քարտերը ըստ շաբաթվա օրերի")
    plt.xlabel("Բոնուս քարտերի քանակ")
    plt.ylabel("Շաբաթվա օրեր")
    plt.tight_layout()
    plt.show()


def branch_turnover():
    branch = c_df.groupby(["Տրվել.է.խանութ."])

    turnover_br = branch["Շրջանառություն"].sum()
    turnover_br.sort_values(ascending=True, inplace=True)
    x = turnover_br.index
    y = turnover_br.values / 1000
    plt.ticklabel_format(useOffset=False, style='plain')
    plt.barh(x, y)
    plt.title("Շրջանառությունը ըստ մասնաճյուղերի (մլն․/դրամ)")
    plt.xlabel("Շրջանառություն (մլն․/դրամ)")
    plt.ylabel("Մասնաճյուղերի անվանում")
    plt.tight_layout()
    plt.show()


def age_turnover():
    age = c_df.groupby(["Տարիքային.խումբ"])
    turnover_age = age["Շրջանառություն"].sum()

    age_order = ["Չնշված", "Մինչև 18 տարեկան", "25-ից 45 տարեկան", "18-ից 25 տարեկան", "45-ից բարձր"]
    turnover_age = turnover_age.reindex(index=age_order)

    new_x = ["Չնշված", "Մինչև 18 տ.", "25-ից 45 տ.", "18-ից 25 տ.", "45-ից բարձր"]
    y = turnover_age.values / 1000000

    plt.bar(new_x, y)
    plt.title("Շրջանառությունը ըստ տարիքային խմբերի (մլն/դրամ)")
    plt.xlabel("Տարիքային խմբեր")
    plt.ylabel("Շրջանառություն(մլն/դրամ)")
    plt.tight_layout()
    plt.show()


def age_card_holder():
    age_count = c_df["Տարիքային.խումբ"].value_counts()

    age_order = ["Չնշված", "Մինչև 18 տարեկան", "18-ից 25 տարեկան", "25-ից 45 տարեկան", "45-ից բարձր"]
    age_count = age_count.reindex(index=age_order)

    new_x = ["Չնշված", "Մինչև 18 տ.", "25-ից 45 տ.", "18-ից 25 տ.", "45-ից բարձր"]
    explode = [0.1, 0.15, 0.1, 0.1, 0.1]
    colors = ["blue", "orange", "green", "yellow", "red"]

    y = age_count.values

    plt.pie(y, labels=new_x, explode=explode, wedgeprops={'edgecolor': 'black'}, autopct='%.2f%%', colors=colors)
    plt.title("Քարտապանները ըստ տարիքային խմբերի")
    plt.tight_layout()
    plt.show()


def sms_agree_rcv():
    sms_agr = c_df["Համաձայն.է.ստանալ.SMS"].value_counts()
    labels = ['Համաձայն.է.ստանալ.SMS', 'Համաձայան չէ']
    explode = [0, 0.1]

    plt.pie(sms_agr, labels=labels, wedgeprops={'edgecolor': 'black'}, shadow=True, autopct='%.2f%%', explode=explode)
    plt.title("Քարտապանները ըստ SMS ստանալու համաձայնության")
    plt.tight_layout()
    plt.show()


def sms_agree_turnover():
    sms_agr = c_df.groupby(["Համաձայն.է.ստանալ.SMS"])
    sms_agr_turnover = sms_agr["Շրջանառություն"].sum()

    new_x = ["Համաձայն է", "Համաձայն չէ"]
    y = sms_agr_turnover.values

    explode = [0.1, 0.1]
    colors = ["blue", "red"]

    plt.pie(y, labels=new_x, explode=explode, wedgeprops={'edgecolor': 'black'}, autopct='%.2f%%', colors=colors,
            shadow=True)

    plt.title("Շրջանառությունը ըստ SMS ստանալու համաձայնության (մլն/դրամ)")
    plt.tight_layout()
    plt.show()


def city_village_distrb():
    n = 10
    city_village = r_df['Քաղաք/Գյուղ'].value_counts()
    city_village = city_village.nlargest(n).sort_values(ascending=True)
    x = city_village.index
    y = city_village.values
    plt.barh(x, y)

    plt.title(f"{n} ամենաշատ հաճախորդ ունեցող քաղաք/գյուղեր")
    plt.xlabel("Հաճախորդների քանակը")
    plt.ylabel("Քաղաք/Գյուղ")
    plt.tight_layout()
    plt.show()


def c_v_turnover():
    n = 10
    result = pd.merge(r_df, c_df[["id", "Շրջանառություն"]], on="id")
    c_v = result.groupby(["Քաղաք/Գյուղ"])
    c_v_turn = c_v["Շրջանառություն"].sum()
    c_v_turn = c_v_turn.nlargest(n).sort_values(ascending=True)

    x = c_v_turn.index
    y = c_v_turn.values / 1000
    plt.barh(x, y)

    plt.title(f"{n} ամենամեծ շրջանառություն ունեցող քաղաք/գյուղեր")
    plt.xlabel("Շրջանառություն(մլն/դրամ)")
    plt.ylabel("Քաղաք/Գյուղ")
    plt.tight_layout()
    plt.show()


def bonus_p_np():
    bonus = c_df["Կուտակում"].sum()
    payed = c_df["Վճարում"].sum()

    unused = bonus - payed
    y = [unused, payed]
    new_x = ["Չվճարված", "Վճարված"]
    explode = [0.1, 0.1]
    colors = ["blue", "red"]
    plt.pie(y, labels=new_x, explode=explode, wedgeprops={'edgecolor': 'black'}, autopct='%.2f%%', colors=colors,
            shadow=True)

    plt.title("Կուտակված բոնուսների վճարված և չվճարված մասերը")
    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    data_prep()
    gender_distrb()
    given_by_week_days()
    branch_turnover()
    age_turnover()
    age_card_holder()
    sms_agree_rcv()
    sms_agree_turnover()
    city_village_distrb()
    c_v_turnover()
    bonus_p_np()
