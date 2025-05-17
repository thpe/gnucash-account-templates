import pandas as pd
from pprint import pprint
import sys
import numpy as np
from piecash import create_book, Account, Commodity
from piecash.core.factories import create_currency_from_ISO


grupp = "Unnamed: 2"
konto = "Unnamed: 6"
gdesc = "Unnamed: 3"
kdesc = "Unnamed: 7"
f = sys.argv[1]
df = pd.read_excel(f)
df = df.applymap(lambda x: x.replace('\n', ' ').replace('\r', ' ') if isinstance(x, str) else x)
print(df)

acc_hira = {}

lastgroup = 0

for index, row in df.iterrows():
    g = row[grupp]
    k = row[konto]
    gd = row[gdesc]
    kd = row[kdesc]
    if isinstance(g, float) and g.is_integer():
        g = int(g)
        if len(str(int(abs(g)))) == 1:
            acc_hira[g] = {'desc' : gd, 'subacc' : {}}
        if len(str(int(abs(g)))) == 2:
            sup = int(str(int(abs(g)))[0])
            acc_hira[sup]['subacc'][g] = {'desc' : gd, 'subacc' : {}}
        if len(str(int(abs(g)))) == 3:
            sup1 = int(str(int(abs(g)))[0])
            sup2 = int(str(int(abs(g)))[0:2])
            print(f'account 3 {sup1} {sup2} {g}')
            acc_hira[sup1]['subacc'][sup2]['subacc'][g] = {'desc' : gd, 'subacc' : {}}
            lastgroup = g

    if isinstance(k, float) and k.is_integer():
        k = int(k)
        kd = ''.join(kd)
        sup1 = int(str(int(abs(k)))[0])
        sup2 = int(str(int(abs(k)))[0:2])
        sup3 = int(str(int(abs(k)))[0:3])
        acc_hira[sup1]['subacc'][sup2]['subacc'][sup3]['subacc'][k] = {'desc' : kd}
    print(f"Index: {index}, grupp: {row[grupp]}, konto: {row[konto]} {kd}")


pprint(acc_hira)
print(acc_hira)

sek = create_currency_from_ISO("SEK")

with create_book("my_books.gnucash", overwrite=True) as book:
    root = book.root_account

    t = 'ASSET'
    for k, v in acc_hira.items():
        if k == 2:
            t = 'LIABILITY'
        if k == 3:
            t = 'INCOME'
        if k >= 4:
            t = 'EXPENSE'
        a = Account(name=str(k) + ' ' + v['desc'], code = k, parent=root, commodity=sek, type=t)
        acc_hira[k]['gcacc'] = a
        for k1, v1 in v['subacc'].items():
            a1 = Account(name=str(k1) + ' ' + v1['desc'], code = k1, parent=a, commodity=sek, type=t)
            for k2, v2 in v1['subacc'].items():
                a2 = Account(name=str(k2) + ' ' + v2['desc'], code = k2, parent=a1, commodity=sek, type=t)
                for k3, v3 in v2['subacc'].items():
                    a3 = Account(name=str(k3) + ' ' + v3['desc'], code = k3, parent=a2, commodity=sek, type=t)

    book.save()
