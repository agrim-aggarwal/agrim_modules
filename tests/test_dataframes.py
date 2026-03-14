import pandas as pd
import numpy as np
import itertools
from agrim_modules import create_sheet

outfile = 'test.xlsx'
sheet_name = 'Sheet1'

ser_a = [f"a{i}" for i in range(3)]
ser_b = [f"b{i}" for i in range(2)]
ser_c = [f"k{i}" for i in range(2)]

a = pd.DataFrame(itertools.product(ser_a, ser_b, ser_c), columns=['c1', 'c2', 'c3'])
a['c4'] = a.apply(lambda x: np.random.rand(), axis=1)
a['c5'] = a.apply(lambda x: np.random.rand(), axis=1)

b = a.pivot_table(index=['c1'], columns = ['c3'], values='c4', aggfunc=['min', 'max'])
b2 = a.pivot_table(index=['c1', 'c2'], columns = ['c3'], values='c4', aggfunc=['min', 'max'])
b3 = a.pivot_table(index=['c1'], columns = ['c3','c2'], values='c4', aggfunc=['min', 'max'])
b4 = a.pivot_table(index=['c1'], values='c4', aggfunc=['min', 'max'])
b4.columns = b4.columns.swaplevel(0,1)
c = a.groupby(['c1', 'c2']).agg(max =('c4', 'max'), min=('c4', 'min'))
d = a.groupby(['c1'])[['c4']].sum()

# df = a.copy()
# df = b.copy()
# df = c.copy()
df = d.copy()

with pd.ExcelWriter(outfile, engine='xlsxwriter') as writer:
    create_sheet(a, writer, 'Sheet1')
    create_sheet(b, writer, 'Sheet2')
    create_sheet(b2, writer, 'Sheet2_2')
    create_sheet(b3, writer, 'Sheet2_3')
    create_sheet(b4, writer, 'Sheet2_4', duplicate_header = True)
    create_sheet(c, writer, 'Sheet3')
    create_sheet(d, writer, 'Sheet4')            