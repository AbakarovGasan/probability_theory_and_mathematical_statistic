import olefile 
import pandas as pd
import matplotlib.pyplot as plt
import random
from math import *
import os
import tkinter as tk
from statistics import *
from scipy.stats import kurtosis#эксцесс
from scipy.stats import skew    #ассиметрия (откос)
from scipy.stats import sem
from scipy.stats import t
from scipy.stats import linregress
import scipy
import numpy as np
import statsmodels.api as sm

def norm_dist(a, o=1):
    y=lambda x: exp(-(((x-3)**2)/o))
#g(x)=a * exp( - ( (x - b)^2 ) / (2*(c^2))  )
#https://en.wikipedia.org/wiki/Gaussian_function
#a = 1
#b = 3
#o = 2 * c^2 = random
    q1=y(1)
    q2=y(2)
    q3=y(3)
    q4=(q2+q1)*2+1
    
    i=a/q4
    i1=ceil(i)
    if ((a%2==1) & (i1%2==0))|((i1%2==1) & (a%2==0)):
       i1+=1
    a-=i1
    a/=2
    i2=ceil((a/(q2+q1))*q2)
    i3=a-i2
    if i3==0:
        i3=1
        i2-=1
    if i2==0:
        i2=1
        i1-=2
    return int(i1), int(i2), int(i3)

r=tk.Tk()
filename = tk.filedialog.askopenfilename(
                initialdir= os.getcwd(),
                title= "Выберите файл задания",
                filetypes=(("xls files","*.xls"),))
r.destroy()
print(filename)

ole = olefile.OleFileIO(open(filename, 'rb'))
#https://olefile.readthedocs.io/en/latest/OLE_Overview.html
#if ole.exists("Workbook"):
x = pd.read_excel(ole.openstream("Workbook")) 

_print = print

def print(a='', *b, **c):
    a = '   '+str(a)
    return _print(a, *b, **c)

print()
print()
    
def parse(j):
    j = str(j)
    if j=='nan':
        raise Exception("")
    o=float(j.replace(',','.'))
    return o

a = []    
k = x.keys()
print(k.values)

book={}
book[1]=[]
book[2]=[]
book[3]=[]

b_in = 0 

Y=1
X1=2
X2=3

d = {'1': 'X1', '2':'X2', '3':'Y'}
for u in k[1:]:
    i = ''
    b_in+=1
    while not i in d.keys():
        print(u, d, end=': ')
        i = input()
    globals()[d.pop(i)] = b_in

d = {X1: 'X1', X2:'X2', Y:'Y'}
print()
print()

for i in x.values:
    try:
        c = [i[0], parse(i[1]), parse(i[2]), parse(i[3])]
        if c[0].lower()!='страна': a.append(c)
    except Exception as e:
        pass

def out1(i):
    t=str(i[0])
    print(t, ' '*(30-len(t)), end='')
    t=str(i[1])
    print(t, ' '*(10-len(t)), end='')
    t=str(i[2])
    print(t, ' '*(10-len(t)), end='')
    t=str(i[3])
    print(t, ' '*(10-len(t)), end='')
    print()
    
def out2(c, i):
    t=str(c)+':'
    print(t, ' '*(5-len(t)), end='')
    for t in i:
        t=str(t)
        print(t, ' '*(8-len(t)), end='')
    print()

for i in a:
   # print(i)
    book[3].append(i[3])
    book[2].append(i[2])
    book[1].append(i[1])
    
x1=book[X1]
x2=book[X2]
y=book[Y]

x1vr=x1.copy()
x1vr.sort()
print()
print('X1вр =' , x1vr)
print('размах: ', x1vr[-1]-x1vr[0])
print()
x2vr=x2.copy()
x2vr.sort()
print('X2вр =' , x2vr)
print('размах: ', x2vr[-1]-x2vr[0])
print()
yvr=y.copy()
yvr.sort()
print('Yвр =' , yvr)
print('размах: ', yvr[-1]-yvr[0])
print()
print()


print('Y:', k[Y])
print('X1:', k[X1])
print('X2:', k[X2])
out1(['Cтрана', d[1], d[2], d[3]]) 
for i in a:
    out1(i)
    
class glob:
    def refresh(s):
        a1, a2, a3 = norm_dist(len(y), 5*random.random() )
        k1=a3-1
        k2=k1+a2
        k3=k2+a1
        k4=k3+a2
        k5=k4+a3
        s.k1=k1
        s.k2=k2
        s.k3=k3
        s.k4=k4
        s.k5=k5
glob=glob()
glob.refresh()

#print(glob.__dict__)

def rand(y, x):
    g=floor(y-x)
 #   print('______________________')
 #   print(y, x)
    if g==0:
 #       print(x)
        return x
    x+=random.randint(0, g-1)
    l=0
    if g>100000:
        l=100000
    if g>10000:
        l=10000
    if g>1000:
        l=1000
    elif g>100:
        l=100
    elif g>10:
        l=10
    else:
 #       print(x)
        return x
    x1=x+(l-(x%l))
    if x1>y:
        x1-=l
 #   print(x1)
    return x1



def generate_interval(x):
    k = [ [ rand(x[glob.k1+1],x[glob.k1]),
    rand(x[glob.k2+1],x[glob.k2]),
    rand(x[glob.k3+1],x[glob.k3]),
    rand(x[glob.k4+1],x[glob.k4]),
    x[glob.k5] ] ]
    for i in range(3):
        j=[]
        z=glob.k5+1
        o=-1
        k.append(j)
  #      print('----------------new ---------------')
        for i in range(6):
   #         print(z)
            l=int(o+(z/2))
            if (l)>(o+1):
                o=random.randint(o+1, l)
            else:
                o=o+1
   #         print('o', o)
            z-=o
            if z <=0:
                break
            if o==glob.k5:
                b=x[o]
                if not b in j:
                    j.append(b)
            else:
    #            print('xo', x[o])
     #           print('xo1', x[o+1])
                b=(rand(x[o+1], x[o]))
                if not b in j:
                    j.append(b)
        j.append(x[-1])
    return k

#print('lockup')
#print(generate_interval(x2vr))


print()
print("""Разбиения интервала""")
print("Для параметра Y")
gY=generate_interval(yvr)#[ [1000, 2500, 150000, 450000, 500000],
    # [1500, 15000, 300000, 500000], 
    # [5000, 10000, 30000, 300000, 500000], 
    # [2500, 20000, 100000, 300000, 500000] ]
for i in range(len(gY)): out2(str(i+1), gY[i])
print("""наилучший интервал - 1,
   получается гистограмма с наименьшим разбросом частот, график 
   распределения наиболее соответствует закону распределения, 
   заданной функцией Гаусса (нормальному распределению)""")
print()

print("Для параметра X1")
gX1=generate_interval(x1vr)#[ [74, 90, 100],
      #[62, 75, 80, 100], 
      #[65, 72, 85, 100], 
      #[62, 75, 82, 92, 100] ]
for i in range(len(gX1)): out2(str(i+1), gX1[i])
print("""наилучший интервал - 1,
   получается гистограмма с наименьшим разбросом частот, график 
   распределения наиболее соответствует закону распределения, 
   заданной функцией Гаусса (нормальному распределению)""")
print()

print("Для параметра Х2")
gX2=generate_interval(x2vr)#[ [25, 30, 35, 40, 52],
     # [15, 20, 35, 40, 52], 
      #[23, 32, 41, 49, 52], 
      #[25, 35, 45, 49, 52] ]
      
for i in range(len(gX2)): out2(str(i+1), gX2[i])
print("""наилучший интервал - 1,
   получается гистограмма с наименьшим разбросом частот, график 
   распределения наиболее соответствует закону распределения, 
   заданной функцией Гаусса (нормальному распределению)""")
print()
#input()

def show_graph(array, labels, name=''):
    ax = plt.gca()
    x=range(len(array))
    ax.bar(x, array, align='edge') 
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    plt.ylabel('частота')
    plt.xlabel('карман')
    plt.savefig(name+'.png')
    plt.clf()

def format(a, b):
    a = str(a)
    b -= len(a)
    return a + ' ' * b

def allocation(a1, a2):
    out=[]
    l=0
    for i in a1:
       z=0
       for u in a2:
           if u<=i and u>l:
               z+=1
       out.append(z)
       l=i
    return out

def show_alloc(x, y, c=''):
  #  print(*allocation(x, y))
  #  print(*x)
  #  print(end='\n\n')
    show_graph(allocation(x, y), x, c)
 
if not False:
 show_alloc(gY[0], y, 'для Y(1)')
 show_alloc(gY[1], y, 'для Y(2)')
 show_alloc(gY[2], y, 'для Y(3)')
 show_alloc(gY[3], y, 'для Y(4)')

 show_alloc(gX1[0], x1, 'для X1(1)')
 show_alloc(gX1[1], x1, 'для X1(2)')
 show_alloc(gX1[2], x1, 'для X1(3)')
 show_alloc(gX1[3], x1, 'для X1(4)')

 show_alloc(gX2[0], x2, 'для X2(1)')
 show_alloc(gX2[1], x2, 'для X2(2)')
 show_alloc(gX2[2], x2, 'для X2(3)')
 show_alloc(gX2[3], x2, 'для X2(4)')

def out3(c, a):
    print(c)
    print('выборочное среднее:                     ', mean(a)) 
    print('выборочная медианa:                     ', median(a)) 
    print('выборочная модa:                        ', mode(a)) 
    print('стандартное отклонение:                 ', stdev(a)) 
    print('дисперсия выборки:                      ', variance(a)) 
    print('выборочный коэффициент эксцесa:         ', kurtosis(a)) 
    print('выборочный коэффициент асимметричности: ', skew(a), end='\n\n') 
    
print('Числовые характеристики для Y, X1, X2')
out3('Для параметра Y', y)
out3('Для параметра Х1', x1)
out3('Для параметра Х2', x2)

def rel_l(data, confidence=0.95):
    s = scipy.stats.sem(data)
    s = s * scipy.stats.t.ppf((1 + confidence) / 2., len(data)-1)
    return s

ix1=t.interval(0.95, len(x1)-1, loc=mean(x1), scale=sem(x1))
ix2=t.interval(0.95, len(x2)-1, loc=mean(x2), scale=sem(x2))

print('Точечные и интервальные оценки неизвестных математических ожиданий ')
print('                             x1                   x2')
print('cреднее                     ', format(mean(x1), 20), mean(x2))
print('уровень надежности (95%)    ', format(rel_l(x1), 20), rel_l(x2))

print('x1 = ', mean(x1), '+-', rel_l(x1), sep='')
print('x2 = ', mean(x2), '+-', rel_l(x2), sep='')
print()

print('Матрица корреляции ')

f = np.corrcoef([y, x1, x2])

def link(a, b, c):
    o=abs(c)
    return 'Между '+a+' и '+b+' '+( 'cлабая' if (o<0.30)else ('умеренная' if (o<0.70)else 
    'сильная')  ) + ' ' + ( 'положительная' if (c>0) else 'отрицательная' )+' cвязь'

print('   |        Y        |       X1        |        X2       |')
print('-----------------------------------------------------------')
print('Y  |', format('%.0f' % f[0][0], 15), '|', format('%.10f' % f[0][1], 15), '|', format('%.10f' % f[0][2], 15), '|')
print('-----------------------------------------------------------')
print('X1 |', format('%.10f' % f[1][0], 15),'|', format('%.0f' % f[1][1], 15),'|', format('%.10f' % f[1][2], 15), '|')
print('-----------------------------------------------------------')
print('X2 |', format('%.10f' % f[2][0], 15),'|', format('%.10f' % f[2][1], 15),'|', format('%.0f' % f[2][2], 15), '|')
print('-----------------------------------------------------------')
print('Анализ полученных данных')
print(link('Y ('+k[Y]+')', 'X1 ('+k[X1]+')', f[1][0]))
print(link('Y ('+k[Y]+')', 'X2 ('+k[X2]+')', f[2][0]))
print(link('X1 ('+k[X1]+')', 'X2 ('+k[X2]+')', f[2][1]))



x = [x1]

def reg_m(y, x):
    ones = np.ones(len(x[0]))
    X = sm.add_constant(np.column_stack((x[0], ones)))
    for ele in x[1:]:
        X = sm.add_constant(np.column_stack((ele, X)))
    results = sm.OLS(y, X).fit()
    return results


def out4(y, x, name):
    print()
    print('Pегрессия ', name)
    f = reg_m(y, [x])
    print('R-квадрат:', f.rsquared)
    print('Значимость F:', f.f_pvalue)
    Y=f.params[1]
    X=f.params[0]
    print('Коэффицент '+name+': ', X)
    print('Коэффицент Y-пересечения:', Y)
    print('p-значение '+name+':', f.pvalues[0])
    print('p-значение Y-пересечения:', f.pvalues[1])
    r_sq=int(round(f.rsquared*100))
    print()
    print('Примерно '+str(r_sq)+'% разброса описываются линейной зависимостью')
     #y='+('%0.5f'%Y)+(('+') if (X>=0) else (''))+('%0.5f'%X)+'*x')
    print('Остальные '+str(100-r_sq)+'% приходятся на случайные факторы')
    if (f.f_pvalue>0.05):
        print('Значимость F больше 0.05, следовательно, регрессия статически НЕ значима')
    else:
        print('Значимость F НЕ больше 0.05, следовательно, регрессия статически значима')
    y2 = Y+X*x
    print('уравнение линейной зависимости: y=' + ('%.4f' % Y) + ('+' if (X>=0) else'') + ('%.4f' % X)+'*x')
    out5(x, y, y2)
    plt.savefig('график подбора ' + name)
    plt.clf()
    plt.plot(x, (y)-(Y+X*x), 's')
    plt.grid()
    plt.savefig('график остатков ' + name)
    plt.clf()

def out5(x, y, y2):
    plt.plot(x,y,'o')
    plt.plot(x,y2,'-')
    plt.grid()

def out6(y, x1, x2, name, name2):
    print()
    print('Общая регрессия ')
    f = reg_m(y, [x1, x2])
    print('R-квадрат:', f.rsquared)
    print('Значимость F:', f.f_pvalue)
    Y=f.params[2]
    X1=f.params[1]
    X2=f.params[0]
    print('Коэффицент '+name+': ', X1)
    print('Коэффицент '+name2+': ', X2)
    print('Коэффицент Y-пересечения:', Y)
    print('p-значение '+name+':', f.pvalues[1])
    print('p-значение '+name2+':', f.pvalues[0])
    print('p-значение Y-пересечения:', f.pvalues[2])
    r_sq=int(round(f.rsquared*100))
    print()
    print('Примерно '+str(r_sq)+'% разброса описываются линейной зависимостью')
     #y='+('%0.5f'%Y)+(('+') if (X>=0) else (''))+('%0.5f'%X)+'*x')
    print('Остальные '+str(100-r_sq)+'% приходятся на случайные факторы')
    if (f.f_pvalue>0.05):
        print('Значимость F больше 0.05, следовательно, регрессия статически НЕ значима')
    else:
        print('Значимость F НЕ больше 0.05, следовательно, регрессия статически значима')
    print('уравнение линейной зависимости: y=' + ('%.4f' % Y) + ('+' if (X1>=0) else'') + ('%.4f' % X1)+'*x1' , ('+' if (X2>=0) else'') + ('%.4f' % X2)+'*x2', sep='')
#    y2 = Y+X*x
#    out5(x, y, y2)
#    plt.savefig('график подбора ' + name)
#    plt.clf()
#    plt.plot(x, (y)-(Y+X*x), 's')
#    plt.grid()
#    plt.savefig('график остатков ' + name)
#    plt.clf()

x1=np.array(x1)
x2=np.array(x2)
#f = reg_m(y, x)
out4(y, x1, 'X1')
out4(y, x2, 'X2')
out6(y, x1, x2, 'X1', 'X2')


#x1 = np.array([1, 2, 3, 5])
#x2 = np.array([2, 1, 4, 7])
#y = [1, 2, 3, 4]





