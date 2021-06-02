import pandas as pd
import xlsxwriter
#RENAME SESUAI PATH restoran.xlsx
df = pd.read_excel (r'D:\Tel U\Semester 4\Pengantar AI\Tupro 2\restoran.xlsx')

def fuzzification(p, m):
    if (m>=0 and m<=5):
        te =1
    elif (m>5 and m<=6):
        te = (-(m-6))/(6-5)
    else:
        te = 0

    if (m>5 and m<=6):
        biasa =(m-5)/(6-5)
    elif (m>6 and m<=8):
        biasa = (-(m-8))/(8-6)
    else:
        biasa = 0

    if (m>7 and m<9):
        enak=(m-7)/(9-7)
    elif (m>=9 and m<=10):
        enak=1
    else:
        enak = 0

    if (p>=0 and p<=60):
        tp=1
    elif (p>60 and p<=70):
        tp = (-(p-70))/(70-60)
    else:
        tp = 0
    
    if (p>50 and p<=70):
        puas = (p-50)/(70-50)
    elif (p>70 and p<=90):
        puas = (-(p-90))/(90-70)
    else :
        puas = 0
    
    if (p>85 and p<95):
        sangat=(p-85)/(95-85)
    elif (p>=95 and p<=100):
        sangat=1
    else :
        sangat = 0
    return te, biasa, enak, tp, puas,sangat

def inference(te, biasa, enak, tp, puas,sangat):
    aturan = [
        ['tidak',min(te,tp)],
        ['tidak',min(te,puas)],
        ['tidak',min(te,sangat)],
        ['tidak',min(biasa,tp)],
        ['mungkin',min(biasa,puas)],
        ['mungkin',min(biasa,sangat)],
        ['mungkin',min(enak,tp)],
        ['ya',min(enak,puas)],
        ['ya',min(enak,sangat)]
    ]
    
    mungkin = [] 
    tidak = [] 
    ya = [] 

    for i in range(len(aturan)):
        if aturan[i][0] == 'tidak':
            tidak.append(aturan[i][1])
        elif aturan[i][0] == 'mungkin':
            mungkin.append(aturan[i][1])
        elif aturan[i][0] == 'ya':
            ya.append(aturan[i][1])
    return max(ya), max(mungkin), max(tidak)

def defuzzyfication(ya,mungkin,tidak):
    defuzz = []
    for i in range(10):
        n = 10*(i+1)
        
        if 0<=n and n<=35:
            rej = 1
        elif 35<n and n<=55:
            rej = (-(n-55))/(55-35)
        else:
            rej = 0
        
        if 35<n and n<55 :
            may = (n-35)/(55-35)
        elif (n>=55 and n<=55):
            may =1
        elif 55<n and n<=75:
            may = (-(n-75))/(75-55)
        else :
            may = 0
        
        if 55<n and n<=75 :
            acc = (n-55)/(75-55)
        elif 75<n and n<=100:
            acc = 1
        else :
            acc = 0
        
        if (rej>tidak):
            rej = tidak
        if may>mungkin:
            may = mungkin
        if acc>ya:
            acc=ya

        h = [n,max(rej,may,acc),(n*max(rej,may,acc))]
        defuzz.append(h)
    miu = 0
    perkalian = 0
    for i in range(10):
        miu += defuzz[i][1]
        perkalian += defuzz[i][2]
    if (perkalian!=0 or miu!=0):
        return perkalian/miu
    else:
        return 0
        
hasil = []
for i in range(len(df)):
    #print(i)
    te, biasa, enak, tp, puas,sangat =fuzzification(df['pelayanan'][i], df['makanan'][i])
    ya,mungkin,tidak = inference(te, biasa, enak, tp, puas,sangat)
    defuzz = defuzzyfication(ya,mungkin,tidak)
    hasil.append([defuzz,i+1])

hasil.sort(reverse=True)
hasilakhir=[]
for i in range(10):
    hasilakhir.append(hasil[i][1])
df = pd.DataFrame(
    {
        'Peringkat Restoran': hasilakhir
        }
    )
writer = pd.ExcelWriter('peringkat.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
