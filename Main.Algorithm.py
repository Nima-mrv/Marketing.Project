# -*- coding: utf-8 -*-
from datetime import datetime, timedelta

def subtract_days(date_str, days):
    date = datetime.strptime(date_str, '%Y/%m/%d')
    new_date = date - timedelta(days=days)
    return new_date.strftime('%Y/%m/%d')


def is_between(date, start, end):
    date = datetime.strptime(date, '%Y/%m/%d')
    start = datetime.strptime(start, '%Y/%m/%d')
    end = datetime.strptime(end, '%Y/%m/%d')
    return start <= date <= end


import pandas

def duration(date1, date2):
    
    year1 = int(date1[:4])
    month1 = int(date1[5:7])
    day1 = int(date1[8:10])
    year2 = int(date2[:4])
    month2 = int(date2[5:7])
    day2 = int(date2[8:10])
    day_dif = 0
    if year1 == year2 :
        if month1 == month2:
            day_dif += day2 - day1
        else:
            day_dif += ((month2 - month1) - 1) * 30 + (30 - day1) + day2
        return day_dif
    else:
        year_dif = year2 - year1
        day_dif += (year_dif - 1) * 364 + ((12-month1)*30) + (30 - day1) + (month2 - 1) * 30 + day2
        return day_dif
    



df = pandas.read_excel('Reports1402.xlsx', dtype= str)
final_dict = {}
priority = {}
nor = len(df.index)
for i in range(nor):
    service = df.loc[[i], ['نوع تماس']]
    service2 = str(service).split()
    if service2[3] == 'تعمیر' :
        model = df.loc[[i], ['نام مدل']]
        model2 = str(model).split()
        if model2[3] in priority:
            priority[model2[3]] += 1
        else:
            priority[model2[3]] = 1
        install_date = df.loc[[i], ['تاریخ شروع گارانتی']]
        install_date2 = str(install_date).split()
        repair_date = df.loc[[i], ['تاریخ پذیرش']]
        repair_date2 = str(repair_date).split()
        if len(install_date2[4]) == 10 and len(repair_date2[3]) == 10:
            duration1 = duration(install_date2[4], repair_date2[3])
            if model2[3] in final_dict and duration1 >= 0:
                final_dict[model2[3]][0].append(duration1)
                final_dict[model2[3]][1] += 1
            if model2[3] not in final_dict and duration1 >= 0:
                final_dict[model2[3]] = [[duration1], 1]
        else:
            continue
        
for i in final_dict:
    mid = sum(final_dict[i][0]) // len(final_dict[i][0]) 
    final_dict[i][0].clear()
    final_dict[i][0] = mid

final_dict2 = dict(sorted(final_dict.items(), key=lambda item : item[1][1]))
final3 = {}
for i in final_dict2:
    if final_dict2[i][1] >= 10:
        final3[i] = final_dict2[i]
        
print(final3)
            
        
    
finallist = []
for j in range(nor):
    initiallist = []
    service = df.loc[[j], ['نوع تماس']]
    service2 = str(service).split()
    if service2[3] == 'نصب' :
        model = df.loc[[j], ['نام مدل']]
        model2 = str(model).split()
        for t in final3:
            
            if model2[3] == t:
                subtract1 = subtract_days('1402/05/17', final3[t][0])
                subtract2 = subtract_days('1402/08/17', final3[t][0])
                if subtract1[5:7] == '02' and int(subtract1[8:10]) > 28:
                    subtract1 = subtract1.replace(subtract1[8:10], '28')
                elif subtract1[5:7] in ['04', '06', '09', '11'] and int(subtract1[8:10]) > 30:
                    subtract1 = subtract1.replace(subtract1[8:10], '30')
                if subtract2[5:7] == '02' and int(subtract2[8:10]) > 28:
                    subtract2 = subtract2[4].replace(subtract2[8:10], '28')
                elif subtract2[5:7] in ['04', '06', '09', '11'] and int(subtract2[8:10]) > 30:
                     subtract2 = subtract2.replace(subtract2[8:10], '30')
                
                install_date = df.loc[[j], ['تاریخ شروع گارانتی']]
                install_date2 = str(install_date).split()
                if install_date2[4][5:7] == '02' and int(install_date2[4][8:10]) > 28:
                    newdate = install_date2[4].replace(install_date2[4][8:10], '28')
                    install_date2.remove(install_date2[4])
                    install_date2.insert(4, newdate)
                elif install_date2[4][5:7] in ['04', '06', '09', '11'] and int(install_date2[4][8:10]) > 30:
                     newdate = install_date2[4].replace(install_date2[4][8:10], '30')
                     install_date2.remove(install_date2[4])
                     install_date2.insert(4, newdate) 
                if len(install_date2[4]) == 10:
                    if is_between(install_date2[4], subtract1, subtract2) == True: 
                      customer_mobile = df.loc[[j], ['موبایل سایت']]
                      customer_mobile2 = str(customer_mobile).split()
                      initiallist.append(customer_mobile2[3][1:11])
                      customer_name = df.loc[[j], ['نام مشتری']]
                      customer_name2 = str(customer_name).split()
                      customer_name3 = ''
                      for i in range(3, len(customer_name2)):
                          customer_name3 += customer_name2[i]
                          customer_name3 += ' '
                      initiallist.append(customer_name3)
                      end_g = df.loc[[j], ['تاریخ پایان گارانتی']]
                      end_g2 = str(end_g).split()
                      initiallist.append(end_g2[4])
                      brand = str(df.loc[[j], ['برند']]).split()
                      brand2 = ''
                      for z in range(2,len(brand)):
                          brand2 += brand[z]
                          brand2 += '\u200b'
                      initiallist.append(brand2)
                      finallist.append(initiallist)
                
df = pandas.DataFrame(finallist, columns=['Phone', 'Name', 'Dead Line', 'Brand'])
df.to_excel('GuaranteeCustomers2.xlsx')








