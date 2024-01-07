# -*- coding: utf-8 -*-

import pandas


model = ['43XT745', '50P6US', '55XT515']
final_list = []
for model_name in model: 
 data = pandas.read_excel('1.xls', dtype = str)
 num_of_rows = len(data.index)
 for i in range(num_of_rows): 
    customer = data.loc[[i],['نام مدل']]
    customer2 = str(customer)
    customer3 = customer2.split()
    customer_mobile = data.loc[[i], ['تلفن سایت']]
    customer_mobile2 = str(customer_mobile).split()
    initial_list = []
    if customer3[3] == model_name and len(customer_mobile2[3]) == 11 and customer_mobile2[3].startswith('09') : 
        customer_mobile = data.loc[[i], ['تلفن سایت']]
        customer_mobile2 = str(customer_mobile).split()
        initial_list.append(customer_mobile2[3][1::])
        customer_name = data.loc[[i], ['نام مشتری']]
        customer_name2 = str(customer_name).split()
        customer_name3 = ''
        for i in range(3, len(customer_name2)):
            customer_name3 += customer_name2[i]
            customer_name3 += ' '
        initial_list.append(customer_name3)
        final_list.append(initial_list)
    elif customer3[3] == model_name and len(customer_mobile2[3]) > 11 and customer_mobile2[3].startswith('09') : 
            customer_mobile = data.loc[[i], ['تلفن سایت']]
            customer_mobile2 = str(customer_mobile).split()
            initial_list.append(customer_mobile2[3][1:11])
            customer_name = data.loc[[i], ['نام مشتری']]
            customer_name2 = str(customer_name).split()
            customer_name3 = ''
            for i in range(3, len(customer_name2)):
                customer_name3 += customer_name2[i]
                customer_name3 += ' '
            initial_list.append(customer_name3)
            final_list.append(initial_list)
            
 data = pandas.read_excel('2.xls', dtype = str)
 num_of_rows = len(data.index)

 for i in range(num_of_rows): 
    customer = data.loc[[i],['نام مدل']]
    customer2 = str(customer)
    customer3 = customer2.split()
    customer_mobile = data.loc[[i], ['تلفن سایت']]
    customer_mobile2 = str(customer_mobile).split()
    initial_list = []
    if customer3[3] == model_name and len(customer_mobile2[3]) == 11 and customer_mobile2[3].startswith('09') : 
        customer_mobile = data.loc[[i], ['تلفن سایت']]
        customer_mobile2 = str(customer_mobile).split()
        initial_list.append(customer_mobile2[3][1::])
        customer_name = data.loc[[i], ['نام مشتری']]
        customer_name2 = str(customer_name).split()
        customer_name3 = ''
        for i in range(3, len(customer_name2)):
            customer_name3 += customer_name2[i]
            customer_name3 += ' '
        initial_list.append(customer_name3)
        final_list.append(initial_list)
    elif customer3[3] == model_name and len(customer_mobile2[3]) > 11 and customer_mobile2[3].startswith('09') : 
            customer_mobile = data.loc[[i], ['تلفن سایت']]
            customer_mobile2 = str(customer_mobile).split()
            initial_list.append(customer_mobile2[3][1:11])
            customer_name = data.loc[[i], ['نام مشتری']]
            customer_name2 = str(customer_name).split()
            customer_name3 = ''
            for i in range(3, len(customer_name2)):
                customer_name3 += customer_name2[i]
                customer_name3 += ' '
            initial_list.append(customer_name3)
            final_list.append(initial_list)
            
            
 data = pandas.read_excel('3.xls', dtype = str)
 num_of_rows = len(data.index)

 for i in range(num_of_rows): 
    customer = data.loc[[i],['نام مدل']]
    customer2 = str(customer)
    customer3 = customer2.split()
    customer_mobile = data.loc[[i], ['تلفن سایت']]
    customer_mobile2 = str(customer_mobile).split()
    initial_list = []
    if customer3[3] == model_name and len(customer_mobile2[3]) == 11 and customer_mobile2[3].startswith('09') : 
        customer_mobile = data.loc[[i], ['تلفن سایت']]
        customer_mobile2 = str(customer_mobile).split()
        initial_list.append(customer_mobile2[3][1::])
        customer_name = data.loc[[i], ['نام مشتری']]
        customer_name2 = str(customer_name).split()
        customer_name3 = ''
        for i in range(3, len(customer_name2)):
            customer_name3 += customer_name2[i]
            customer_name3 += ' '
        initial_list.append(customer_name3)
        final_list.append(initial_list)
    elif customer3[3] == model_name and len(customer_mobile2[3]) > 11 and customer_mobile2[3].startswith('09') : 
            customer_mobile = data.loc[[i], ['تلفن سایت']]
            customer_mobile2 = str(customer_mobile).split()
            initial_list.append(customer_mobile2[3][1:11])
            customer_name = data.loc[[i], ['نام مشتری']]
            customer_name2 = str(customer_name).split()
            customer_name3 = ''
            for i in range(3, len(customer_name2)):
                customer_name3 += customer_name2[i]
                customer_name3 += ' '
            initial_list.append(customer_name3)
            final_list.append(initial_list)
 df = pandas.DataFrame(final_list, columns= ['Phone', 'Names'])
 mname = model_name + '.xlsx'
 df.to_excel(mname)
# print(final_list)
   
