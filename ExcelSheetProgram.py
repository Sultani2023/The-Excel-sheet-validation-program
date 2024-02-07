import pandas as pd
# import numpy as np


Depart=[[1,'A',8,6,100],
        [2,'B',5,3,10],
        [3,'C',9,6,1000],
        [4,'D',8,6,50]]


print("____________________________________________________________________________")
print("****************************************************************************")
print("****************************************************************************")
print("******************** Welcome To Excel Validation Program *******************")
print("****************************************************************************")
print("****************************************************************************")
print("____________________________________________________________________________")
depart_df = pd.DataFrame(Depart,columns=['id','education','max_hours','min_hours','cost_per_hour' ])
print("****************************************************************************")
print("*********************** Department List *************************************")
print("****************************************************************************")
print(depart_df)
print("****************************************************************************")
print("____________________________________________________________________________")


# Read the employee data from an Excel file
print("____________________________________________________________________________")
print("****************************************************************************")
print("*********************** Excel Sheet Before Validation **********************")
print("****************************************************************************")  
employee_df = pd.read_excel('employee.xlsx')
print(employee_df)
print("****************************************************************************")
print("____________________________________________________________________________")


 

for index, row in employee_df.iterrows():

     firstName= row['first name']
     lastName= row['last name']
     age= row['age']
     depart=row['department']
     work_hour=int(row['hour'])
     emp_education=row['education']
     emp_salary=int(row['salary'])
     invalid=False

     


     

     if not firstName.isalpha():
        print(f"Error:first name '{firstName}'must be letters.")
        invalid=True
        
     if not lastName.isalpha():
       print(f"Error:Last name '{lastName}'must be letters.")
       invalid=True
       
      
     if 60 >= age <=20:
       print(f"Erorr: age'{age}' must be between 20 and 60 years.")
       invalid=True
       
      
     if  depart not in depart_df['id'].to_list():
        print(f"Error! Invalid department id '{depart}' the department must be one up to four.")
        invalid=True
        
      
     if emp_education not in depart_df['education'].to_list():
         print(f"Error! Invalid grade of education '{emp_education}' the education must be in A up to D ")
         invalid=True
      
         

     if not ((depart_df['min_hours'] <= work_hour) & (work_hour <= depart_df['max_hours'])).any():
        print(f"Error! Invalid work hours '{work_hour}' the work hours must be between(max hours & min hours) ")
        invalid=True
       

   #   if not np.isclose(emp_salary, work_hour * depart_df['cost_per_hour']).any():
   #      print(f"invalid  employee salary '{emp_salary}' " )
   #      print(emp_salary,work_hour,depart_df['cost_per_hour'])
   #      invalid=True
        if ((emp_salary != (work_hour * depart_df['cost_per_hour']))).any():
          print(f"Error! Invalid  employee salary '{emp_salary}'" )
          invalid=True
          user_input=input("Do you want to delete this invalid value? y/n ")
          if user_input=='y':
              print("The Excel Sheet Validation Completed Successfully.")
          
          else:
             print("Error! validation has been failed!")
             exit()
             

     if invalid==True:
          employee_df=employee_df.drop(index)
   
     
print("____________________________________________________________________________")
print("****************************************************************************")

employee_df.to_excel('employee15.xlsx',index=False)
print("****************************************************************************")
print("************************ Excel Sheet after Validation **********************")
print("****************************************************************************")  
new_employee_df = pd.read_excel('employee15.xlsx')
print(new_employee_df)
print("****************************************************************************")
print("____________________________________________________________________________")

















































