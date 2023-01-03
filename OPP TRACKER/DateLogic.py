# Import date class from datetime module

from datetime import datetime
from datetime import date
from datetime import timedelta



    
def  Is_Weekend(PythonDateString,PythonReminderFrequency):
    
#PythonDateString="12-25-2022"    
    PythonDateTime=datetime.strptime(PythonDateString,"%d-%m-%Y")
#PythonReminderFrequency=1
    ReminderPythonDateTime=PythonDateTime + timedelta(days=PythonReminderFrequency)


    day=ReminderPythonDateTime.weekday()
    print(day)
    if(day<=4):
        print("Weekday  ",day)
        print(ReminderPythonDateTime)
        CalculatedDate=ReminderPythonDateTime
        return  CalculatedDate
    
    else:
        if(day==5):
            print("Saturday =  ",day)
            NextWeekdaydate=ReminderPythonDateTime + timedelta(days=2)
            print(NextWeekdaydate)
            CalculatedDate=NextWeekdaydate
            return  CalculatedDate
            
        else:
            print("Sunday = ",day)
            NextWeekdaydate=ReminderPythonDateTime + timedelta(days=1)
            print(NextWeekdaydate) 
            CalculatedDate=NextWeekdaydate
            return  CalculatedDate