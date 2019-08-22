#!/usr/bin/python

from tkinter import *
import datetime
import string
import random
from  xlwt import Workbook

class GarbageReporter(object):

    def __init__(self):
        self.masterGui = Tk()
        self.masterGui.title("Garbage Reporter 1.0")
        return

    def main(self):
        self.SetupLabels()
        self.SetupEntries()
        self.SetupOptionMenu()
        self.SetupButtons()
        
        self.masterGui.mainloop()
    
    def GenerateReport(self):
        #print ("report:")
       # print self.IdEntry.get()
        #print self.AmountEntry.get()
        #print self.FromEntry.get()
        #print self.ToEntry.get()
        #print self.FreqOptionValue.get()
        #print datetime.datetime(2018,3,13).today().weekday()
        List = self.GetBusinessDaysList(self.GetDayCountValue(self.FreqOptionValue.get()))
        #print(List)
        AmountList = self.generate_random_integers(int(self.AmountEntry.get()), len(List))
        #print(len(List))

        dictionary = dict(zip(List, AmountList))
        self.CreateExcelFile(dictionary)
        #print (List)
        #print (AmountList)

    def CreateExcelFile(self, outDict):
        workbook = Workbook()       
        worksheet = workbook.add_sheet("Sheet 1")
        row = 0
        column = 0
        
        for date, amount in (outDict.items()):
            worksheet.write(row, column, date)
            worksheet.write(row, column +1, amount)
            row += 1
        workbook.save("Report_" + self.IdEntry.get()+".xls")
        
    def generate_random_integers(self,_sum, n):  
        #print( _sum)
        #print (n)
        mean = (_sum) // n
        variance = int(0.25 * mean)

        min_v = mean - variance
        max_v = mean + variance
        array = [min_v] * n

        diff = _sum - min_v * n
        while diff > 0:
            a = random.randint(0, n - 1)
            if array[a] >= max_v:
                continue
            array[a] += 1
            diff -= 1
        #print (array)
        #print(sum(array))
        #print (len(array))
        return array

    def GetBusinessDaysList(self, day_count):
        StartDate = []
        EndDate = []
        DateList = []

        if self.ValidateDate(self.FromEntry.get()) and self.ValidateDate(self.ToEntry.get()):
            #print ("valid")
            StartDate = self.FromEntry.get().split(".")
            EndDate = self.ToEntry.get().split(".")
            
            StartDateValue = datetime.date(int(StartDate[2]), int(StartDate[1]), int(StartDate[0]))
            EndDateValue = datetime.date(int(EndDate[2]), int(EndDate[1]), int(EndDate[0]))
            
            DateList = self.CreateDateList(StartDateValue, EndDateValue, day_count)
        else:
            print ("not valid date format")
            
        return DateList

    def GetDayCountValue(self, freq_option):
        freq = 5 #denne
        if freq_option == "2x tyzdene":
            freq = 2
        elif freq_option == "3x tyzdene":
            freq = 3
            
        return freq

    def CreateDateList(self, start_date, end_date, day_count):
        WeekList = []
        OutList = []
        for single_date in (start_date + datetime.timedelta(n) for n in range((end_date - start_date).days +1)):
            if 0 <= single_date.weekday() <= 4:
                if single_date.weekday() == 0 or single_date == end_date:
                    if single_date == end_date:
                        WeekList.append(single_date)
                    #print ("week")
                    #print (WeekList)
                    random.shuffle(WeekList)
                #    OutList.append(WeekList[0:int(day_count)])
                    for listitem in WeekList[0:int(day_count)]:
                        OutList.append(listitem)
                    WeekList = []

                WeekList.append(single_date)
            elif single_date == end_date:
                # print (single_date.strftime("%d.%m.%Y"))
                random.shuffle(WeekList)
                # print ("weekend")
                # print (WeekList)
                #    OutList.append(WeekList[0:int(day_count)])
                for listitem in WeekList[0:int(day_count)]:
                    OutList.append(listitem)
                #print (single_date.strftime("%d.%m.%Y"))
        SortedList= []
        for item in sorted(OutList):
            SortedList.append(item.strftime("%d/%m/%Y"))
        return SortedList
    
    def ValidateDate(self, date):
        IsValid = False
        try:
            datetime.datetime.strptime(str(date), '%d.%m.%Y')
            IsValid = True
        except ValueError:
            print ("not valid date")
            
        return IsValid
        
    def FreqValueCallback(self,*args):
        print (self.FreqOptionValue.get())
    
    def SetupLabels(self):
        Label(self.masterGui, text= "Kod odpadu:").grid(row=0)
        Label(self.masterGui, text= "Mnozstvo:").grid(row=1)
        Label(self.masterGui, text= "Od(DD.MM.RRRR):").grid(row=2, column=0)
        Label(self.masterGui, text= "Do(DD.MM.RRRR):").grid(row=2, column=2)
        Label(self.masterGui, text= "Frekvencia:").grid(row=4, column=0)
        
    def SetupEntries(self):
        self.IdEntry= Entry(self.masterGui)
        self.AmountEntry = Entry(self.masterGui)
        self.FromEntry = Entry(self.masterGui)
        self.ToEntry = Entry(self.masterGui) 
        
        self.IdEntry.grid(row=0, column=1)
        self.AmountEntry.grid(row=1, column=1)
        self.FromEntry.grid(row=2, column=1)
        self.ToEntry.grid(row=2, column=3)
        
    def SetupButtons(self):
        self.GenerateButton = Button(self.masterGui, text="Generuj Vystup", command=self.GenerateReport)
        self.GenerateButton.grid(row=6, column=0)

        self.ExitButton = Button(self.masterGui, text="Exit", command=self.masterGui.quit)
        self.ExitButton.grid(row=6, column=3)
        
    def SetupOptionMenu(self):
        self.FreqOptionValue = StringVar(self.masterGui)
        self.FreqOptionValue.set("denne")
        self.FreqOptionValue.trace("w", self.FreqValueCallback)
        self.FreqOptionMenu= OptionMenu(self.masterGui, self.FreqOptionValue, "denne", "2x tyzdene","3x tyzdene")
        
        self.FreqOptionMenu.grid(row=4, column=1)
        self.FreqOptionMenu.configure(width=12)
    
if __name__ == "__main__":

    master = GarbageReporter()
    master.main()
    
