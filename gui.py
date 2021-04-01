import tkinter as tk
from tkinter import ttk
import pandas as pd


class MainRAD:

    def __init__(self, window):
        # Components
        self.labelName = tk.Label(window, text="Student name: ")
        self.labelGrade1 = tk.Label(window, text="Grade 1: ")
        self.labelGrade2 = tk.Label(window, text="Grade 2: ")
        self.labelAverage = tk.Label(window, text="Average: ")
        self.textName = tk.Entry(bd=3)
        self.textGrade1 = tk.Entry()
        self.textGrade2 = tk.Entry()
        self.buttonCalculate = tk.Button(window, text="Calculate average", command=self.calculateAverage)

        # Tree view
        self.columnData = ("Student", "Grade1", "Grade2", "Average", "Situation")
        self.treeMedias = ttk.Treeview(window, columns=self.columnData, selectmode="browse")

        self.scrollBar = ttk.Scrollbar(window, orient="vertical", command=self.treeMedias.yview)
        self.scrollBar.pack(side="right", fill="x")
        self.treeMedias.configure(yscrollcommand=self.scrollBar.set)

        # Components position
        self.labelName.place(x=100, y=50)
        self.textName.place(x=200, y=50)

        self.labelGrade1.place(x=100, y=100)
        self.textGrade1.place(x=200, y=100)

        self.labelGrade2.place(x=100, y=150)
        self.textGrade2.place(x=200, y=150)

        self.buttonCalculate.place(x=100, y=200)

        self.treeMedias.place(x=100, y=300)
        self.scrollBar.place(x=805, y=300, height=225)

        self.id = 0
        self.iid = 0

        self.loadData()

    def loadData(self):
        try:
            save = "students.xlsx"
            data = pd.read_excel(save)
            qtStudents = len(data['Student'])

            for i in range(qtStudents):
                name = data['Student'][i]
                grade1 = data['Grade1'][i]
                grade2 = data['Grade1'][i]
                average = data['Grade1'][i]
                situation = data['Situation'][i]

                self.treeMedias.insert('', 'end',
                                       iid= self.iid,
                                       values=(name, grade1, grade2, average, situation))

                self.iid = self.iid + 1
                self.id = self.id + 1


        except (Exception) as error:
            print(error)

    def calculateAverage(self):
        try:
            name = self.textName.get()
            grade1 = float(self.textGrade1.get())
            grade2 = float(self.textGrade2.get())

            average, situation = self.verifySituation(grade1, grade2)

            self.treeMedias.insert('', 'end',
                                   iid=self.iid,
                                   values=(name, str(grade1), str(grade2), str(average), situation))

        except ValueError as error:
            print(error)

        finally:
            self.textName.delete(0, 'end')
            self.textGrade1.delete(0, 'end')
            self.textGrade2.delete(0, 'end')

    def verifySituation(self, grade1, grade2):
        average = (grade1+grade2)/2

        if average >= 7:
            situation = "Approved"
        elif average >= 5:
            situation = "Summer school"
        else:
            situation = "Not approved"

        return average, situation

