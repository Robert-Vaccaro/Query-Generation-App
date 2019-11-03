import tkinter 
from tkinter import Tk, Label, Entry, W, E, S, N
import xlrd
import xlsxwriter
import tkinter.messagebox
from tkinter import simpledialog as simpledialog
import tkinter.simpledialog as simpledialog
from tkinter.messagebox import showinfo
from tkinter import messagebox, PhotoImage
import sys






root = Tk()
root.title("Query Creation")
root.geometry("500x300")


def starting(self):
    def makeLabel(query):
        if query != "":
            messagebox.showinfo("Error", "{} does not exist. Please try again.".format(query))
        if query == "":
            messagebox.showinfo("Error", "All Entry Boxes Are Required!\n\nThere cannot be anything left blank!")
    import requests
    answer3 = messagebox.askquestion("Creating Queries", "Are these the correct queries? \n\nIf so, press Yes.\n\n(Once you press yes, please wait a few moments for the queries to be generated)")
    if answer3 == 'yes':    
        firstArr=[]
        secondArr=[]
        thirdArr=[]

        def main(userInput):
            arr=[]
            query = userInput

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()
            
            datRes = data['response']
            
            dataLen = len(datRes)


            def separateString(strin):
                newStrin = strin.split('|') #making the new array, read below
                return newStrin 

            for i in range(dataLen):
                syno=data['response'][i]['list']['synonyms'] #iterating through JSON response
                newsplit = syno.split() #creating array from response
                if "|" in newsplit[0]: #if it has | in array item
                    newWords = separateString(newsplit[0]) #run it through this function to make new array of that element in original array
                    arrLen = len(newWords) #length of new Array
                    for z in range(arrLen):
                        arr.append(newWords[z]) #adding each new item from new array to original array if it has |

                else:
                    arr.append(newsplit[0]) #if it just a single word, it is appended to the original array


            arr = set(arr) #getting unique values
            arr=list(arr) # turning into list to be iterating through

            for y in arr:
                firstArr.append(y)
            
            print(firstArr)
        def main2(userInput):
            arr=[]
            query = userInput

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()
            
            datRes = data['response']
            
            dataLen = len(datRes)


            def separateString(strin):
                newStrin = strin.split('|') #making the new array, read below
                return newStrin 

            for i in range(dataLen):
                syno=data['response'][i]['list']['synonyms'] #iterating through JSON response
                newsplit = syno.split() #creating array from response
                if "|" in newsplit[0]: #if it has | in array item
                    newWords = separateString(newsplit[0]) #run it through this function to make new array of that element in original array
                    arrLen = len(newWords) #length of new Array
                    for z in range(arrLen):
                        arr.append(newWords[z]) 

                else:
                    arr.append(newsplit[0]) 


            arr = set(arr) #getting unique values
            arr=list(arr) # turning into list to be iterating through

            for y in arr:
                secondArr.append(y)
            
            print(secondArr)
        def main3(userInput):
            arr=[]
            query = userInput

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()
            
            datRes = data['response']
            
            dataLen = len(datRes)


            def separateString(strin):
                newStrin = strin.split('|') #making the new array, read below
                return newStrin 

            for i in range(dataLen):
                syno=data['response'][i]['list']['synonyms'] #iterating through JSON response
                newsplit = syno.split() #creating array from response
                if "|" in newsplit[0]: #if it has | in array item
                    newWords = separateString(newsplit[0]) #run it through this function to make new array of that element in original array
                    arrLen = len(newWords) #length of new Array
                    for z in range(arrLen):
                        arr.append(newWords[z]) 

                else:
                    arr.append(newsplit[0]) #if it just a single word, it is appended to the original array


            arr = set(arr) #getting unique values
            arr=list(arr) # turning into list to be iterating through

            for y in arr:
                thirdArr.append(y)
            
            print(thirdArr)

        def search(inputty):
            query = inputty

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()

            while "error" in data:
                makeLabel(query)
                print("'" + query + "'" + " does not exist in the dictionary")
                query = input("Please enter another search term: ")
                response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
                data = response.json()
                if "error" not in data:
                    print("Results for {}: ".format(query))
                    main(query)
                    print()
                    break
                else:
                    pass
            else:
                print("Results for {}: ".format(query))
                main(query)
                print()

        def search2(inputty2):
            query = inputty2

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()

            while "error" in data:
                makeLabel(query)
                print("'" + query + "'" + " does not exist in the dictionary")
                query = input("Please enter another search term: ")
                response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
                data = response.json()
                if "error" not in data:
                    print("Results for {}: ".format(query))
                    main2(query)
                    print()
                    break
                else:
                    pass
            else:
                print("Results for {}: ".format(query))
                main2(query)
                print()

        def search3(inputty3):
            query = inputty3

            response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
            data = response.json()

            while "error" in data:
                makeLabel(query)
                print("'" + query + "'" + " does not exist in the dictionary")
                query = input("Please enter another search term: ")
                response = requests.get("https://thesaurus.altervista.org/thesaurus/v1?word={}&language=en_US&output=json&key=9HgndzVU9vZ19KTtUEFn".format(query))
                data = response.json()
                if "error" not in data:
                    print("Results for {}: ".format(query))
                    main3(query)
                    print()
                    break
                else:
                    pass
            else:
                print("Results for {}: ".format(query))
                main3(query)
                print()

        inputty=query1.get()
        inputty2=query2.get()
        inputty3=query3.get()
        search(inputty)
        search2(inputty2)
        search3(inputty3)

        resultSet=set(firstArr)
        resultSet2=set(secondArr)
        resultSet3=set(thirdArr)
        resultList=list(resultSet)
        resultList2=list(resultSet2)
        resultList3=list(resultSet3)
        print("first ", resultList)
        print("second ", resultList2)
        print("third ", resultList3)
        totalLen = len(resultList)+len(resultList2)+len(resultList3)
        finalArr=[]
        d=1
        for a in resultList:
            for b in resultList2:
                for c in resultList3:
                    result = a+" "+b+" "+c
                    finalArr.append(result)
                    worksheet.write(d, 0, result)                   
                    d=d+1
        workbook.close()
        answer2 = messagebox.askquestion("Exiting Program", "Program has completed.\n\nExcel File has been create with all queries. \n\nPlease go to the folder in which this program is in and the excel will be there. \n\nClick Yes to Exit. \n\nClick No to redo queries.")
        if answer2 == 'yes':
            sys.exit()
            workbook.close()
            root.destroy()

def exitLast():
    answer = messagebox.askquestion("Exiting Program", "Would you like to exit the program?\nYour Progress will saved up to this point.\n\nClick Yes to Exit.")
    if answer == 'yes': 
        sys.exit()
        workbook.close()
        root.destroy()




workbook = xlsxwriter.Workbook('Query_Results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.add_table(0, 0, 199 , 2)

worksheet.set_column('A:A', 45, None)
worksheet.set_column('B:B', 45, None)

worksheet.add_table("A1:B1", {'columns': [{'header': 'Query'},
                                        {'header': 'Job Type'}
                                        ]})


hos_lab = Label(root, text = "Descriptor: ")
hos_lab.grid(row=10,column=0, sticky = 'E')

hos_lab2 = Label(root, text = "Job Type: ")
hos_lab2.grid(row=11,column=0, sticky = 'E')

hos_lab3 = Label(root, text = "Vertical: ")
hos_lab3.grid(row=12,column=0, sticky = 'E')

query1= Entry(root)
query1.grid(row=10, column=1, padx = 15, pady=3, sticky=W+E)

query1.focus()
query2= Entry(root)
query2.grid(row=11, column=1, padx = 15, pady=3, sticky=W+E)


query3= Entry(root)
query3.grid(row=12, column=1, padx = 15, pady=3, sticky=W+E)

hos_lab = Label(root, text = "Golden Set Query Generation")
hos_lab.grid(row=0, column=0+1, sticky = 'W')
hos_lab2 = Label(root, text = "Created By: Robert Vaccaro")
hos_lab2.grid(row=1, column=0+1)
hos_lab3 = Label(root, text = "Dated Created: November 2, 2019")
hos_lab3.grid(row=2, column=0+1)
hos_lab.config(font=("Courier", 24))
root.protocol("WM_DELETE_WINDOW", exitLast)
root.bind('<Return>', starting)
root.mainloop()