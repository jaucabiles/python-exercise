import mysql.connector
from openpyxl import load_workbook
myDB = mysql.connector.connect(
  host="localhost",
  user="root",
  passwd="jaucabiles"
)


    
workbook_name = 'ticketsystem.xlsx'
book = load_workbook(workbook_name)
sheet = book.active

myCursor = myDB.cursor()
myCursor.execute("USE ticketSystem")


print("Ticketing System")

print("\n1. CI\n2. PR\n3. IN\n4. SR\n5. CR")

tktClassificationInpt = 0

def classificationChoice(tktClassification):
    print("\n" + tktClassification)
    chooseAction()

def chooseAction():
    print("\n1. ADD\n2. RETRIEVE\n3. EDIT")
    action = input("Choose an action: ")
    if action == "1":
        addToDatabase()
    
    elif action == "2":
        retrievefromDatabase()

    elif action == "3":
        editFromDatabase()

def addToDatabase():

    tktClass = tktClassification
    if tktClass == "CI":
        resultSet = "SELECT count(*) FROM tickets WHERE tktID LIKE 'CI%'"
    if tktClass == "PR":
        resultSet = "SELECT count(*) FROM tickets WHERE tktID LIKE 'PR%'"
    if tktClass == "IN":
        resultSet = "SELECT count(*) FROM tickets WHERE tktID LIKE 'IN%'"
    if tktClass == "SR":
        resultSet = "SELECT count(*) FROM tickets WHERE tktID LIKE 'SR%'"
    if tktClass == "CR":
        resultSet = "SELECT count(*) FROM tickets WHERE tktID LIKE 'CR%'"

    myCursor.execute(resultSet)
    valResultSet = myCursor.fetchone()
    getValue = str(valResultSet[0])
    ticketID = tktClassification + getValue
    print(ticketID)
    insertTicket = "INSERT INTO tickets tktID values (?)"

    query = "INSERT INTO tickets (tktID, tktDescription, tktAttachment) VALUES (%s, %s, %s)"
    tktDescription = input("Put a description: ")
    tktAttachment = input("Copy the path of the file you want to attach: ")
    items = (ticketID, tktDescription, tktAttachment)
    itemss = (ticketID, tktDescription)
    myCursor.execute(query, items)
    myDB.commit()
    
    for item in itemss:
        sheet.append([ticketID, tktDescription])
        
    book.save('ticketsystem.xlsx')
    


def retrievefromDatabase():
    query = "SELECT * from tickets"
    myCursor.execute(query)
    ticketList = myCursor.fetchall()

    print("TICKET LIST\n\n")
    for row in ticketList:
        print("Ticket ID: ", row[1])
        print("Description: ", row[2], "\n")

def editFromDatabase():
    query = "UPDATE tickets SET tktDescription = %s WHERE tktID = %s"
    idToBeChanged = input("Ticket ID: ")
    descToBeChanged = input("New Description: ")
    itemsToBeChanged = (descToBeChanged, idToBeChanged)
    myCursor.execute(query, itemsToBeChanged)
    myDB.commit()


tktClassificationInpt = input("Pick a category: ")

if tktClassificationInpt == "1":
    tktClassification = "CI"
    classificationChoice(tktClassification)
    
elif tktClassificationInpt == "2":
    tktClassification = "PR"
    classificationChoice(tktClassification)

elif tktClassificationInpt == "3":
    tktClassification = "IN"
    classificationChoice(tktClassification)

elif tktClassificationInpt == "4":
    tktClassification = "SR"
    classificationChoice(tktClassification)

elif tktClassificationInpt == "5":
    tktClassification = "CR"
    classificationChoice(tktClassification)

""" hello

"""