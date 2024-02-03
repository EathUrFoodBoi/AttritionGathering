import tbapy, xlsxwriter
tba = tbapy.TBA('GZ9erQoKBsgLfKlXDGClNWrdyVMq5Tpk3qab732jxDwaHAyAsL67WHl5Y8e5UtXY')
currentYear = input("What is the current year?")
allInfo = []
workbook = xlsxwriter.Workbook('Team_Attrition_Data_2024.xlsx')
worksheet = workbook.add_worksheet()

try:
    currentYear = int(currentYear)
except:
    print("Invalid number! Please restart program!")
    exit()


print("Getting the teams...")
teams = tba.teams()
print("Teams returned!")
def Earliest(teamNumber):
    teamYears = tba.team_years(teamNumber)
    earliestYear = teamYears[0]
    return earliestYear
def Latest(teamNumber):
    teamYears = tba.team_years(teamNumber)
    latestYear = teamYears[len(teamYears)-1]
    return latestYear
def stillActive(teamNumber):
    latestYear = int(Latest(teamNumber))
    if latestYear == currentYear:
        return True
    else:
        return False
def teamLeft(teamNumber):
    teamYears = tba.team_years(teamNumber)
    for i in range(len(teamYears) - 1):
        if not (int(teamYears[i+1] == int(len(teamYears)-1))):
            if not (int(teamYears[i+1]) == int(teamYears[i]) +1):
                return (teamYears[i])
    return False
def teamReturned(teamNumber):
    teamYears = tba.team_years(teamNumber)
    if not teamLeft(teamNumber) == False:
        yearLeft = teamLeft(teamNumber)
        indexOfYear = int(teamYears.index(yearLeft))
        yearReturned = teamYears[indexOfYear+1]
        return yearReturned
    else:
        return False
    
def getCountry(whereInList):
    country = teams[whereInList]["country"]
    return country

def getState(whereInList):
    state_prov = teams[whereInList]["state_prov"]
    return state_prov

def getCity(whereInList):
    city = teams[whereInList]["city"]
    return city

print("Writing info to worksheet...")
worksheet.write(0, 0, "Team Number")
worksheet.write(0, 1, "Year they Joined")
worksheet.write(0, 2, "Most recent year")
worksheet.write(0, 3, "Still active?")
worksheet.write(0, 4, "Did they take a break? When did it start?")
worksheet.write(0, 5, "When did they return?")
worksheet.write(0, 6, "Country")
worksheet.write(0, 7, "State/Province")
worksheet.write(0, 8, "City")
for i in range((len(teams) - 1)):
    teamNumber = int(teams[i]["team_number"])
    teamYears = tba.team_years(teamNumber)
    if not(len(teamYears) == 0):
        worksheet.write(i+1, 0, teamNumber)
        worksheet.write(i+1, 1, str(Earliest(teamNumber)))
        worksheet.write(i+1, 2, str(Latest(teamNumber)))
        worksheet.write(i+1, 3, str(stillActive(teamNumber)))
        worksheet.write(i+1, 4, str(teamLeft(teamNumber)))
        worksheet.write(i+1, 5, str(teamReturned(teamNumber)))
        worksheet.write(i+1, 6, str(getCountry(i)))
        worksheet.write(i+1, 7, str(getState(i)))
        worksheet.write(i+1, 8, str(getCity(i)))
    if i%50 == 0:
        print("The code is still running!")
    if i==8000:
        print("Almost done!")
workbook.close()
print("Done!")