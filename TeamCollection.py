import tbapy, csv

tba = tbapy.TBA('GZ9erQoKBsgLfKlXDGClNWrdyVMq5Tpk3qab732jxDwaHAyAsL67WHl5Y8e5UtXY')

currentYear = 2024

teams = tba.teams()
rows = []
state = 1
for i in teams:
    print(i.team_number)
    years = tba.team_years(i.team_number)

    try:
        started = years[0]
        last = years[-1]
        competed = last - started + 1
        if len(years) != (last - started + 1):
            for j in range(len(years)):
                if years[j+1] > years[j] + 1:
                    last = years[j]
                    returned = years[j+1]
                    break
        else:
            returned = 0
        years_range = [0] * (currentYear-1993+1)
        for play_year in range(1993,currentYear,1):
            if play_year in years:
                years_range[play_year-1993] = 1
            if started > play_year:
                years_range[play_year-1993] = float("nan")
        row = [i.team_number,i.city, i.state_prov, i.postal_code,i.lat,i.lng, started]
        row.extend(years_range)
    except:
        print(".", end="")
    rows.append(row)
header = ["team number", "city", "state", "zip code", "latitude", "longitude", "rookie year", "1993", "1994", "1995", "1996", "1997", "1998", "1999","2000",
          "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017",
          "2018", "2019", "2020", "2021", "2022", "2023", "2024" ]
with open("AllTeamData_years.csv", 'w', encoding="utf-8", newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter=",")
    writer.writerow(header)
    writer.writerows(rows)

print("Done!")