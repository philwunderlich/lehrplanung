# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
import pandas as pd
import sys
import os
print(sys.version)
os.chdir(sys.path[0])

## HEADER DATA ENTRY
term = "Sommersemester 2023"
filename = "Lehrplanung-S23-Master.xlsx"

## read LP Table1
data = pd.read_excel(filename, sheet_name="Lehrangebot")

metadata = pd.read_excel(filename, sheet_name="Metadaten")



cats = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', 'NN']
data['Tag'] = pd.Categorical(data['Tag'], categories=cats, ordered=True)
data = data.sort_values(by=['Tag', 'Von'])

data['Von'] = pd.to_numeric(data["Von"])
data['Bis'] = pd.to_numeric(data["Bis"])

# HTML String
m4 = data[data['Modul']=="M4"]
m5 = data[data['Modul']=="M5"]
m6 = data[data['Modul']=="M6"]
m8 = data[data['Modul']=="M8"]
mc = data[data['Modul']=="Colloquium"]

# read css file

css = open("style.css", "r")


html = f"""<meta http-equiv="content-type" content="text/html; charset=utf-8">
<meta charset="utf-8" >
<!--<link rel="stylesheet" href="style.css">-->
<style>
{css.read()}
</style>
<!--End of Style sheet-->
<body>
<div class="content">
<h1>{metadata.loc[0,"Heading"]} (Version {metadata.loc[0,"Version"]})<h1>
<h2>Semester 2</h2>
<table class="modultable">
<caption class="m4">Module 4: Social Structure and Inequalities in European Societies</caption>
<thead><tr>
<th>Title</th>
<th>Lecturer</th>
<th>Day</th>
<th>Time</th>
</tr></thead>"""

for index,row in m4.iterrows():
     if pd.notnull(row['Dozent2']):
        dozent2 = f" & {row['Dozent2']}" 
     else:
        dozent2 = ""
     html += f"<tr><td class='title-cell'>{row['Titel']}</td><td>{row['Dozent']}{dozent2}</td><td>{row['Tag']}</td><td>{int(row['Von'])}-{int(row['Bis'])}</td></tr>"

html += "</table>"

html += """<table class="modultable" style="break-before:auto">
<caption class="m5">Module 5: Values and Culture from a European Comparative Perspective</caption>
<thead><tr>
<th>Title</th>
<th>Lecturer</th>
<th>Day</th>
<th>Time</th>
</tr></thead>"""

for index,row in m5.iterrows():
     if pd.notnull(row['Dozent2']):
        dozent2 = f" & {row['Dozent2']}" 
     else:
        dozent2 = ""
     html += f"<tr><td class='title-cell'>{row['Titel']}</td><td>{row['Dozent']}{dozent2}</td><td>{row['Tag']}</td><td>{int(row['Von'])}-{int(row['Bis'])}</td></tr>"
html += "</table>"

html += """<table class="modultable" style="break-before:auto">
<caption class="m6">Module 6: Globalization and Regional Development</caption>
<thead><tr>
<th>Title</th>
<th>Lecturer</th>
<th>Day</th>
<th>Time</th>
</tr></thead>"""

for index,row in m6.iterrows():
     if pd.notnull(row['Dozent2']):
        dozent2 = f" & {row['Dozent2']}" 
     else:
        dozent2 = ""
     html += f"<tr><td class='title-cell'>{row['Titel']}</td><td>{row['Dozent']}{dozent2}</td><td>{row['Tag']}</td><td>{int(row['Von'])}-{int(row['Bis'])}</td></tr>"
html += "</table>"


html += """<p> </p>
<h2 style="break-before:always">Semester 4</h2>
<table class="modultable">
<caption  class="m8">Module 8: Specialisation</caption>
<thead><tr>
<th>Title</th>
<th>Lecturer</th>
<th>Day</th>
<th>Time</th>
</tr></thead>"""

for index,row in m8.iterrows():
     html += f"<tr><td class='title-cell'>{row['Titel']}</td><td>{row['Dozent']}</td><td>{row['Tag']}</td><td>{int(row['Von'])}-{int(row['Bis'])}</td></tr>"
html += "</table>"


html += """<table class="modultable">
<caption  class="mc">Colloquia</caption>
<thead><tr>
<th>Title</th>
<th>Lecturer</th>
<th>Day</th>
<th>Time</th>
</tr></thead> """

for index,row in mc.iterrows():
     html += f"<tr><td class='title-cell'>{row['Titel']}</td><td>{row['Dozent']}</td><td>{row['Tag']}</td><td>{int(row['Von'])}-{int(row['Bis'])}</td></tr>"
html += "</table>"


### include ready made timetable
#timetable = open("/Users/pwunderlich/Documents/00-FU/LEHRPLANUNG/timetable_s22.html", "r")
#html += timetable.read()





#### tests

print(data)

# count overlaps per event for one day and semester
# returns dataframe
def daily_data(data, day, semester):
     modules = []
     if semester==2:
          modules = ["M4", "M5", "M6"]
     if semester==4:
          modules = ["M8", "Colloquium"]
     daydata = data[(data["Tag"]==day) & (data["Modul"].isin(modules))]
     def overlaps(row, data):
          count = -1
          for i,j in daydata.iterrows():
               start = j["Von"]
               end = j["Bis"]
               start1 = row["Von"]
               end1 = row["Bis"]
               if ((start < end1) and (end > start1)):
                    count += 1
          return(count)
     col = daydata.apply(lambda row: overlaps(row, daydata), axis=1)
     daydata =  daydata.assign(overlaps = col.values)
     return(daydata)

# amount of parallel events.
def tot_overlap(data):
     x = data["overlaps"].max()
     x += 1
     if x == 3:
          x = x*2
     if x == 4:
          x = x*3
     return(x)

tot_overlap(daily_data(data, "Mon", 2))


# create a function that generates rows with all the many rules

def timeslot(day, time, semester, data):
     df = daily_data(data, day, semester)
     daily_range = tot_overlap(df)
     n_courses = len(df.loc[df['Von']==time,])
     string = ""
     # if n_courses == 0:
     #      string += f"<td colspan='{daily_range}'></td>" # nicht mehr benötigt da unten inbegriffen
     ## count secondary overlaps
     counter_secondary = 0
     for index, row in df.iterrows():
          if ((row["Von"] < time) and (row["Bis"]>time)):
               counter_secondary += 1
     #create cells and count
     counter_colspan = 0
     for i,j in df[df['Von']==time].iterrows():
          #print(j)
          rowspan = int((j["Bis"] - j["Von"])*2)
          # hier den column span aus der daily range und der zahl der überlappenden kurse berechnend
          # oder eben aus counter secondary und n_courses
          # problem, beides gibt nicht die maximale zahl gleichzeitiger Überlappungen des Kurses an.
          # das ist übrigens auch ein Problem bei der definition der header breite. Wenn fünf kurse überlappen heißt das dennoch nicht, dass man 5 spalten braucht.
          colspan = daily_range / (j['overlaps']+1)
          if pd.notnull(j['Dozent2']):
               dozent2 = f" & {j['Dozent2']}"
          else: 
               dozent2 = ""
          string += f"""
          <td colspan="{colspan}" rowspan="{rowspan}" class="{j['Modul']}" > {j['Titel']} <span>{j['Dozent']}{dozent2}</span></td>
          """
          counter_colspan += colspan
     if ((counter_colspan + counter_secondary) < daily_range):
          empties = daily_range - counter_colspan - counter_secondary
          string += f'<td colspan="{empties}"></td>'
     return(string)


# first table

semester = 2
semester2 = f"""<p> </p>
<p> </p>
<div class="timetable">
<table style="break-before:always">
<caption>Semester {semester}</caption>
<thead style="font-weight: bold;">
<tr>
<th><strong>Time</strong></th>"""

days = ["Mon", "Tue", "Wed", "Thu", "Fri"]

# create a row of headers for all weekdays
for day in days:
     daily = daily_data(data, day, 2)
     range_header = tot_overlap(daily)
     semester2 += f'<th colspan="{range_header}"><strong>{day}</strong></th>'


# semester2 += """
#     <tr>
#       <th>08:00</th>
# """

# for day in days:
#      semester2 += timeslot(day, 8,2,data)

# semester2 += """
#     </tr>
#     <tr>
#       <th>08:30</th>
#     </tr>
#     <tr>
#       <th>09:00</th>
#     </tr>
#     <tr>
#       <th>09:30</th>
#     </tr>
# """

semester2 += """<tr>
<th>10:00</th>"""

for day in days:
     semester2 += timeslot(day, 10,2,data)

semester2 += """</tr>
<tr>
<th>10:30</th>
</tr>
<tr>
<th>11:00</th>
</tr>
<tr>
<th>11:30</th>
</tr>
<tr>
<th>12:00</th>"""

for day in days:
     semester2 += timeslot(day, 12,2,data)


semester2 += """</tr>
<tr>
<th>12:30</th>
</tr>
<tr>
<th>13:00</th>
</tr>
<tr>
<th>13:30</th>
</tr>
<tr>
<th>14:00</th>"""

for day in days:
     semester2 += timeslot(day, 14,2,data)

semester2 += """</tr>
<tr>
<th>14:30</th>
</tr>
<tr>
<th>15:00</th>
</tr>
<tr>
<th>15:30</th>
</tr>
<tr>
<th>16:00</th>"""

for day in days:
     semester2 += timeslot(day, 16,2,data)

semester2 += """</tr>
<tr>
<th>16:30</th>
</tr>
<tr>
<th>17:00</th>
</tr>
<tr>
<th>17:30</th>
</tr>
<tr>
<th>18:00</th>"""

for day in days:
     semester2 += timeslot(day, 18,2,data)

semester2 += """</tr>
<tr>
<th>18:30</th>
</tr>
<tr>
<th>19:00</th>
</tr>
<tr>
<th>19:30</th>
</tr>"""
semester2 += "</table>"


print(semester2)
html += semester2




# second table

semester = 4
semester4 = f"""<table style="break-before:always">
<caption>Semester {semester}</caption>
<thead style="font-weight: bold;">
<tr>
<th><strong>Time</strong></th>"""

days = ["Mon", "Tue", "Wed", "Thu", "Fri"]

# create a row of headers for all weekdays
for day in days:
     daily = daily_data(data, day, 4)
     range_header = tot_overlap(daily)
     semester4 += f'<th colspan="{range_header}"><strong>{day}</strong></th>'


# semester4 += """
#     <tr>
#       <th>08:00</th>
# """

# for day in days:
#      semester4 += timeslot(day, 8,4,data)

# semester4 += """
#     </tr>
#     <tr>
#       <th>08:30</th>
#     </tr>
#     <tr>
#       <th>09:00</th>
#     </tr>
#     <tr>
#       <th>09:30</th>
#     </tr>
# """

semester4 += """
     <tr>
     <th>10:00</th>
"""

for day in days:
     semester4 += timeslot(day, 10,4,data)

semester4 += """</tr>
<tr>
<th>10:30</th>
</tr>
<tr>
<th>11:00</th>
</tr>
<tr>
<th>11:30</th>
</tr>
<tr>
<th>12:00</th>"""

for day in days:
     semester4 += timeslot(day, 12,4,data)


semester4 += """</tr>
<tr>
<th>12:30</th>
</tr>
<tr>
<th>13:00</th>
</tr>
<tr>
<th>13:30</th>
</tr>
<tr>
<th>14:00</th>"""

for day in days:
     semester4 += timeslot(day, 14,4,data)

semester4 += """</tr>
<tr>
<th>14:30</th>
</tr>
<tr>
<th>15:00</th>
</tr>
<tr>
<th>15:30</th>
</tr>
<tr>
<th>16:00</th>"""

for day in days:
     semester4 += timeslot(day, 16,4,data)

semester4 += """</tr>
<tr>
<th>16:30</th>
</tr>
<tr>
<th>17:00</th>
</tr>
<tr>
<th>17:30</th>
</tr>
<tr>
<th>18:00</th>"""

for day in days:
     semester4 += timeslot(day, 18,4,data)

semester4 += """</tr>
<tr>
<th>18:30</th>
</tr>
<tr>
<th>19:00</th>
</tr>
<tr>
<th>19:30</th>
</tr>"""
semester4 += "</table></div>"


html += "<p></p>"
html += semester4

html += "</div></body>"

# Write HTML String to file.html
with open("summer-term.html", "w") as file:
     file.write(html)

#### export dekanat table

print(data)

import numpy as np

df = data

# create a list of our conditions
conditions = [
    (df['Modul'] == "M4") & (df["Typ"] == "VL"),
    (df['Modul'] == "M4") & (df["Typ"] == "Sem"),
    (df['Modul'] == "M5") & (df["Typ"] == "VL"),
    (df['Modul'] == "M5") & (df["Typ"] == "Sem"),
    (df['Modul'] == "M6") & (df["Typ"] == "VL"),
    (df['Modul'] == "M6") & (df["Typ"] == "Sem"),
    (df['Modul'] == "M8"),
    (df['Modul'] == "Colloquium")
    ]

# create a list of the values we want to assign for each condition
values = [
     '0181bB.1.1.1',
     '0181bB.1.1.2',
     '0181bB.1.2.1',
     '0181bB.1.2.2',
     '0181bB.1.3.1',
     '0181bB.1.3.2',
     '0181bC.1.2.1',
     '0181bE.1.2.1']

# create a new column and use np.select to assign values to it using our lists as arguments
df['Lehre'] = np.select(conditions, values)

values2 = [
     'Vertiefungsvorlesung',
     'Hauptseminar',
     'Vertiefungsvorlesung',
     'Hauptseminar',
     'Vertiefungsvorlesung',
     'Hauptseminar',
     'Vertiefungsseminar',
     'Kolloquium']

df['LV-Art'] = np.select(conditions, values2)

# display updated DataFrame
df.dtypes
df = df.rename(columns={"Nummer": "LV-Nummer",
     "Dozent": "Dozent*in"})

df1 = df.reindex(columns=[
     "Lehre", "LV-Nummer",
     "LV-Art", "SWS",
     "Dozent*in", "Titel",
     "Bemerkungen LA 1",
     "Bemerkungen LA 2",
     "Lehrauftrag besoldet",
     "Lehrauftrag unbesoldet",
     "Titellehre", "TN Max",
     "Block", "Präsenz",
     "online-synchron",
     "online-asynchron",
     "Grün", "Gelb", "Rot",
     "LV geöffnet für Geflüchtete"
     ]).sort_values(by=["Lehre"])

print(df)

## Sheet 2 (tn beschränkt)
print(df.dtypes)
df = df.astype({"LV-Nummer":"string"})

# filter for only "platzbeschränkt" and select other columns, and sort by number.
df2 = df[df["Platzbeschränkt"]=="ja"].reindex(columns=[
     "LV-Art",
     "LV-Nummer",
     "Dozent*in",
     "Titel",
     "TN Max",
     "Begründung"
]).sort_values(by=["LV-Nummer"])

## Sheet 3 (Blockveranstaltungen)

df3 = df[df["Block"]=="ja"].reindex(columns=[
     "LV-Art",
     "LV-Nummer",
     "Dozent*in",
     "Titel",
     "Durchführungsturnus mit Zeit/Ort/Datum und Pausenzeiten",
     "Begründung"
]).sort_values(by=["LV-Nummer"])

df.dtypes
df4 = df[(df["Lehrauftrag unbesoldet"]=="ja") | (df["Lehrauftrag besoldet"]=="ja")].reindex(columns=[
     "LV-Art",
     "LV-Nummer",
     "Dozent*in",
     "Titel",
     "im Zuge des Aufwuchses",
     "Bemerkungen/Finanzierung (MvBZ alternativ. Finanzierung?)"
]).sort_values(by=["LV-Nummer"])

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Lehrplanung_IfS_SS22_fuer_Dekanat.xlsx', engine='xlsxwriter')

# Write the dataframe data to XlsxWriter. Turn off the default header and
# index and skip one row to allow us to insert a user defined header.
df1.to_excel(writer, sheet_name='alle LVen nach Modulen', startrow=1, header=True, index=False)
df2.to_excel(writer, sheet_name='teilnahmebeschränkte LVen', startrow=1, header=True, index=False)
df3.to_excel(writer, sheet_name='Blockveranstaltungen', startrow=1, header=True, index=False)
df4.to_excel(writer, sheet_name='Lehraufträge', startrow=1, header=True, index=False)


# Get the xlsxwriter workbook and worksheet objects.
#workbook = writer.book
#worksheet = writer.sheets['alle LVen nach Modulen', 'teilnahmebeschränkte LVen']

# Close the Pandas Excel writer and output the Excel file.
writer.save()


## convert to pdf


# import pdfkit

# options = {
#     'page-size': 'A4',
#     'margin-top': '2cm',
#     'margin-right': '2cm',
#     'margin-bottom': '2cm',
#     'margin-left': '2cm',
#     'encoding': "UTF-8"
#     }
html

# pdfkit.from_file('summer-term.html', 'summer-term.pdf', options=options)

#from weasyprint import HTML

#HTML(string=html).write_pdf("summer-term.pdf")
