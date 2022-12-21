from parser import *
from openpyxl.chart import (
    PieChart,
    BarChart,
    Reference
)



def write_graph():
    rapport = load_workbook("../conformity-report/" + file_name + ".xlsx")
    all_sheet = rapport.worksheets[4]

    # PIE 1

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=1, max_row=2)
    data = Reference(all_sheet, min_col=2, min_row=1, max_row=2)
    pie.add_data(data)
    pie.set_categories(labels)

    if all_sheet['B1'].value != 0:
        title = int((100 * all_sheet['B2'].value) / (all_sheet['B1'].value + all_sheet['B2'].value) )
        pie.title = str(title) + " % DE POSTES NON CONFORMES"
    all_sheet.add_chart(pie, "E1")

    # PIE 2

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=3, max_row=4)
    data = Reference(all_sheet, min_col=2, min_row=3, max_row=4)
    pie.add_data(data)
    pie.set_categories(labels)
    num = all_sheet['B4'].value
    title = 100 - int((100 *  + all_sheet['B3'].value) / (all_sheet['B3'].value + all_sheet['B4'].value))
    pie.title = str(num) + " (" + str(title) + "%)" + " POSTES BUREAUTIQUES NON CONFORMES"

    all_sheet.add_chart(pie, "E16")

    # PIE 3

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=5, max_row=6)
    data = Reference(all_sheet, min_col=2, min_row=5, max_row=6)
    pie.add_data(data)
    pie.set_categories(labels)
    title = all_sheet['B6'].value
    num = 100 - int((100 * all_sheet['B5'].value )/ (all_sheet['B5'].value + all_sheet['B4'].value))
    pie.title = str(title) + " (" + str(num) + "%)" + " POSTES CRITIQUES NON CONFORMES"
    all_sheet.add_chart(pie, "E31")

    # PIE 4

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=7, max_row=8)
    data = Reference(all_sheet, min_col=2, min_row=7, max_row=8)
    pie.add_data(data)
    pie.set_categories(labels)
    pie.title = "Rapport postes bureautique et critique en échec"
    all_sheet.add_chart(pie, "E46")

    # PIE 5

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=14, max_row=23)
    data = Reference(all_sheet, min_col=2, min_row=14, max_row=23)
    pie.add_data(data)
    pie.set_categories(labels)
    pie.title = "Rapport taux d'échec par région"
    all_sheet.add_chart(pie, "Y46")

    # PIE 6

    pie = PieChart()
    labels = Reference(all_sheet, min_col=1, min_row=47, max_row=51)
    data = Reference(all_sheet, min_col=2, min_row=47, max_row=51)
    pie.add_data(data)
    pie.set_categories(labels)
    pie.title = "Rapport typologie de non conformité"
    all_sheet.add_chart(pie, "O46")

    # BAR 1

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "POSTES EN BULLES"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'TYPE DE POSTES'

    data = Reference(all_sheet, min_col=2, min_row=9, max_row=12, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=10, max_row=12)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    all_sheet.add_chart(chart1, "AI1")

    # BAR 2

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "POSTES EN ECHEC PAR REGIONS"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'LOCATIONS'

    data = Reference(all_sheet, min_col=2, min_row=13, max_row=23, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=14, max_row=23)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "Y1")

    # BAR 3

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 7
    chart1.title = "POSTES BUREAUTIQUE EN ECHEC PAR REGIONS"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'LOCATIONS'

    data = Reference(all_sheet, min_col=2, min_row=24, max_row=34, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=25, max_row=34)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "Y16")

    # BAR 4

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 7
    chart1.title = "POSTES CRITIQUES EN ECHEC PAR REGIONS"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'LOCATIONS'

    data = Reference(all_sheet, min_col=2, min_row=35, max_row=45, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=36, max_row=45)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "Y31")

    # BAR 5

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "TYPOLOGIE DE NON CONFORMITE"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'TYPE'

    data = Reference(all_sheet, min_col=2, min_row=46, max_row=51, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=47, max_row=51)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "O1")

    # BAR 6

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "TYPOLOGIE DE NON CONFORMITE :\n POSTES BUREAUTIQUES"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'TYPE'

    data = Reference(all_sheet, min_col=2, min_row=52, max_row=57, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=53, max_row=57)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "O16")

    # BAR 7

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "TYPOLOGIE DE NON CONFORMITE :\n POSTES CRITIQUES"
    chart1.y_axis.title = 'NOMBRE DE POSTES'
    chart1.x_axis.title = 'TYPE'

    data = Reference(all_sheet, min_col=2, min_row=58, max_row=63, max_col=2)
    cats = Reference(all_sheet, min_col=1, min_row=59, max_row=63)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 1
    all_sheet.add_chart(chart1, "O31")

    rapport.save("../conformity-report/" + file_name + ".xlsx")


init_graph()
write_graph()
