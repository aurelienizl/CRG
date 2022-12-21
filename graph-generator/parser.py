from datetime import date

from openpyxl import *

today = date.today()
file_name = today.strftime("%b-%d-%Y")

def init_graph():
    rapport = load_workbook("../conformity-report/" + file_name + ".xlsx")
    all_sheet = rapport.worksheets[0]
    issues_sheet = rapport.worksheets[1]
    graph_sheet = rapport.create_sheet()
    graph_sheet.title = "GRAPHIQUES"

    total = -1
    default = 0

    av = 0
    sccm = 0
    update = 0
    wsus = 0
    ou = 0

    avb = 0
    sccmb = 0
    updateb = 0
    wsusb = 0
    oub = 0

    avc = 0
    sccmc = 0
    updatec = 0
    wsusc = 0
    ouc = 0

    total_critique = 0
    total_workstation = 0

    critique = 0
    workstation = 0

    isolatedWindows10 = 0
    isolatedWindows7 = 0
    isolatedRemediation = 0

    #paris 0 , guade 1, vouvellec 2, polynes 3, reuni 4, walli 5, spm 6, guy 7, marti 8, mayotte 9

    default_agency = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    critique_agency = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    workstation_agency = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    for el in all_sheet:
        if el[0].value != 'NAME':
            if el[8].value == "CRITIQUE":
                total_critique += 1
            elif el[8].value == "BUREAUTIQUE":
                total_workstation += 1

    for el in all_sheet:
        total += 1
        if "POSTE" in el[10].value:
            if "10" in el[9].value:
                isolatedWindows10 += 1
            elif "7" in el[9].value:
                isolatedWindows7 += 1
        if "EN" in el[11].value:
            isolatedRemediation += 1
    for el in issues_sheet:
        if el[0].value != 'NAME':
            default += 1
            if el[8].value == "CRITIQUE":
                critique += 1
            elif el[8].value == "BUREAUTIQUE":
                workstation += 1

    for el in issues_sheet:
        if el[3].value[0] == "W":
            av += 1
            if "BUR" in el[8].value:
                avb += 1
            else:
                avc += 1
        if el[4].value[1] == "A":
            wsus += 1
            if "BUR" in el[8].value:
                wsusb += 1
            else:
                wsusc += 1
        if el[5].value[0] == "W":
            sccm += 1
            if "BUR" in el[8].value:
                sccmb += 1
            else:
                sccmc += 1
        if el[6].value[0] == 'W':
            ou += 1
            if "BUR" in el[8].value:
                oub += 1
            else:
                ouc += 1
        if el[7].value[0] == "W":
            update += 1
            if "BUR" in el[8].value:
                updateb += 1
            else:
                updatec += 1

    for el in issues_sheet:
        if "SPM" in el[2].value:
            default_agency[6] += 1
            if "CRI" in el[8].value:
                critique_agency[6] += 1
            else:
                workstation_agency[6] += 1
        elif "Paris" in el[2].value:
            default_agency[0] += 1
            if "CRI" in el[8].value:
                critique_agency[0] += 1
            else:
                workstation_agency[0] += 1
        elif "Reunion" in el[2].value:
            default_agency[4] += 1
            if "CRI" in el[8].value:
                critique_agency[4] += 1
            else:
                workstation_agency[4] += 1
        elif "Nouvelle" in el[2].value:
            default_agency[2] += 1
            if "CRI" in el[8].value:
                critique_agency[2] += 1
            else:
                workstation_agency[2] += 1
        elif "Guadelo" in el[2].value:
            default_agency[1] += 1
            if "CRI" in el[8].value:
                critique_agency[1] += 1
            else:
                workstation_agency[1] += 1
        elif "Mayotte" in el[2].value:
            default_agency[9] += 1
            if "CRI" in el[8].value:
                critique_agency[9] += 1
            else:
                workstation_agency[9] += 1
        elif "Polyn" in el[2].value:
            default_agency[3] += 1
            if "CRI" in el[8].value:
                critique_agency[3] += 1
            else:
                workstation_agency[3] += 1
        elif "Guyan" in el[2].value:
            default_agency[7] += 1
            if "CRI" in el[8].value:
                critique_agency[7] += 1
            else:
                workstation_agency[7] += 1
        elif "Martini" in el[2].value:
            default_agency[8] += 1
            if "CRI" in el[8].value:
                critique_agency[8] += 1
            else:
                workstation_agency[8] += 1
        elif "Walli" in el[2].value:
            default_agency[5] += 1
            if "CRI" in el[8].value:
                critique_agency[5] += 1
            else:
                workstation_agency[5] += 1

    graph_sheet.append(["NOMBRE DE POSTES : " + str(total), total - default])
    graph_sheet.append(["NOMBRE DE POSTES EN ECHEC : " + str(default), default])

    graph_sheet.append(["NOMBRE DE POSTES BUREAUTIQUE : " + str(total_workstation), total_workstation])
    graph_sheet.append(["NOMBRE DE POSTES BUREAUTIQUE EN ECHEC : " + str(workstation), workstation])

    graph_sheet.append(["NOMBRE DE POSTES CRITIQUES: " + str(total_critique), total_critique])
    graph_sheet.append(["NOMBRE DE POSTES CRITIQUES EN ECHEC : " + str(critique), critique])

    graph_sheet.append(["NOMBRE DE POSTES BUREAUTIQUE EN ECHEC : " + str(workstation), workstation])
    graph_sheet.append(["NOMBRE DE POSTES CRITIQUES EN ECHEC : " + str(critique), critique])

    graph_sheet.append(["TYPE", "POSTES ISOLES"])

    graph_sheet.append(["Win 10 : " + str(isolatedWindows10), isolatedWindows10])
    graph_sheet.append(["Win 7 : " + str(isolatedWindows7) , isolatedWindows7])
    graph_sheet.append(["REMEDIATION : " + str(isolatedRemediation), isolatedRemediation])

    # paris 0 , guade 1, vouvellec 2, polynes 3, reuni 4, walli 5, spm 6, guy 7, marti 8, mayotte 9

    graph_sheet.append(["", "POSTES EN ECHEC"])
    graph_sheet.append(["Paris : " + str(default_agency[0]), default_agency[0]])
    graph_sheet.append(["Guadeloupe : " + str(default_agency[1]), default_agency[1]])
    graph_sheet.append(["Nouvelle Calédonie : " + str(default_agency[2]), default_agency[2]])
    graph_sheet.append(["Polynésie : " + str(default_agency[3]), default_agency[3]])
    graph_sheet.append(["Réunion : " + str(default_agency[4]), default_agency[4]])
    graph_sheet.append(["Wallis : " + str(default_agency[5]), default_agency[5]])
    graph_sheet.append(["SPM : " + str(default_agency[6]), default_agency[6]])
    graph_sheet.append(["Guyane : " + str(default_agency[7]), default_agency[7]])
    graph_sheet.append(["Martinique : " + str(default_agency[8]), default_agency[8]])
    graph_sheet.append(["Mayotte : " + str(default_agency[9]), default_agency[9]])

    graph_sheet.append(["", "POSTES EN ECHEC"])
    graph_sheet.append(["Paris : " + str(workstation_agency[0]), workstation_agency[0]])
    graph_sheet.append(["Guadeloupe : " + str(workstation_agency[1]), workstation_agency[1]])
    graph_sheet.append(["Nouvelle Calédonie : " + str(workstation_agency[2]), workstation_agency[2]])
    graph_sheet.append(["Polynésie : " + str(workstation_agency[3]), workstation_agency[3]])
    graph_sheet.append(["Réunion : " + str(workstation_agency[4]), workstation_agency[4]])
    graph_sheet.append(["Wallis : " + str(workstation_agency[5]), workstation_agency[5]])
    graph_sheet.append(["SPM : " + str(workstation_agency[6]), workstation_agency[6]])
    graph_sheet.append(["Guyane : " + str(workstation_agency[7]), workstation_agency[7]])
    graph_sheet.append(["Martinique : " + str(workstation_agency[8]), workstation_agency[8]])
    graph_sheet.append(["Mayotte : " + str(workstation_agency[9]), workstation_agency[9]])

    graph_sheet.append(["", "POSTES EN ECHEC"])
    graph_sheet.append(["Paris : " + str(critique_agency[0]), critique_agency[0]])
    graph_sheet.append(["Guadeloupe : " + str(critique_agency[1]), critique_agency[1]])
    graph_sheet.append(["Nouvelle Calédonie : " + str(critique_agency[2]), critique_agency[2]])
    graph_sheet.append(["Polynésie : " + str(critique_agency[3]), critique_agency[3]])
    graph_sheet.append(["Réunion : " + str(critique_agency[4]), critique_agency[4]])
    graph_sheet.append(["Wallis : " + str(critique_agency[5]), critique_agency[5]])
    graph_sheet.append(["SPM : " + str(critique_agency[6]), critique_agency[6]])
    graph_sheet.append(["Guyane : " + str(critique_agency[7]), critique_agency[7]])
    graph_sheet.append(["Martinique : " + str(critique_agency[8]), critique_agency[8]])
    graph_sheet.append(["Mayotte : " + str(critique_agency[9]), critique_agency[9]])

    graph_sheet.append(["TYPOLOGIE NON CONFORMITE", "TYPE"])
    graph_sheet.append(["ANTIVIRUS : " + str(av), av])
    graph_sheet.append(["WSUS : " + str(wsus), wsus])
    graph_sheet.append(["SCCM : " + str(sccm), sccm])
    graph_sheet.append(["OU : " + str(ou), ou])
    graph_sheet.append(["UPDATE : " + str(update), update])

    graph_sheet.append(["TYPOLOGIE NON CONFORMITE", "TYPE"])
    graph_sheet.append(["ANTIVIRUS : " + str(avb), avb])
    graph_sheet.append(["WSUS : " + str(wsusb), wsusb])
    graph_sheet.append(["SCCM : " + str(sccmb), sccmb])
    graph_sheet.append(["OU : " + str(oub), oub])
    graph_sheet.append(["UPDATE : " + str(updateb), updateb])

    graph_sheet.append(["TYPOLOGIE NON CONFORMITE", "TYPE"])
    graph_sheet.append(["ANTIVIRUS : " + str(avc), avc])
    graph_sheet.append(["WSUS : " + str(wsusc), wsusc])
    graph_sheet.append(["SCCM : " + str(sccmc), sccmc])
    graph_sheet.append(["OU : " + str(ouc), ouc])
    graph_sheet.append(["UPDATE : " + str(updatec), updatec])

    rapport.save("../conformity-report/" + file_name + ".xlsx")


