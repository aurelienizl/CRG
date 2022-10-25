import datetime
from parser import *

WSUS_STATUS = []
BITDEFENDER_STATUS = []
PC_CRITIQUE = []
LANSWEEPER_STATUS = []
SCCM_STATUS = []
OU_STATUS = []
UPDATE_STATUS = []
REMEDIATION = []
BULLE = []


def lansweeper_status():
    for el in LANSWEEPER:
        if el[2] + datetime.timedelta(days=14) > datetime.datetime.now():
            LANSWEEPER_STATUS.append(el)


def pc_critique_status():
    for pc in PARC_NAME:
        if pc in FIDUCIA or pc in GIPOM or pc in QUALIAC or "XPERT" in pc:
            if pc not in PC_CRITIQUE:
                PC_CRITIQUE.append(pc)


def wsus_status():
    for pc in WSUS:
        if float(pc[1]) == 0:
            WSUS_STATUS.append(pc[0])


def sccm_status():
    for pc in LANSWEEPER_STATUS:
        if pc[0] not in SCCM:
            SCCM_STATUS.append(pc[0])


def ou_status():
    for el in OU:
        if "REMEDIATION" in el[1]:
            REMEDIATION.append(el[0])
        elif "WINDOWS 10" in el[1]:
            pass
        elif "BULLE" in el[1]:
            BULLE.append(el[0])
        elif "AGENT DE" in el[1]:
            pass
    for el in LANSWEEPER_STATUS:
        if el[1] == "WIN 7":
            BULLE.append(el[0])


def update_status():
    for el in UPDATE:
        if not (el[1] == "NOT APPLICABLE" or el[1] == "INSTALLED" or el[1] == "STATUS"):
            UPDATE_STATUS.append(el[0])
        else:
            pass
    UPDATE_LIST = []
    for el in UPDATE:
        UPDATE_LIST.append(el[0])
    for el in PARC_NAME:
        if el not in UPDATE_LIST:
            UPDATE_STATUS.append(el)


def bitdefender_status():
    versions = []
    for el in BITDEFENDER:
        versions.append(el[1])
    versions.remove("Version du Produit")
    versions = list(set(versions))
    versions.sort()
    versions.reverse()
    for i in range(2, len(versions)):
        versions.pop()
    for el in BITDEFENDER:
        if el[1] not in versions:
            if "NOM DE" not in el[0]:
                BITDEFENDER_STATUS.append(el[0])
    BITDEFENDER_NAME = []
    for el in BITDEFENDER:
        BITDEFENDER_NAME.append(el[0])
    for pc in PARC_NAME:
        if pc not in BITDEFENDER_NAME:
            BITDEFENDER_STATUS.append(pc)



def status():
    pc_critique_status()
    wsus_status()
    lansweeper_status()
    ou_status()
    bitdefender_status()
    sccm_status()
    update_status()

