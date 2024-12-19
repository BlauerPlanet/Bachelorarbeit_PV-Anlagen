import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
from os import listdir, path
from PIL import Image as img
import math
import numpy as np
from scipy import stats as st
import statsmodels.stats.multitest as mt

param = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]","mL","mF","mR","mN","mM","mW","mTr", "AntThero [%]", "n_Thero", "AntHemi [%]", "n_Hemi","n_Geo","n_Cha","n_Pha","E [%]","Hs","AbS","Artzahl", "Bruch_keinGras_Gras"] #,"mT"
vegparam = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]"]
staparam = ["mL","mF","mR","mN"]#"mT",
divparam=["AntThero [%]", "n_Thero", "AntHemi [%]", "n_Hemi","n_Geo","n_Cha","n_Pha","E [%]","Hs","AbS","Artzahl","Bruch_keinGras_Gras"]
nutzparam=["mM","mW","mTr"]
#paramDict = {"Dges [%]":"Gesamtdeckung (Dges) in %","DKr [%]":"Deckung der Krautschicht (DKr) in %","DMoos [%]":"Deckung der Moosschicht (DMoos) in %","DStreu [%]":"Deckung der Streuschicht (DStreu) in %","HKr [cm]":"Höhe der Krautschicht (HKr) in cm","mT":"mittlere Temperaturzahl (mT)","mL":"mittlere Lichtzahl (mL)","mF":"mittlere Feuchtezahl (mF)","mR":"mittlere Reaktionszahl (mR)","mN":"mittlere Nährstoffzahl (mN)","mM":"mittlere Mahdverträglichkeitszahl (mM)","mW":"mittlere Weideverträglichkeitszahl (mW)","mTr":"mittlere Trittverträglichkeitszahl (mTr)", "AntThero [%]":"Anteil der Therophyten an der Gesamtartenzahl (AntThero) in %", "n_Thero":"n_Thero", "AntHemi [%]":"Anteil der Hemikryptophyten an der Gesamtartenzahl (AntHemi) in %", "n_Hemi":"n_Hemi","n_Geo":"n_Geo","n_Cha":"n_Cha","n_Pha":"n_Pha","E [%]":"Pielou-Evenness (E) in %","Hs":"Shannon-Index (Hs)","AbS":"Abundanzsumme in %","Artzahl":"Gesamtartenzahl", "Bruch_keinGras_Gras":"Verhältnis der Nicht-Grasarten zu den Grasarten"}
# Anpassung Länge wegen y-Achsengröße
paramDict = {"Dges [%]":"Gesamtdeckung (Dges) in %","DKr [%]":"Deckung der Krautschicht (DKr) in %","DMoos [%]":"Deckung der Moosschicht (DMoos) in %","DStreu [%]":"Deckung der Streuschicht (DStreu) in %","HKr [cm]":"Höhe der Krautschicht (HKr) in cm","mT":"mittlere Temperaturzahl (mT)","mL":"mittlere Lichtzahl (mL)","mF":"mittlere Feuchtezahl (mF)","mR":"mittlere Reaktionszahl (mR)","mN":"mittlere Nährstoffzahl (mN)","mM":"mittlere Mahdverträglichkeitszahl (mM)","mW":"mittlere Weideverträglichkeitszahl (mW)","mTr":"mittlere Trittverträglichkeitszahl (mTr)", "AntThero [%]":"Anteil der Therophyten an der\n Gesamtartenzahl (AntThero) in %", "n_Thero":"n_Thero", "AntHemi [%]":"Anteil der Hemikryptophyten an der\n Gesamtartenzahl (AntHemi) in %", "n_Hemi":"n_Hemi","n_Geo":"n_Geo","n_Cha":"n_Cha","n_Pha":"n_Pha","E [%]":"Pielou-Evenness (E) in %","Hs":"Shannon-Index (Hs)","AbS":"Abundanzsumme in %","Artzahl":"Gesamtartenzahl", "Bruch_keinGras_Gras":"Verhältnis der Nicht-Grasarten zu den Grasarten"}
importantParam = ["Artzahl", "Bruch_keinGras_Gras", "mF","DMoos [%]","DStreu [%]","DKr [%]"]
paramKorr = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]","mF","mR","mN","mM","mW", "AntThero [%]", "AntHemi [%]", "Artzahl", "Bruch_keinGras_Gras"]

def Kopfdaten(selection):
    Stat = 0
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
               
    for i in param:
        KP=pd.DataFrame(K.loc[f"{i}"])
#        print(KP)
           
        if selection == "DS":
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indM in DSMList:
#                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indM}"]))
            for indH in DSHList:
                DSH.append(float(KP.loc[f"{indH}"]))
            for indS in DSSList:
                DSS1.append(float(KP.loc[f"{indS}"]))
            
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)
            #print(f"{i} - {DSS}")
            KPM=pd.DataFrame(DSM)
            KPH=pd.DataFrame(DSH)
            KPS=pd.DataFrame(DSS)
            pd.DataFrame(data=[DSM,DSH,DSS], index=["M","H","S"]).to_excel(f"XLSX\Korr\{selection}\{i}.xlsx")

            #KWlist=[]
            #KWlist.append(st.kruskal(KPM, KPH)[1][0])
            #KWlist.append(st.kruskal(KPM, KPS)[1][0])
            #KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)
            # np.std | ddof changes divisor -> ddof=1 -> divosr=N-1 => as is statgraphics
            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0]] # st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0]] # ,"","", "","",""
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0]] # ,"","", "","",""
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std"]).to_excel(f"XLSX\{selection}\{i}.xlsx")#,"KW","p","M-H","M-S","H-S"
            
            #print(f"{i} - {st.kruskal(DSM,DSH,DSS)} - [M-H,M-S,H-S] {mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0]}")
            #print(Mdata)

            #print(KPS)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True,medianprops={"color":"blue"},meanprops={"marker":"+"})
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,3,3,sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()
        
        elif selection == "TH":
            THM=[]
            THH=[]
            THS=[]
            THMList=["THM","THM.1","THM.2","THM.3","THM.4"]
            THHList=["THH","THH.1","THH.2","THH.3","THH.4"] 
            THSList=["THS","THS.1","THS.2","THS.3","THS.4"]
            for indM in THMList:
#                print(KP.loc[f"{ind}"])
                THM.append(float(KP.loc[f"{indM}"]))
            for indH in THHList:
                THH.append(float(KP.loc[f"{indH}"]))
            for indS in THSList:
                THS.append(float(KP.loc[f"{indS}"]))

            KPM=pd.DataFrame(THM)
            KPH=pd.DataFrame(THH)
            KPS=pd.DataFrame(THS)
            
            pd.DataFrame(data=[THM,THH,THS], index=["M","H","S"]).to_excel(f"XLSX\Korr\{selection}\{i}.xlsx")
            
            # KWlist=[]
            # KWlist.append(st.kruskal(KPM, KPH)[1][0])
            # KWlist.append(st.kruskal(KPM, KPS)[1][0])
            # KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)

            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0]] # ,st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0]] # ,"","", "","",""
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0]] # ,"","", "","",""
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std"]).to_excel(f"XLSX\{selection}\{i}.xlsx") # ,"KW","p","M-H","M-S","H-S"

            #print(f"{i} - {st.kruskal(THM,THH,THS)}")

#            print(KPM)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()    

        elif selection == "ZS":
            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indM in ZSMList:
#                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indM}"]))
            for indH in ZSHList:
                ZSH.append(float(KP.loc[f"{indH}"]))
            for indS in ZSSList:
                ZSS.append(float(KP.loc[f"{indS}"]))

            KPM=pd.DataFrame(ZSM)
            KPH=pd.DataFrame(ZSH)
            KPS=pd.DataFrame(ZSS)
            #print(f"{i} - {st.kruskal(ZSM,ZSH,ZSS)}")
            pd.DataFrame(data=[ZSM,ZSH,ZSS], index=["M","H","S"]).to_excel(f"XLSX\Korr\{selection}\{i}.xlsx")
            # KWlist=[]
            # KWlist.append(st.kruskal(KPM, KPH)[1][0])
            # KWlist.append(st.kruskal(KPM, KPS)[1][0])
            # KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)

            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0]] #,st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0]] # ,"","", "","",""
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0]] # ,"","", "","",""
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std"]).to_excel(f"XLSX\{selection}\{i}.xlsx") #,"KW","p","M-H","M-S","H-S"]

            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()   
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()   

        elif selection == "Anlage":
            DS=[]
            DS1=[]
            TH=[]
            ZS=[]
            DSList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6', 'DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6', 'DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6']
            THList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4', 'THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4', 'THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4']
            ZSList=['ZSM', 'ZSM.1', 'ZSM.2', 'ZSM.3', 'ZSH', 'ZSH.1', 'ZSH.2', 'ZSH.3', 'ZSS', 'ZSS.1', 'ZSS.2', 'ZSS.3']
            for indM in DSList:
#                print(KP.loc[f"{ind}"])
                DS1.append(float(KP.loc[f"{indM}"]))
            for indH in THList:
                TH.append(float(KP.loc[f"{indH}"]))
            for indS in ZSList:
                ZS.append(float(KP.loc[f"{indS}"]))

            for iK in DS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DS.append(iK)

            KPDS=pd.DataFrame(DS)#, index=list(range(11,32)))
            KPTH=pd.DataFrame(TH)#, index=list(range(11,26)))
            KPZS=pd.DataFrame(ZS)#, index=list(range(11,23)))
#            print(KPTH)
            #print(f"{i} - {st.kruskal(DS,TH,ZS)}")
            
            # KWlist=[]
            # KWlist.append(st.kruskal(KPDS, KPTH)[1][0])
            # KWlist.append(st.kruskal(KPDS, KPZS)[1][0])
            # KWlist.append(st.kruskal(KPTH, KPZS)[1][0])

            Mdata=[i,len(KPDS),np.mean(KPDS)[0],np.median(KPDS),np.std(KPDS, ddof=1)[0]] # ,st.kruskal(KPDS,KPTH,KPZS)[0],st.kruskal(KPDS,KPTH,KPZS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]
            Hdata=[i,len(KPTH),np.mean(KPTH)[0],np.median(KPTH),np.std(KPTH, ddof=1)[0]] # ,"","", "","",""
            Sdata=[i,len(KPZS),np.mean(KPZS)[0],np.median(KPZS),np.std(KPZS, ddof=1)[0]] # ,"","", "","",""

            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["DS","TH","ZS"], columns=["Parameter","n","mean","median","std"]).to_excel(f"XLSX\{selection}\{i}.xlsx") # ,"KW","p","DS-TH","DS-ZS","TH-ZS"
            
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - M, H & S ", loc="left")
            plt.boxplot(KPDS, labels="D", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPTH, labels="T", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPZS, labels="Z", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()

        elif selection == "AnlageFTyp":
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            THM=[]
            THH=[]
            THS=[]
            ZSM=[]
            ZSH=[]
            ZSS=[]
            DSMList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6']
            DSHList=['DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6']
            DSSList=['DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6']
            THMList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4']
            THHList=['THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4']
            THSList=['THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4']
            ZSMList=['ZSM', 'ZSM.1', 'ZSM.2', 'ZSM.3']
            ZSHList=['ZSH', 'ZSH.1', 'ZSH.2', 'ZSH.3']
            ZSSList=['ZSS', 'ZSS.1', 'ZSS.2', 'ZSS.3']

            for indDSM in DSMList:
#                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indDSM}"]))
            for indDSH in DSHList:
                DSH.append(float(KP.loc[f"{indDSH}"]))
            for indDSS in DSSList:
                DSS1.append(float(KP.loc[f"{indDSS}"]))                
            for indTHM in THMList:
                THM.append(float(KP.loc[f"{indTHM}"]))
            for indTHH in THHList:
                THH.append(float(KP.loc[f"{indTHH}"]))
            for indTHS in THSList:
                THS.append(float(KP.loc[f"{indTHS}"])) 
            for indZSM in ZSMList:
                ZSM.append(float(KP.loc[f"{indZSM}"]))
            for indZSH in ZSHList:
                ZSH.append(float(KP.loc[f"{indZSH}"]))
            for indZSS in ZSSList:
                ZSS.append(float(KP.loc[f"{indZSS}"])) 
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)

            KPDSM=pd.DataFrame(DSM)
            KPDSH=pd.DataFrame(DSH)
            KPDSS=pd.DataFrame(DSS)
            KPTHM=pd.DataFrame(THM)
            KPTHH=pd.DataFrame(THH)
            KPTHS=pd.DataFrame(THS)
            KPZSM=pd.DataFrame(ZSM)
            KPZSH=pd.DataFrame(ZSH)
            KPZSS=pd.DataFrame(ZSS)

            pd.DataFrame(data=[THM,ZSM,DSM], index=["TH","ZS","DS"]).to_excel(f"XLSX\Korr\M\{i}.xlsx")
            pd.DataFrame(data=[THH,ZSH,DSH], index=["TH","ZS","DS"]).to_excel(f"XLSX\Korr\H\{i}.xlsx")
            pd.DataFrame(data=[THS,ZSS,DSS], index=["TH","ZS","DS"]).to_excel(f"XLSX\Korr\S\{i}.xlsx")

        elif selection == "Ende":
            Stat=2
            return Stat
    Stat=1
    return Stat

def Shinozaki():
    dictAnlage = {"TH": "Tannenhübel", "DS": "Davidschacht", "ZS": "Ziegelscheune 2", "Alle Anlagen":"alle Anlagen"}
    ShinoSel=str(input("M+H - j/n? "))
    S=pd.read_excel("SORTstatistics.xlsx", sheet_name="Shinozaki - Anlage", header=2, index_col=0)
    Anlagen=["TH","ZS","DS","Alle Anlagen"]

#    print(S)
    plt.xticks(S.index, labels=[1,"",3,"",5,"",7,"",9,"",11,"",13,"",15,"",17,"",19,"",21])
    plt.plot(S["DS - S"], label="DS", marker="o")
    plt.plot(S["TH - S"], label="TH", marker="s")
    plt.plot(S["ZS - S"], label="ZS", marker="X")
    plt.xlabel("Anzahl Untersuchungsflächen (n)")
    plt.ylabel("Artzahl (n)")
    plt.legend()
    plt.title("Shinozaki-Kurve")
    plt.savefig(rf"Figures\Shinozaki.png", bbox_inches="tight")
#    plt.show()
    plt.close()

    for i in Anlagen:
        SZ=pd.read_excel("ShinozakiMHS.xlsx", sheet_name=i, header=1, index_col=0)
        plt.xticks(SZ.index, labels=SZ.index.values)
        plt.xlabel("Anzahl Untersuchungsflächen (n)")
        plt.ylabel("Artzahl (n)")
        plt.plot(SZ["M"], label="M", marker="o")
        plt.plot(SZ["H"], label="H", marker="s")
        if i!="Alle Anlagen" and ShinoSel=="j":
            plt.plot(SZ["M+H"], label="M+H", marker="X")
        plt.plot(SZ["S"], label="S", marker="D")
        plt.legend()
        # Bachelorarbeit
        #plt.title(f"Shinozaki-Kurve | {i}")
        # FECO
        plt.title(f"{dictAnlage[i]}")
        #plt.show()
        # Bachelorarbeit
        plt.savefig(rf"Figures\Shinozaki-{i}.png", bbox_inches="tight")
        # FECO
        plt.savefig(rf"D:\Daten_Noah\Desktop_Dateien\Studium_Freiberg\Study\FECO\finalPlots\Shinozaki-{i}.png", bbox_inches="tight")
        plt.close()

    Stat=2
    return Stat

def compareFTyp():
    Stat = 0
    w=0.5
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
    for i in param:
        KP=pd.DataFrame(K.loc[f"{i}"])
#        print(KP)
        print(i)

        def three():  
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indMDS in DSMList:
    #                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indMDS}"]))
            for indHDS in DSHList:
                DSH.append(float(KP.loc[f"{indHDS}"]))
            for indSDS in DSSList:
                DSS1.append(float(KP.loc[f"{indSDS}"]))
            
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)
            
            THM=[]
            THH=[]
            THS=[]
            THMList=["THM","THM.1","THM.2","THM.3","THM.4"]
            THHList=["THH","THH.1","THH.2","THH.3","THH.4"] 
            THSList=["THS","THS.1","THS.2","THS.3","THS.4"]
            for indMTH in THMList:
    #                print(KP.loc[f"{ind}"])
                THM.append(float(KP.loc[f"{indMTH}"]))
            for indHTH in THHList:
                THH.append(float(KP.loc[f"{indHTH}"]))
            for indSTH in THSList:
                THS.append(float(KP.loc[f"{indSTH}"]))

            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indMZS in ZSMList:
    #                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indMZS}"]))
            for indHZS in ZSHList:
                ZSH.append(float(KP.loc[f"{indHZS}"]))
            for indSZS in ZSSList:
                ZSS.append(float(KP.loc[f"{indSZS}"]))

            KPMDS=pd.DataFrame(DSM)
            KPHDS=pd.DataFrame(DSH)
            KPSDS=pd.DataFrame(DSS)
            KPMTH=pd.DataFrame(THM)
            KPHTH=pd.DataFrame(THH)
            KPSTH=pd.DataFrame(THS)
            KPMZS=pd.DataFrame(ZSM)
            KPHZS=pd.DataFrame(ZSH)
            KPSZS=pd.DataFrame(ZSS)
            if i == "n_Pha": #ValueError: All numbers are identical in kruskal H: 0,0,0 but in U-test it works with H: 0,0?!
                KPDSZSutest=[st.mannwhitneyu(KPMDS,KPMZS)[0],st.mannwhitneyu(KPMDS,KPMZS)[1],st.mannwhitneyu(KPHDS,KPHZS)[0],st.mannwhitneyu(KPHDS,KPHZS)[1],st.mannwhitneyu(KPSDS,KPSZS)[0],st.mannwhitneyu(KPSDS,KPSZS)[1]]
                pd.DataFrame(data=[KPDSZSutest], index=[i], columns=["U: MDS-MZS","p: MDS-MZS","U: HDS-HZS","p: HDS-HZS","U: SDS-SZS","p: SDS-SZS"]).to_excel(f"XLSX\DSZS\{i}.xlsx")
                KPDSTHZSkwtest=[st.kruskal(KPMDS,KPMTH,KPMZS)[0],st.kruskal(KPMDS,KPMTH,KPMZS)[1],"ValueError: All numbers are identical in kruskal","ValueError: All numbers are identical in kruskal",st.kruskal(KPSDS,KPSTH,KPSZS)[0],st.kruskal(KPSDS,KPSTH,KPSZS)[1]]
                #print(f"{i} - {KPMDS}")
                #print(f"{i} - {max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))}")
                pd.DataFrame(data=[KPDSTHZSkwtest], index=[i], columns=["KW: M-M","p: M-M","KW: H-H","p: H-H","KW: S-S","p: S-S"]).to_excel(f"XLSX\DSTHZS\{i}.xlsx")
                KPTHZSutest=[st.mannwhitneyu(KPMTH,KPMZS)[0],st.mannwhitneyu(KPMTH,KPMZS)[1],st.mannwhitneyu(KPHTH,KPHZS)[0],st.mannwhitneyu(KPHTH,KPHZS)[1],st.mannwhitneyu(KPSTH,KPSZS)[0],st.mannwhitneyu(KPSTH,KPSZS)[1]]
                KPTHDSutest=[st.mannwhitneyu(KPMTH,KPMDS)[0],st.mannwhitneyu(KPMTH,KPMDS)[1],st.mannwhitneyu(KPHTH,KPHDS)[0],st.mannwhitneyu(KPHTH,KPHDS)[1],st.mannwhitneyu(KPSTH,KPSDS)[0],st.mannwhitneyu(KPSTH,KPSDS)[1]]
                pd.DataFrame(data=[KPTHZSutest], index=[i], columns=["U: MTH-MZS","p: MTH-MZS","U: HTHS-HZS","p: HTH-HZS","U: STH-SZS","p: STH-SZS"]).to_excel(f"XLSX\THZS\{i}.xlsx")
                pd.DataFrame(data=[KPTHDSutest], index=[i], columns=["U: MTH-MSS","p: MTH-MDS","U: HTH-HDS","p: HTH-HDS","U: STH-SDS","p: STH-SDS"]).to_excel(f"XLSX\THDS\{i}.xlsx")
            else:
                KPDSZSutest=[st.mannwhitneyu(KPMDS,KPMZS)[0],st.mannwhitneyu(KPMDS,KPMZS)[1],st.mannwhitneyu(KPHDS,KPHZS)[0],st.mannwhitneyu(KPHDS,KPHZS)[1],st.mannwhitneyu(KPSDS,KPSZS)[0],st.mannwhitneyu(KPSDS,KPSZS)[1]]
                pd.DataFrame(data=[KPDSZSutest], index=[i], columns=["U: MDS-MZS","p: MDS-MZS","U: HDS-HZS","p: HDS-HZS","U: SDS-SZS","p: SDS-SZS"]).to_excel(f"XLSX\DSZS\{i}.xlsx")
                KPDSTHZSkwtest=[st.kruskal(KPMDS,KPMTH,KPMZS)[0],st.kruskal(KPMDS,KPMTH,KPMZS)[1],st.kruskal(KPHDS,KPHTH,KPHZS)[0],st.kruskal(KPHDS,KPHTH,KPHZS)[1],st.kruskal(KPSDS,KPSTH,KPSZS)[0],st.kruskal(KPSDS,KPSTH,KPSZS)[1]]
                #print(f"{i} - {KPMDS}")
                #print(f"{i} - {max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))}")
                pd.DataFrame(data=[KPDSTHZSkwtest], index=[i], columns=["KW: M-M","p: M-M","KW: H-H","p: H-H","KW: S-S","p: S-S"]).to_excel(f"XLSX\DSTHZS\{i}.xlsx")
                KPTHZSutest=[st.mannwhitneyu(KPMTH,KPMZS)[0],st.mannwhitneyu(KPMTH,KPMZS)[1],st.mannwhitneyu(KPHTH,KPHZS)[0],st.mannwhitneyu(KPHTH,KPHZS)[1],st.mannwhitneyu(KPSTH,KPSZS)[0],st.mannwhitneyu(KPSTH,KPSZS)[1]]
                KPTHDSutest=[st.mannwhitneyu(KPMTH,KPMDS)[0],st.mannwhitneyu(KPMTH,KPMDS)[1],st.mannwhitneyu(KPHTH,KPHDS)[0],st.mannwhitneyu(KPHTH,KPHDS)[1],st.mannwhitneyu(KPSTH,KPSDS)[0],st.mannwhitneyu(KPSTH,KPSDS)[1]]
                pd.DataFrame(data=[KPTHZSutest], index=[i], columns=["U: MTH-MZS","p: MTH-MZS","U: HTHS-HZS","p: HTH-HZS","U: STH-SZS","p: STH-SZS"]).to_excel(f"XLSX\THZS\{i}.xlsx")
                pd.DataFrame(data=[KPTHDSutest], index=[i], columns=["U: MTH-MSS","p: MTH-MDS","U: HTH-HDS","p: HTH-HDS","U: STH-SDS","p: STH-SDS"]).to_excel(f"XLSX\THDS\{i}.xlsx")

            yMAX=max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))
            yMIN=min(min(KPMDS[0]),min(KPHDS[0]),min(KPSDS[0]),min(KPMTH[0]),min(KPHTH[0]), min(KPSTH[0]),min(KPMZS[0]),min(KPHZS[0]),min(KPSZS[0]))
           
            if i in vegparam:
                yADD=15-(math.ceil(yMAX)%10) 
                ySUB=5+(yMIN%5)
            elif i in staparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            elif i in nutzparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            elif i in divparam:
                if i == "Hs":
                    yADD=1.2-(yMIN%1)
                    ySUB=0.2+(yMIN%0.2)
                elif i == "Bruch_keinGras_Gras":
                    yADD=0
                    ySUB=0
                    yMAX=10
                    yMIN=0
                else:
                    yADD=15-(yMIN%10)
                    ySUB=5+(yMIN%5)
            
            plt.figure()
            ax1=plt.subplot(1,11,1, ylim=(yMIN-ySUB,yMAX+yADD))
            if i in vegparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            elif i in staparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
            elif i in divparam:
                if i == "Hs":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
                elif i=="Bruch_keinGras_Gras":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.5))
                else:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            #plt.suptitle(f"{paramDict[i]}\n", ha="center", weight="bold")
            plt.ylabel(f"{paramDict[i]}", weight="bold")
            plt.title(f"TH", loc="left")
            plt.boxplot(KPMTH, labels="M", widths=w,showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,11,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHTH, labels="H", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,11,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSTH, labels="S", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax4=plt.subplot(1,11,5, sharey=ax1)
            plt.title("ZS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMZS, labels="M", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})  
            ax5=plt.subplot(1,11,6, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHZS, labels="H", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,11,7, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSZS, labels="S", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax7=plt.subplot(1,11,9, sharey=ax1)
            plt.title("DS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMDS, labels="M", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax8=plt.subplot(1,11,10, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHDS, labels="H", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax9=plt.subplot(1,11,11, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSDS, labels="S", widths=w, showmeans=True, medianprops={"color":"blue"},meanprops={"marker":"+"})
            if i=="Bruch_keinGras_Gras":
                ax1.get_yticklabels()[2].set_color("red")
                ax4.get_yticklabels()[2].set_color("red")
                ax7.get_yticklabels()[2].set_color("red")
                plt.text(-14.5, -1.5, "< 1: mehr Grasarten | = 1: Anzahl Nicht-Grasarten = Anzahl Grasarten | > 1: mehr Nicht-Grasarten")
            plt.savefig(rf"Figures\Compare\MHS\{i}.png", bbox_inches="tight")
            plt.close()  

        def two():
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indMDS in DSMList:
    #                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indMDS}"]))
            for indHDS in DSHList:
                DSH.append(float(KP.loc[f"{indHDS}"]))
            for indSDS in DSSList:
                DSS1.append(float(KP.loc[f"{indSDS}"]))
            
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)

            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indMZS in ZSMList:
    #                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indMZS}"]))
            for indHZS in ZSHList:
                ZSH.append(float(KP.loc[f"{indHZS}"]))
            for indSZS in ZSSList:
                ZSS.append(float(KP.loc[f"{indSZS}"]))

            KPMDS=pd.DataFrame(DSM)
            KPHDS=pd.DataFrame(DSH)
            KPSDS=pd.DataFrame(DSS)
            KPMZS=pd.DataFrame(ZSM)
            KPHZS=pd.DataFrame(ZSH)
            KPSZS=pd.DataFrame(ZSS)
            KPDSZSutest=[st.mannwhitneyu(KPMDS,KPMZS)[0],st.mannwhitneyu(KPMDS,KPMZS)[1],st.mannwhitneyu(KPHDS,KPHZS)[0],st.mannwhitneyu(KPHDS,KPHZS)[1],st.mannwhitneyu(KPSDS,KPSZS)[0],st.mannwhitneyu(KPSDS,KPSZS)[1]]
            pd.DataFrame(data=[KPDSZSutest], index=[i], columns=["U: MDS-MZS","p: MDS-MZS","U: HDS-HZS","p: HDS-HZS","U: SDS-SZS","p: SDS-SZS"]).to_excel(f"XLSX\DSZS\{i}.xlsx")        
            
            yMAX=max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))
            yMIN=min(min(KPMDS[0]),min(KPHDS[0]),min(KPSDS[0]),min(KPMZS[0]),min(KPHZS[0]),min(KPSZS[0]))

            if i in nutzparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            
            plt.figure()
            ax1=plt.subplot(1,7,1, ylim=(yMIN-ySUB,yMAX+yADD))
            if i in vegparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            elif i in staparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
            elif i in divparam:
                if i == "Hs":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
                else:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            
            #plt.suptitle(f"{paramDict[i]}\n", ha="center", weight="bold")
            plt.ylabel(f"{paramDict[i]}", weight="bold")
            plt.title(f" DS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMDS, labels="M", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,7,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHDS, labels="H", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,7,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSDS, labels="S", widths=w, showmeans=True, medianprops={"color":"blue"},meanprops={"marker":"+"})
            ax4=plt.subplot(1,7,5, sharey=ax1)
            plt.title("ZS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMZS, labels="M", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})  
            ax5=plt.subplot(1,7,6, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHZS, labels="H", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,7,7, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSZS, labels="S", widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            plt.savefig(rf"Figures\Compare\MHS\{i}.png", bbox_inches="tight")
            plt.close()  

        if i in staparam:
            three()
        elif i in vegparam:
            three()
        elif i in divparam:
            three()
        elif i in nutzparam:
            two()

    Stat=2
    return Stat

def compareAnlagen():
    Stat = 0
    w=0.5
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
    for i in param:
        KP=pd.DataFrame(K.loc[f"{i}"])
#        print(KP)
        print(i)

        def three():  
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indMDS in DSMList:
    #                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indMDS}"]))
            for indHDS in DSHList:
                DSH.append(float(KP.loc[f"{indHDS}"]))
            for indSDS in DSSList:
                DSS1.append(float(KP.loc[f"{indSDS}"]))
            
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)
            
            THM=[]
            THH=[]
            THS=[]
            THMList=["THM","THM.1","THM.2","THM.3","THM.4"]
            THHList=["THH","THH.1","THH.2","THH.3","THH.4"] 
            THSList=["THS","THS.1","THS.2","THS.3","THS.4"]
            for indMTH in THMList:
    #                print(KP.loc[f"{ind}"])
                THM.append(float(KP.loc[f"{indMTH}"]))
            for indHTH in THHList:
                THH.append(float(KP.loc[f"{indHTH}"]))
            for indSTH in THSList:
                THS.append(float(KP.loc[f"{indSTH}"]))

            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indMZS in ZSMList:
    #                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indMZS}"]))
            for indHZS in ZSHList:
                ZSH.append(float(KP.loc[f"{indHZS}"]))
            for indSZS in ZSSList:
                ZSS.append(float(KP.loc[f"{indSZS}"]))

            KPMDS=pd.DataFrame(DSM)
            KPHDS=pd.DataFrame(DSH)
            KPSDS=pd.DataFrame(DSS)
            KPMTH=pd.DataFrame(THM)
            KPHTH=pd.DataFrame(THH)
            KPSTH=pd.DataFrame(THS)
            KPMZS=pd.DataFrame(ZSM)
            KPHZS=pd.DataFrame(ZSH)
            KPSZS=pd.DataFrame(ZSS)
            
            yMAX=max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))
            yMIN=min(min(KPMDS[0]),min(KPHDS[0]),min(KPSDS[0]),min(KPMTH[0]),min(KPHTH[0]), min(KPSTH[0]),min(KPMZS[0]),min(KPHZS[0]),min(KPSZS[0]))
           
            if i in vegparam:
                yADD=15-(math.ceil(yMAX)%10) 
                ySUB=5+(yMIN%5)
            elif i in staparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            elif i in nutzparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            elif i in divparam:
                if i == "Hs":
                    yADD=1.2-(yMIN%1)
                    ySUB=0.2+(yMIN%0.2)
                elif i == "Bruch_keinGras_Gras":
                    yADD=0
                    ySUB=0
                    yMAX=10
                    yMIN=0
                else:
                    yADD=15-(yMIN%10)
                    ySUB=5+(yMIN%5)
            
            plt.figure()
            ax1=plt.subplot(1,11,1, ylim=(yMIN-ySUB,yMAX+yADD))
            if i in vegparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            elif i in staparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
            elif i in divparam:
                if i == "Hs":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
                elif i=="Bruch_keinGras_Gras":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.5))
                else:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            #plt.suptitle(f"{paramDict[i]}\n", ha="center", weight="bold")
            plt.ylabel(f"{paramDict[i]}", weight="bold")            
            plt.title(f"M", loc="left")
            plt.boxplot(KPMTH, labels=["TH"], widths=w,showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,11,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPMZS, labels=["ZS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,11,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPMDS, labels=["DS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax4=plt.subplot(1,11,5, sharey=ax1)
            plt.title("H", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPHTH, labels=["TH"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})  
            ax5=plt.subplot(1,11,6, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHZS, labels=["ZS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,11,7, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHDS, labels=["DS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax7=plt.subplot(1,11,9, sharey=ax1)
            plt.title("S", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPSTH, labels=["TH"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax8=plt.subplot(1,11,10, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSZS, labels=["ZS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax9=plt.subplot(1,11,11, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSDS, labels=["DS"], widths=w, showmeans=True, medianprops={"color":"blue"},meanprops={"marker":"+"})
            if i=="Bruch_keinGras_Gras":
                ax1.get_yticklabels()[2].set_color("red")
                ax4.get_yticklabels()[2].set_color("red")
                ax7.get_yticklabels()[2].set_color("red")
                plt.text(-14.5, -1.5, "< 1: mehr Grasarten | = 1: Anzahl Nicht-Grasarten = Anzahl Grasarten | > 1: mehr Nicht-Grasarten")
            plt.savefig(rf"Figures\Compare\THZSDS\{i}.png", bbox_inches="tight")
            plt.close()  

        def two():
            DSM=[]
            DSH=[]
            DSS=[]
            DSS1=[]
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indMDS in DSMList:
    #                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indMDS}"]))
            for indHDS in DSHList:
                DSH.append(float(KP.loc[f"{indHDS}"]))
            for indSDS in DSSList:
                DSS1.append(float(KP.loc[f"{indSDS}"]))
            
            for iK in DSS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    DSS.append(iK)

            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indMZS in ZSMList:
    #                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indMZS}"]))
            for indHZS in ZSHList:
                ZSH.append(float(KP.loc[f"{indHZS}"]))
            for indSZS in ZSSList:
                ZSS.append(float(KP.loc[f"{indSZS}"]))

            KPMDS=pd.DataFrame(DSM)
            KPHDS=pd.DataFrame(DSH)
            KPSDS=pd.DataFrame(DSS)
            KPMZS=pd.DataFrame(ZSM)
            KPHZS=pd.DataFrame(ZSH)
            KPSZS=pd.DataFrame(ZSS)
            
            yMAX=max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))
            yMIN=min(min(KPMDS[0]),min(KPHDS[0]),min(KPSDS[0]),min(KPMZS[0]),min(KPHZS[0]),min(KPSZS[0]))

            if i in nutzparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            
            plt.figure()
            ax1=plt.subplot(1,8,1, ylim=(yMIN-ySUB,yMAX+yADD))
            if i in vegparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            elif i in staparam:
                ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
            elif i in divparam:
                if i == "Hs":
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
                else:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            
            #plt.suptitle(f"{paramDict[i]}\n", ha="center", weight="bold")
            plt.ylabel(f"{paramDict[i]}", weight="bold")
            plt.title(f"M", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMZS, labels=["ZS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,8,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPMDS, labels=["DS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,8,4, sharey=ax1)
            plt.title("H", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPHZS, labels=["ZS"], widths=w, showmeans=True, medianprops={"color":"blue"},meanprops={"marker":"+"})
            ax4=plt.subplot(1,8,5, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHDS, labels=["DS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})  
            ax5=plt.subplot(1,8,7, sharey=ax1)
            plt.title("S", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPSZS, labels=["ZS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,8,8, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSDS, labels=["DS"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            plt.savefig(rf"Figures\Compare\THZSDS\{i}.png", bbox_inches="tight")
            plt.close()  

        if i in staparam:
            three()
        elif i in vegparam:
            three()
        elif i in divparam:
            three()
        elif i in nutzparam:
            two()

    Stat=2
    return Stat

def compareFECO():
    # inserted into code 24.10.24 & 25.10.24 & 06.11.24
    Stat = 0
    w=0.6
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)

    FECOparam = ["mW","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]","mL","mF","mN", "AntThero [%]", "AntHemi [%]", "Artzahl", "Bruch_keinGras_Gras"] 
    FECOvegparam = ["DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]"]
    FECOstaparam = ["mL","mF","mN"]
    FECOdivparam=["AntThero [%]", "AntHemi [%]","Artzahl","Bruch_keinGras_Gras"]
    paramDictFECO = {"DKr [%]":"Deckung der Krautschicht in %","DMoos [%]":"Deckung der Moosschicht in %","DStreu [%]":"Deckung der Streuschicht in %","HKr [cm]":"Höhe der Krautschicht in cm","mL":"mittlere Lichtzahl","mF":"mittlere Feuchtezahl","mN":"mittlere Nährstoffzahl","mW":"mittlere Weideverträglichkeitszahl","AntThero [%]":"Anteil der Therophyten an der\n Gesamtartenzahl in %", "AntHemi [%]":"Anteil der Hemikryptophyten an der\n Gesamtartenzahl in %","Artzahl":"Gesamtartenzahl", "Bruch_keinGras_Gras":"Verhältnis der Nicht-Grasarten\nzu den Grasarten"}
    
    font = {'family' : 'sans-serif',
        'weight' : 'regular',
        'size'   : 13}

    plt.rc('font', **font)
    
    for i in FECOparam:
        KP=pd.DataFrame(K.loc[f"{i}"])
#        print(KP)
        print(i)

        if i != "mW":

            def three():  
                DSM=[]
                DSH=[]
                DSS=[]
                DSS1=[]
                DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
                DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
                DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
                for indMDS in DSMList:
        #                print(KP.loc[f"{ind}"])
                    DSM.append(float(KP.loc[f"{indMDS}"]))
                for indHDS in DSHList:
                    DSH.append(float(KP.loc[f"{indHDS}"]))
                for indSDS in DSSList:
                    DSS1.append(float(KP.loc[f"{indSDS}"]))
                
                for iK in DSS1:
                    #print({str(iK).replace(".","",1).isdigit()})
                    if str(iK).replace(".","",1).isdigit():
                        DSS.append(iK)
                
                THM=[]
                THH=[]
                THS=[]
                THMList=["THM","THM.1","THM.2","THM.3","THM.4"]
                THHList=["THH","THH.1","THH.2","THH.3","THH.4"] 
                THSList=["THS","THS.1","THS.2","THS.3","THS.4"]
                for indMTH in THMList:
        #                print(KP.loc[f"{ind}"])
                    THM.append(float(KP.loc[f"{indMTH}"]))
                for indHTH in THHList:
                    THH.append(float(KP.loc[f"{indHTH}"]))
                for indSTH in THSList:
                    THS.append(float(KP.loc[f"{indSTH}"]))

                ZSM=[]
                ZSH=[]
                ZSS=[]
                ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
                ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
                ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
                for indMZS in ZSMList:
        #                print(KP.loc[f"{ind}"])
                    ZSM.append(float(KP.loc[f"{indMZS}"]))
                for indHZS in ZSHList:
                    ZSH.append(float(KP.loc[f"{indHZS}"]))
                for indSZS in ZSSList:
                    ZSS.append(float(KP.loc[f"{indSZS}"]))

                KPMDS=pd.DataFrame(DSM)
                KPHDS=pd.DataFrame(DSH)
                KPSDS=pd.DataFrame(DSS)
                KPMTH=pd.DataFrame(THM)
                KPHTH=pd.DataFrame(THH)
                KPSTH=pd.DataFrame(THS)
                KPMZS=pd.DataFrame(ZSM)
                KPHZS=pd.DataFrame(ZSH)
                KPSZS=pd.DataFrame(ZSS)
                
                fig, ax1=plt.subplots() #ylim=(yMIN-ySUB,yMAX+yADD)
                if i in FECOvegparam:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
                    ax1.set_ylim(bottom=-4, top=104)
                elif i in FECOstaparam:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
                    ax1.set_ylim(bottom=3, top=8)
                elif i in FECOdivparam:
                    if i=="Bruch_keinGras_Gras":
                        ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
                        ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.5))
                        ax1.set_ylim(bottom=-0.3, top=8.4)
                    elif i =="Artzahl":
                        ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                        ax1.yaxis.set_minor_locator(ticker.MultipleLocator(1))
                        ax1.set_ylim(bottom=5, top=23)
                    else:
                        ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                        ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
                        ax1.set_ylim(bottom=-4, top=104)
                ax1.set_ylabel(f"{paramDictFECO[i]}", weight="bold")
                #font-size:13
                ax1.set_title("          TH", loc="left")
                ax1.set_title("ZS", loc="center")
                ax1.set_title("DS          ", loc="right")
                #font-size: 10
                # ax1.set_title("             TH", loc="left")
                # ax1.set_title("ZS", loc="center")
                # ax1.set_title("DS             ", loc="right")
                ax1.boxplot([THM, THH, THS], labels=["M","H","S"], widths=w,showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"}, positions=[0,1,2])
                ax1.axvline(x=2.5, ls="--", c=(0.6,0.6,0.6))
                ax1.boxplot([ZSM,ZSH,ZSS], labels=["M","H","S"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"}, positions=[3,4,5])  
                ax1.axvline(x=5.5, ls="--", c=(0.6,0.6,0.6))
                ax1.boxplot([DSM,DSH,DSS], labels=["M","H","S"], widths=w, showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"}, positions=[6,7,8])
                if i=="Bruch_keinGras_Gras":
                    ax1.get_yticklabels()[2].set_color("red")
                    # font-size: 13
                    plt.text(-0.3, -2, "< 1: mehr Grasarten | > 1: mehr Nicht-Grasarten\n = 1: Anzahl Nicht-Grasarten = Anzahl Grasarten")
                    # font-size:10
                    # plt.text(-2.5, -1.5, "< 1: mehr Grasarten | = 1: Anzahl Nicht-Grasarten = Anzahl Grasarten | > 1: mehr Nicht-Grasarten")
                plt.savefig(rf"D:\Daten_Noah\Desktop_Dateien\Studium_Freiberg\Study\FECO\finalPlots\{i}.png", bbox_inches="tight")
                plt.close()  

            if i in FECOstaparam:
                three()
            elif i in FECOvegparam:
                three()
            elif i in FECOdivparam:
                three()
        elif i == "mW":
            THM=[]
            THH=[]
            THS=[]
            THMList=["THM","THM.1","THM.2","THM.3","THM.4"]
            THHList=["THH","THH.1","THH.2","THH.3","THH.4"] 
            THSList=["THS","THS.1","THS.2","THS.3","THS.4"]
            for indMTH in THMList:
    #                print(KP.loc[f"{ind}"])
                THM.append(float(KP.loc[f"{indMTH}"]))
            for indHTH in THHList:
                THH.append(float(KP.loc[f"{indHTH}"]))
            for indSTH in THSList:
                THS.append(float(KP.loc[f"{indSTH}"]))


            ZSM=[]
            ZSH=[]
            ZSS=[]
            ZSMList=["ZSM","ZSM.1","ZSM.2","ZSM.3"]
            ZSHList=["ZSH","ZSH.1","ZSH.2","ZSH.3"] 
            ZSSList=["ZSS","ZSS.1","ZSS.2","ZSS.3"]
            for indMZS in ZSMList:
    #                print(KP.loc[f"{ind}"])
                ZSM.append(float(KP.loc[f"{indMZS}"]))
            for indHZS in ZSHList:
                ZSH.append(float(KP.loc[f"{indHZS}"]))
            for indSZS in ZSSList:
                ZSS.append(float(KP.loc[f"{indSZS}"]))   
     
            fig, ax1=plt.subplots()
            ax1.set_ylim(bottom=3, top=8)
            ax1.yaxis.set_major_locator(ticker.MultipleLocator(1))
            ax1.yaxis.set_minor_locator(ticker.MultipleLocator(.2))
            ax1.set_ylabel(f"{paramDictFECO[i]}", weight="bold")
            #font-size: 13
            ax1.set_title("                TH", loc="left")
            ax1.set_title("ZS                ", loc="right")
            #font-size:10
            # ax1.set_title("                     TH", loc="left")
            # ax1.set_title("ZS                      ", loc="right")
            ax1.boxplot([THM, THH, THS], labels=["M","H","S"], widths=0.5,showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"}, positions=[0,1,2])
            ax1.axvline(x=2.5, ls="--", c=(0.6,0.6,0.6))
            ax1.boxplot([ZSM, ZSH, ZSS], labels=["M","H","S"], widths=0.5,showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"}, positions=[3,4,5])
            plt.savefig(rf"D:\Daten_Noah\Desktop_Dateien\Studium_Freiberg\Study\FECO\finalPlots\{i}.png")
            plt.close()  
    Stat=2
    return Stat


def merge():
    pathDS=r"XLSX\DS"
    pathTH=r"XLSX\TH"
    pathZS=r"XLSX\ZS"
    pathAnlage=r"XLSX\Anlage"
    pathAlle=r"XLSX\Alle"
    pathNutz=r"XLSX\Nutz"
    pathNutzTyp=r"XLSX\NutzTyp"
    pathDSZS=r"XLSX\DSZS"
    pathDSTHZS=r"XLSX\DSTHZS"
    pathTHZS=r"XLSX\THZS"
    pathTHDS=r"XLSX\THDS"

    dfDS=pd.DataFrame()
    dfTH=pd.DataFrame()
    dfZS=pd.DataFrame()
    dfAnlage=pd.DataFrame()
    dfAlle=pd.DataFrame()
    dfNutz=pd.DataFrame()
    dfNutzTyp=pd.DataFrame()
    dfDSZS=pd.DataFrame()
    dfDSTHZS=pd.DataFrame()
    dfTHZS=pd.DataFrame()
    dfTHDS=pd.DataFrame()

    for file1 in listdir(pathDS):
        dfDS = dfDS.append(pd.read_excel(f"{pathDS}\{file1}", header=0, index_col=0))
    for file2 in listdir(pathTH):
        dfTH = dfTH.append(pd.read_excel(f"{pathTH}\{file2}", header=0, index_col=0))
    for file3 in listdir(pathZS):
        dfZS = dfZS.append(pd.read_excel(f"{pathZS}\{file3}", header=0, index_col=0))
    for file4 in listdir(pathAnlage):
        dfAnlage = dfAnlage.append(pd.read_excel(f"{pathAnlage}\{file4}", header=0, index_col=0))
    # for file5 in listdir(pathAlle):
    #     dfAlle = dfAlle.append(pd.read_excel(f"{pathAlle}\{file5}", header=0, index_col=0))
    # for file6 in listdir(pathNutz):
    #     dfNutz = dfNutz.append(pd.read_excel(f"{pathNutz}\{file6}", header=0, index_col=0))
    # for file7 in listdir(pathNutzTyp):
    #     dfNutzTyp = dfNutzTyp.append(pd.read_excel(f"{pathNutzTyp}\{file7}", header=0, index_col=0))
    for file8 in listdir(pathDSZS):
        dfDSZS = dfDSZS.append(pd.read_excel(f"{pathDSZS}\{file8}", header=0, index_col=0))
    for file9 in listdir(pathDSTHZS):
        dfDSTHZS = dfDSTHZS.append(pd.read_excel(f"{pathDSTHZS}\{file9}", header=0, index_col=0))
    for file10 in listdir(pathTHZS):
        dfTHZS = dfTHZS.append(pd.read_excel(f"{pathTHZS}\{file10}", header=0, index_col=0))
    for file11 in listdir(pathTHDS):
        dfTHDS = dfTHDS.append(pd.read_excel(f"{pathTHDS}\{file11}", header=0, index_col=0))
    with pd.ExcelWriter("XLSX\PythonSTATS.xlsx", engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
        dfDS.to_excel(writer, sheet_name="KDS")
        dfTH.to_excel(writer, sheet_name="KTH")
        dfZS.to_excel(writer, sheet_name="KZS")
        dfAnlage.to_excel(writer, sheet_name="KAnlage")
        # dfAlle.to_excel(writer, sheet_name="KAlle")
        # dfNutz.to_excel(writer, sheet_name="KNutz")
        # dfNutzTyp.to_excel(writer, sheet_name="KNutzTyp")
        dfDSZS.to_excel(writer, sheet_name="DSZS")
        dfDSTHZS.to_excel(writer, sheet_name="DSTHZS")
        dfTHZS.to_excel(writer, sheet_name="THZS")
        dfTHDS.to_excel(writer, sheet_name="THDS")
        
    Stat=2
    return Stat

def DWD():
    DWDA1 = pd.read_csv("DWD\produkt_klima_jahr_19490201_20221231_06314.txt", sep=";",decimal=".")
    DWDA2 = pd.read_csv("DWD\produkt_klima_jahr_20190101_20231231_06314.txt", sep=";",decimal=".")

    DWDM1 = pd.read_csv("DWD\produkt_klima_monat_19490201_20221231_06314.txt", sep=";",decimal=".")
    DWDM2 = pd.read_csv("DWD\produkt_klima_monat_20220601_20231231_06314.txt", sep=";",decimal=".")

    lA_TT=[]
    lA_RR=[]
    lMM_TT=[]
    lMM_RR=[]
    lMJE_TT=[]
    lMJE_RR=[]
    lMJY_TT=[]
    lMJY_RR=[]
    lMA_TT=[]
    lMA_RR=[]

    DWDA=DWDA1[["MESS_DATUM_BEGINN","MESS_DATUM_ENDE","JA_TT","JA_RR"]].append(DWDA2[["MESS_DATUM_BEGINN","MESS_DATUM_ENDE","JA_TT","JA_RR"]].loc[4:]).reset_index()
    DWDM=DWDM1[["MESS_DATUM_BEGINN","MESS_DATUM_ENDE","MO_TT","MO_RR"]].append(DWDM2[["MESS_DATUM_BEGINN","MESS_DATUM_ENDE","MO_TT","MO_RR"]].loc[7:]).reset_index()
    #print(DWDA)
    for i in range(0,51):
        if DWDA["JA_TT"].loc[i] != -999:
            lA_TT.append(DWDA["JA_TT"].loc[i])
        if DWDA["JA_RR"].loc[i] != -999:            
            lA_RR.append(DWDA["JA_RR"].loc[i])
    print(f"1950-2023 Nossen | Temp.: {np.mean(lA_TT)} - {74-len(lA_TT)}, Precip.:{np.mean(lA_RR)} - {74-len(lA_RR)}\n2023 Nossen | Temp.: {DWDA['JA_TT'].loc[51]}, Precip.:{DWDA['JA_RR'].loc[51]}")

    for ind in range(0,640):
        #print(str(DWDM["MESS_DATUM_BEGINN"].loc[ind])[4:6])
        if str(DWDM["MESS_DATUM_BEGINN"].loc[ind])[4:6] == "05":
            if DWDM["MO_TT"].loc[ind] != -999:
                lMM_TT.append(DWDM["MO_TT"].loc[ind])
            if DWDM["MO_RR"].loc[ind] != -999:
                lMM_RR.append(DWDM["MO_RR"].loc[ind])
        elif str(DWDM["MESS_DATUM_BEGINN"].loc[ind])[4:6] == "06":
            if DWDM["MO_TT"].loc[ind] != -999:
                lMJE_TT.append(DWDM["MO_TT"].loc[ind])
            if DWDM["MO_RR"].loc[ind] != -999:
                lMJE_RR.append(DWDM["MO_RR"].loc[ind])
        elif str(DWDM["MESS_DATUM_BEGINN"].loc[ind])[4:6] == "07":
            if DWDM["MO_TT"].loc[ind] != -999:
                lMJY_TT.append(DWDM["MO_TT"].loc[ind])
            if DWDM["MO_RR"].loc[ind] != -999:
                lMJY_RR.append(DWDM["MO_RR"].loc[ind])
        elif str(DWDM["MESS_DATUM_BEGINN"].loc[ind])[4:6] == "08":
            if DWDM["MO_TT"].loc[ind] != -999:
                lMA_TT.append(DWDM["MO_TT"].loc[ind])
            if DWDM["MO_RR"].loc[ind] != -999:
                lMA_RR.append(DWDM["MO_RR"].loc[ind])

    print(f"1949-2023 Mai Nossen | Temp.: {np.mean(lMM_TT)} - {75-len(lMM_TT)}, Precip.:{np.mean(lMM_RR)} - {75-len(lMM_RR)}\n2023 Nossen | Temp.: {DWDM['MO_TT'].loc[632]}, Precip.:{DWDM['MO_RR'].loc[632]}")
    print(f"1949-2023 Juni Nossen | Temp.: {np.mean(lMJE_TT)} - {75-len(lMJE_TT)}, Precip.:{np.mean(lMJE_RR)} - {75-len(lMJE_RR)}\n2023 Nossen | Temp.: {DWDM['MO_TT'].loc[633]}, Precip.:{DWDM['MO_RR'].loc[633]}")
    print(f"1949-2023 Juli Nossen | Temp.: {np.mean(lMJY_TT)} - {75-len(lMJY_TT)}, Precip.:{np.mean(lMJY_RR)} - {75-len(lMJY_RR)}\n2023 Nossen | Temp.: {DWDM['MO_TT'].loc[634]}, Precip.:{DWDM['MO_RR'].loc[634]}")
    print(f"1949-2023 August Nossen | Temp.: {np.mean(lMA_TT)} - {75-len(lMA_TT)}, Precip.:{np.mean(lMA_RR)} - {75-len(lMA_RR)}\n2023 Nossen | Temp.: {DWDM['MO_TT'].loc[635]}, Precip.:{DWDM['MO_RR'].loc[635]}")
    print(DWDM[630:])
    Stat=2
    return Stat

def Korr(selection):

    if selection == "DS":
        path=r"XLSX\Korr\DS"
    elif selection == "TH":
        path=r"XLSX\Korr\TH"
    elif selection == "ZS":
        path=r"XLSX\Korr\ZS"
    elif selection == "M":
        path=r"XLSX\Korr\M"
    elif selection == "H":
        path=r"XLSX\Korr\H"
    elif selection == "S":
        path=r"XLSX\Korr\S"
    elif selection == "All": 
        # K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
        # for i in paramKorr:
        #     pd.DataFrame(K.loc[f"{i}"]).to_excel(f"XLSX\Korr\All\{i}.xlsx")
        # path=r"XLSX\Korr\All"
        pTH=r"XLSX\Korr\TH"
        pZS=r"XLSX\Korr\ZS"
        pDS=r"XLSX\Korr\DS"
    elif selection == "Ende": 
        Stat=2
        return Stat

    for ix in paramKorr: #austauschbar in param für andere selections
        for iy in paramKorr: #austauschbar in param für andere selections
            if ix != iy:
                if selection == "TH" or selection == "ZS" or selection == "DS" or selection == "M" or selection == "H" or selection == "S":
                    dfx=pd.read_excel(rf"{pTH}\{ix}.xlsx", index_col=0)
                    dfy=pd.read_excel(rf"{pTH}\{iy}.xlsx", index_col=0)
                elif selection == "All":
                    dfxTH=pd.read_excel(rf"{pTH}\{ix}.xlsx", index_col=0)
                    dfyTH=pd.read_excel(rf"{pTH}\{iy}.xlsx", index_col=0)
                    dfxZS=pd.read_excel(rf"{pZS}\{ix}.xlsx", index_col=0)
                    dfyZS=pd.read_excel(rf"{pZS}\{iy}.xlsx", index_col=0)
                    dfxDS=pd.read_excel(rf"{pDS}\{ix}.xlsx", index_col=0)
                    dfyDS=pd.read_excel(rf"{pDS}\{iy}.xlsx", index_col=0)

                    print(f"{iy}({ix})")              
                    plt.scatter(dfxDS.loc["M"],dfyDS.loc["M"], label="DS", marker="s", c="b")
                    plt.scatter(dfxDS.loc["S"],dfyDS.loc["S"], label="_DS", marker="s", c="b")                       
                    plt.scatter(dfxDS.loc["H"],dfyDS.loc["H"], label="_DS", marker="s", c="b")

                    plt.scatter(dfxTH.loc["M"],dfyTH.loc["M"], label="TH", marker="x", c="r")
                    plt.scatter(dfxTH.loc["H"],dfyTH.loc["H"], label="_TH", marker="x", c="r")
                    plt.scatter(dfxTH.loc["S"],dfyTH.loc["S"], label="_TH", marker="x", c="r")

                    plt.scatter(dfxZS.loc["M"],dfyZS.loc["M"], label="_ZS", marker="o", s=10, c="g")
                    plt.scatter(dfxZS.loc["H"],dfyZS.loc["H"], label="ZS", marker="o", s=10, c="g")
                    plt.scatter(dfxZS.loc["S"],dfyZS.loc["S"], label="_ZS", marker="o", s=10, c="g")
 
                if selection == "TH" or selection == "ZS" or selection == "DS":
                    print(f"{iy}({ix})")
                    plt.title(f"Anlage {selection}")                    
                    plt.scatter(dfx.loc["M"],dfy.loc["M"], label="M", marker="x", c="r")
                    plt.scatter(dfx.loc["H"],dfy.loc["H"], label="H", marker="o", c="g")
                    plt.scatter(dfx.loc["S"],dfy.loc["S"], label="S", marker="s", c="b")
                elif selection == "M" or selection == "H" or selection == "S":
                    print(f"{iy}({ix})")
                    plt.title(f"Flächentyp {selection}")
                    plt.scatter(dfx.loc["TH"],dfy.loc["TH"], label="TH", marker="x", c="r")
                    plt.scatter(dfx.loc["ZS"],dfy.loc["ZS"], label="ZS", marker="o", c="g")
                    plt.scatter(dfx.loc["DS"],dfy.loc["DS"], label="DS", marker="s", c="b")                                                       
                plt.xlabel(ix)
                plt.ylabel(iy)
                plt.legend()
                plt.savefig(rf"Figures\Korr\{selection}\f(x)={iy}({ix}).png", bbox_inches="tight")
                plt.close()
                #plt.show()
    Stat=1
    return Stat

def Discussion():
    D = pd.read_excel("XLSX\PythonSTATS.xlsx", sheet_name="Mediane", header=0, index_col=0)  
    
    plt.figure()
    plt.subplot(6,1,1, ylim=[min(min(D.loc["Artzahl DS"]), min(D.loc["Artzahl ZS"]), min(D.loc["Artzahl TH"]))-1, max(max(D.loc["Artzahl DS"]), max(D.loc["Artzahl ZS"]), max(D.loc["Artzahl TH"]))+1])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["Artzahl DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["Artzahl TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["Artzahl ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[0]]}", rotation="horizontal", ha="right")
    plt.legend(bbox_to_anchor=(0, 1, 1, 0), loc="lower right", ncol=3)
   
    plt.subplot(6,1,3, ylim=[min(min(D.loc["mF DS"]), min(D.loc["mF ZS"]), min(D.loc["mF TH"]))-1, max(max(D.loc["mF DS"]), max(D.loc["mF ZS"]), max(D.loc["mF TH"]))+1])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["mF DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["mF TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["mF ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[2]]}", rotation="horizontal", ha="right")
    
    plt.subplot(6,1,2, ylim=[min(min(D.loc["ngg DS"]), min(D.loc["ngg ZS"]), min(D.loc["ngg TH"]))-1, max(max(D.loc["ngg DS"]), max(D.loc["ngg ZS"]), max(D.loc["ngg TH"]))+1])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["ngg DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["ngg TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["ngg ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[1]]}", rotation="horizontal", ha="right")
    
    plt.subplot(6,1,4, ylim=[min(min(D.loc["DKr DS"]), min(D.loc["DKr ZS"]), min(D.loc["DKr TH"]))-5, max(max(D.loc["DKr DS"]), max(D.loc["DKr ZS"]), max(D.loc["DKr TH"]))+5])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["DKr DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["DKr TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["DKr ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[5]]}", rotation="horizontal", ha="right")
    
    plt.subplot(6,1,5, ylim=[min(min(D.loc["DMoos DS"]), min(D.loc["DMoos ZS"]), min(D.loc["DMoos TH"]))-1, max(max(D.loc["DMoos DS"]), max(D.loc["DMoos ZS"]), max(D.loc["DMoos TH"]))+1])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["DMoos DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["DMoos TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["DMoos ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[3]]}", rotation="horizontal", ha="right")
    
    plt.subplot(6,1,6, ylim=[min(min(D.loc["DStreu DS"]), min(D.loc["DStreu ZS"]), min(D.loc["DStreu TH"]))-1, max(max(D.loc["DStreu DS"]), max(D.loc["DStreu ZS"]), max(D.loc["DStreu TH"]))+1])
    plt.xticks([0,1,2], labels=["M","H","S"])
    plt.plot(D.loc["DStreu DS"], label="DS", marker="o", color="blue", markersize="5")
    plt.plot(D.loc["DStreu TH"], label="TH", marker="s", color="orange", markersize="5")
    plt.plot(D.loc["DStreu ZS"], label="ZS", marker="X", color="green", markersize="5")
    plt.ylabel(f"{paramDict[importantParam[4]]}", rotation="horizontal", ha="right")

    plt.savefig(rf"Figures\Final.png", bbox_inches="tight")
    #plt.show()
    plt.close()
    Status=2
    return Status

if __name__ == "__main__":
    Boot = str(input("Starten: type any string | Ende \n"))
    if Boot != "Ende":
        while Boot != "Ende":
            Status = 0
            Task = str(input("Discussion | Streudiagramme: Korr | Klimadaten: DWD | Kopfdaten auswerten: K | Shinozaki-Kurven: SK | direkter graphischer Vergleich der Anlagen: compare | Zusammenführen der statistischen Analysen: merge | Ende \n"))
            if Task == "K":
                Status=0
                while Status != 2:
                    selection=str(input("DS | TH | ZS | Anlage | AnlageFTyp | Ende \n"))
                    Status=Kopfdaten(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("K beendet")
            elif Task == "SK":
                Status=0
                while Status !=2:
                    Status=Shinozaki()
                    if Status == 2:
                        print("Shinozaki beendet")
            elif Task == "Discussion":
                Status=0
                while Status !=2:
                    Status=Discussion()
                    if Status == 2:
                        print("Discussion beendet")
            elif Task == "compare":
                Status=0
                while Status != 2:
                        selection=str(input("MHS | THZSDS | FECO \n"))
                        if selection == "MHS":
                            Status=compareFTyp()
                        elif selection == "THZSDS":
                            Status=compareAnlagen()
                        elif selection == "FECO":
                            Status=compareFECO()
                if Status == 2:
                    print("compare beendet")
            elif Task == "merge":
                Status=0
                while Status !=2:
                    Status=merge()
                    if Status == 2:
                        print("merge beendet")
            elif Task =="DWD":
                Status=0
                while Status !=2:
                    Status=DWD()
                    if Status == 2:
                        print("DWD beendet")
            elif Task =="Korr":
                while Status !=2:
                    selection=str(input("All | DS | TH | ZS | M | H | S | Ende \n"))
                    Status=Korr(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("Korr beendet")
            elif Task == "Ende":
                Boot = str(input("Starten: type any string | Ende \n"))

    

    
    
