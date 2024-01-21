import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
from os import listdir, path
from PIL import Image as img
import math
import numpy as np
from scipy import stats as st
import statsmodels.stats.multitest as mt

#DSKM_mL= pd.read_csv("DSKM_mL.txt", sep=";",decimal=",")

#paramDict = {"Dges [%]":0,"DKr [%]":4,"DMoos [%]" :8,"DStreu [%]":12,"HKr [cm]":16,"mT":20,"mL":24,"mF":28,"mR":32,"mN":36,"mM":40,"mW":44,"mTr":48,"AntThero":52,"AntHemi":56,"E":60,"Hs":64,"AbS":68,"Artzahl":72}
param = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]","mT","mL", "mL2","mF","mR","mN","mM","mW","mTr","AntThero", "AntThero (Ti)", "n_Thero (Ti)","AntHemi", "AntHemi (Ti)", "n_Hemi (Ti)","E","Hs","AbS","Artzahl"]
vegparam = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]"]
staparam = ["mT","mL", "mL2","mF","mR","mN","mM","mW","mTr"]
divparam=["AntThero", "AntThero (Ti)", "n_Thero (Ti)","AntHemi", "AntHemi (Ti)", "n_Hemi (Ti)","E","Hs","AbS","Artzahl"]
l = []

# def KopfdatenMW(selection):
#     Stat = 0
#     if selection == "DS":
#         KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - DS", header=1, index_col=0)
#     elif selection == "TH":
#         KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - TH", header=1, index_col=0)
#     elif selection == "ZS":
#         KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - ZS", header=1, index_col=0)
#     elif selection == "Anlage":
#         KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - Anlage", header=1, index_col=0)
#     elif selection == "Ende":
#         Stat=2
#         return Stat
#     #print(type(KM.index[1])==type(param[1]))
#     for i in KM.index:
#         l.append(KM.loc[f"{i}"])

#     for ind in range(len(param)):
#     #    print(l[paramDict[f"{param[ind]}"]])
#         ind_df=pd.DataFrame(l[paramDict[f"{param[ind]}"]])
#     #    print(ind_df)
#         plt.errorbar(ind_df["Typ"],ind_df["Mittelwert"], ind_df["StdAbw"], fmt='ok', lw=3, label="Mittelwert und Standardabweichung")
#         plt.errorbar(ind_df["Typ"],ind_df["Median"], fmt="ro", label="Median")
#         plt.title(f"{ind_df.index[0]} - {selection}", loc="left", fontsize=12)
#         plt.legend(bbox_to_anchor=(0,1,1,0), loc="lower right", fontsize=8)
# #        plt.tight_layout()
# #        plt.show()
#         plt.savefig(rf"Figures\KopfdatenMW\{selection}\{ind_df.index[0]}.png", bbox_inches="tight")
#         plt.close()
#     Stat=1
#     return Stat

def Kopfdaten(selection):
    Stat = 0
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
    # if selection == "spear":
    #     SPR=pd.DataFrame()
    #     SPRE=pd.DataFrame(index=param, columns=param)
    #     DSList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6', 'DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6', 'DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6']
    #     THList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4', 'THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4', 'THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4']
    #     ZSList=['ZSM', 'ZSM.1', 'ZSM.2', 'ZSM.3', 'ZSH', 'ZSH.1', 'ZSH.2', 'ZSH.3', 'ZSS', 'ZSS.1', 'ZSS.2', 'ZSS.3']

    #     for i in param:
    #         SPR_single=pd.DataFrame(K.loc[f"{i}"])
    #         SPR=pd.concat([SPR , SPR_single], axis=1)
    #     sel=str(input("Alle | DS | TH | ZS\n"))
    #     if sel=="Alle":
    #         print(SPR)
    #         #print(SPR[param[0]].loc["DSS"])
    #         print(np.array(SPR[param[0]]))
    #         for iHN in range(len(param)):
    #             for iVN in range(len(param)):
    #                 #print(f"{param[iHN]} - {param[iVN]} | {st.mstats.spearmanr(np.array(SPR[param[iHN]], dtype=float),np.array(SPR[param[iVN]], dtype=float), nan_policy='omit')}")
    #                 SPRE[param[iHN]].loc[param[iVN]]=(f"s={st.mstats.spearmanr(np.array(SPR[param[iHN]], dtype=float),np.array(SPR[param[iVN]], dtype=float), nan_policy='omit')[0]}",f"p={st.mstats.spearmanr(np.array(SPR[param[iHN]], dtype=float),np.array(SPR[param[iVN]], dtype=float), nan_policy='omit')[1]}")
    #         #print(SPRE)
    #         with pd.ExcelWriter("XLSX\PythonSTATS.xlsx", engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
    #             SPRE.to_excel(writer, sheet_name="SpearAll")
    #     elif sel == "DS":
    #         #print(SPR)
    #         SPRDS=SPR.iloc[:len(DSList)]
    #         #print(SPRDS)
    #         for iHN in range(len(param)):
    #             for iVN in range(len(param)):
    #                 SPRE[param[iHN]].loc[param[iVN]]=(f"s={st.mstats.spearmanr(np.array(SPRDS[param[iHN]], dtype=float),np.array(SPRDS[param[iVN]], dtype=float), nan_policy='omit')[0]}",f"p={st.mstats.spearmanr(np.array(SPRDS[param[iHN]], dtype=float),np.array(SPRDS[param[iVN]], dtype=float), nan_policy='omit')[1]}")
    #         with pd.ExcelWriter("XLSX\PythonSTATS.xlsx", engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
    #             SPRE.to_excel(writer, sheet_name="SpearDS")           
    #     elif sel == "TH":
    #         SPRTH=SPR.iloc[len(DSList):(len(DSList)+len(THList))]
    #         print(SPRTH)
    #         for iHN in range(len(param)):
    #             for iVN in range(len(param)):
    #                 SPRE[param[iHN]].loc[param[iVN]]=(f"s={st.mstats.spearmanr(np.array(SPRTH[param[iHN]], dtype=float),np.array(SPRTH[param[iVN]], dtype=float), nan_policy='omit')[0]}",f"p={st.mstats.spearmanr(np.array(SPRTH[param[iHN]], dtype=float),np.array(SPRTH[param[iVN]], dtype=float), nan_policy='omit')[1]}")
    #         with pd.ExcelWriter("XLSX\PythonSTATS.xlsx", engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
    #             SPRE.to_excel(writer, sheet_name="SpearTH")
    #     elif sel == "ZS":
    #         SPRZS=SPR.iloc[(len(DSList)+len(THList)):(len(DSList)+len(THList)+len(ZSList))]
    #         print(SPRZS)
    #         for iHN in range(len(param)):
    #             for iVN in range(len(param)):
    #                 SPRE[param[iHN]].loc[param[iVN]]=(f"s={st.mstats.spearmanr(np.array(SPRZS[param[iHN]], dtype=float),np.array(SPRZS[param[iVN]], dtype=float), nan_policy='omit')[0]}",f"p={st.mstats.spearmanr(np.array(SPRZS[param[iHN]], dtype=float),np.array(SPRZS[param[iVN]], dtype=float), nan_policy='omit')[1]}")
    #         with pd.ExcelWriter("XLSX\PythonSTATS.xlsx", engine="openpyxl",mode="a", if_sheet_exists="replace") as writer:
    #             SPRE.to_excel(writer, sheet_name="SpearZS")
           
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

            KWlist=[]
            KWlist.append(st.kruskal(KPM, KPH)[1][0])
            KWlist.append(st.kruskal(KPM, KPS)[1][0])
            KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)
            # np.std | ddof changes divisor -> ddof=1 -> divosr=N-1 => as is statgraphics
            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0],st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0],"","", "","",""]
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0],"","", "","",""]
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std","KW","p","M-H","M-S","H-S"]).to_excel(f"XLSX\{selection}\{i}.xlsx")
            
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
            
            KWlist=[]
            KWlist.append(st.kruskal(KPM, KPH)[1][0])
            KWlist.append(st.kruskal(KPM, KPS)[1][0])
            KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)

            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0],st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0],"","", "","",""]
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0],"","", "","",""]
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std","KW","p","M-H","M-S","H-S"]).to_excel(f"XLSX\{selection}\{i}.xlsx")

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
            KWlist=[]
            KWlist.append(st.kruskal(KPM, KPH)[1][0])
            KWlist.append(st.kruskal(KPM, KPS)[1][0])
            KWlist.append(st.kruskal(KPH, KPS)[1][0])
            #print(KWlist)

            Mdata=[i,len(KPM),np.mean(KPM)[0],np.median(KPM),np.std(KPM, ddof=1)[0],st.kruskal(KPM,KPH,KPS)[0],st.kruskal(KPM,KPH,KPS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]]
            Hdata=[i,len(KPH),np.mean(KPH)[0],np.median(KPH),np.std(KPH, ddof=1)[0],"","", "","",""]
            Sdata=[i,len(KPS),np.mean(KPS)[0],np.median(KPS),np.std(KPS, ddof=1)[0],"","", "","",""]
            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std","KW","p","M-H","M-S","H-S"]).to_excel(f"XLSX\{selection}\{i}.xlsx")

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
            
            KWlist=[]
            KWlist.append(st.kruskal(KPDS, KPTH)[1][0])
            KWlist.append(st.kruskal(KPDS, KPZS)[1][0])
            KWlist.append(st.kruskal(KPTH, KPZS)[1][0])

            Mdata=[i,len(KPDS),np.mean(KPDS)[0],np.median(KPDS),np.std(KPDS, ddof=1)[0],st.kruskal(KPDS,KPTH,KPZS)[0],st.kruskal(KPDS,KPTH,KPZS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]]
            Hdata=[i,len(KPTH),np.mean(KPTH)[0],np.median(KPTH),np.std(KPTH, ddof=1)[0],"","", "","",""]
            Sdata=[i,len(KPZS),np.mean(KPZS)[0],np.median(KPZS),np.std(KPZS, ddof=1)[0],"","", "","",""]

            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["DS","TH","ZS"], columns=["Parameter","n","mean","median","std","KW","p","DS-TH","DS-ZS","TH-ZS"]).to_excel(f"XLSX\{selection}\{i}.xlsx")
            
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

        elif selection == "Alle":
            AM=[]
            AH=[]
            AS=[]
            AS1=[]
            AMList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6','THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4','ZSM','ZSM.1','ZSM.2','ZSM.3']
            AHList=['DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6','THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4','ZSH','ZSH.1','ZSH.2','ZSH.3'] 
            ASList=['DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6','THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4','ZSS','ZSS.1','ZSS.2','ZSS.3']
            for indM in AMList:
#                print(KP.loc[f"{ind}"])
                AM.append(float(KP.loc[f"{indM}"]))
            for indH in AHList:
                AH.append(float(KP.loc[f"{indH}"]))
            for indS in ASList:
                AS1.append(float(KP.loc[f"{indS}"]))

            for iK in AS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    AS.append(iK)

            KAM=pd.DataFrame(AM)
            KAH=pd.DataFrame(AH)
            KAS=pd.DataFrame(AS)

            KWlist=[]
            KWlist.append(st.kruskal(KAM, KAH)[1][0])
            KWlist.append(st.kruskal(KAM, KAS)[1][0])
            KWlist.append(st.kruskal(KAH, KAS)[1][0])

            Mdata=[i,len(KAM),np.mean(KAM)[0],np.median(KAM),np.std(KAM, ddof=1)[0],st.kruskal(KAM,KAH,KAS)[0],st.kruskal(KAM,KAH,KAS)[1], mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(KWlist, alpha=0.05, method='bonferroni')[0][2]]
            Hdata=[i,len(KAH),np.mean(KAH)[0],np.median(KAH),np.std(KAH, ddof=1)[0],"","", "","",""]
            Sdata=[i,len(KAS),np.mean(KAS)[0],np.median(KAS),np.std(KAS, ddof=1)[0],"","", "","",""]

            pd.DataFrame(data=[Mdata,Hdata,Sdata], index=["M","H","S"], columns=["Parameter","n","mean","median","std","KW","p","M-H","M-S","H-S"]).to_excel(f"XLSX\{selection}\{i}.xlsx")

            print(f"{i} - {st.kruskal(AM,AH,AS)}")            
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - DS, TH & ZS ", loc="left")
            plt.boxplot(KAM, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KAH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KAS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()

        elif selection == "Nutz":
            KBe=[]
            KMa=[]
            KMa1=[]
            KBeList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4','ZSM','ZSM.1','ZSM.2','ZSM.3', 'THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4','ZSH','ZSH.1','ZSH.2','ZSH.3','THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4','ZSS','ZSS.1','ZSS.2','ZSS.3' ]
            KMaList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6','DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6', 'DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6'] 

            for indBe in KBeList:
#                print(KP.loc[f"{ind}"])
                KBe.append(float(KP.loc[f"{indBe}"]))
            for indMa in KMaList:
                KMa1.append(float(KP.loc[f"{indMa}"]))
            
            for iK in KMa1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    KMa.append(iK)

            KBE=pd.DataFrame(KBe)
            KMA=pd.DataFrame(KMa)
            #print(f"{i} - {st.kruskal(KBe,KMa)}")

            KBEdata=[i,len(KBE),np.mean(KBE)[0],np.median(KBE),np.std(KBE, ddof=1)[0],st.mannwhitneyu(KBE,KMA)[0],st.mannwhitneyu(KBE,KMA)[1]]
            KMAdata=[i,len(KMA),np.mean(KMA)[0],np.median(KMA),np.std(KMA, ddof=1)[0],"",""]

            #print(KBEdata)
            pd.DataFrame(data=[KBEdata,KMAdata], index=["Be","Ma"], columns=["Parameter","n","mean","median","std","U","p"]).to_excel(f"XLSX\{selection}\{i}.xlsx")

            plt.figure()
            ax1=plt.subplot(1,2,1)
            plt.title(f"{i}", loc="left")
            plt.boxplot(KMA, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,2,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KBE, labels="B", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()

        elif selection == "NutzTyp":
            KBeM=[]
            KBeH=[]
            KBeS=[]
            KMaM=[]
            KMaH=[]
            KMaS=[]
            KMaS1=[]
            KBeMList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4','ZSM','ZSM.1','ZSM.2','ZSM.3']
            KBeHList=['THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4','ZSH','ZSH.1','ZSH.2','ZSH.3']
            KBeSList=['THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4','ZSS','ZSS.1','ZSS.2','ZSS.3' ]
            KMaMList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6'] 
            KMaHList=['DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6']
            KMaSList=['DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6']
            
            for indBeM in KBeMList:
#                print(KP.loc[f"{ind}"])
                KBeM.append(float(KP.loc[f"{indBeM}"]))
            for indBeH in KBeHList:
                KBeH.append(float(KP.loc[f"{indBeH}"]))
            for indBeS in KBeSList:
                KBeS.append(float(KP.loc[f"{indBeS}"]))
            for indMaM in KMaMList:
                KMaM.append(float(KP.loc[f"{indMaM}"]))
            for indMaH in KMaHList:
                KMaH.append(float(KP.loc[f"{indMaH}"]))
            for indMaS in KMaSList:
                KMaS1.append(float(KP.loc[f"{indMaS}"]))

            for iK in KMaS1:
                #print({str(iK).replace(".","",1).isdigit()})
                if str(iK).replace(".","",1).isdigit():
                    KMaS.append(iK)

            KBEM=pd.DataFrame(KBeM)
            KBEH=pd.DataFrame(KBeH)
            KBES=pd.DataFrame(KBeS)
            KMAM=pd.DataFrame(KMaM)
            KMAH=pd.DataFrame(KMaH)
            KMAS=pd.DataFrame(KMaS)

            UlistBE=[]
            UlistBE.append(st.mannwhitneyu(KBEM, KBEH)[1][0])
            UlistBE.append(st.mannwhitneyu(KBEM, KBES)[1][0])
            UlistBE.append(st.mannwhitneyu(KBEH, KBES)[1][0])

            UlistMA=[]
            UlistMA.append(st.mannwhitneyu(KMAM, KMAH)[1][0])
            UlistMA.append(st.mannwhitneyu(KMAM, KMAS)[1][0])
            UlistMA.append(st.mannwhitneyu(KMAH, KMAS)[1][0])

            KBEMdata=[f"{i} - Bew",len(KBEM),np.mean(KBEM)[0],np.median(KBEM),np.std(KBEM, ddof=1)[0],st.kruskal(KBEM,KBEH,KBES)[0],st.kruskal(KBEM,KBEH,KBES)[1], mt.multipletests(UlistBE, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(UlistBE, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(UlistBE, alpha=0.05, method='bonferroni')[0][2],"",st.mannwhitneyu(KBEM,KMAM)[0],st.mannwhitneyu(KBEM,KMAM)[1],st.mannwhitneyu(KBEH,KMAH)[0],st.mannwhitneyu(KBEH,KMAH)[1],st.mannwhitneyu(KBES,KMAS)[0],st.mannwhitneyu(KBES,KMAS)[1]]
            KBEHdata=[f"{i} - Bew",len(KBEH),np.mean(KBEH)[0],np.median(KBEH),np.std(KBEH, ddof=1)[0],"","", "","","","","","",""]
            KBESdata=[f"{i} - Bew",len(KBES),np.mean(KBES)[0],np.median(KBES),np.std(KBES, ddof=1)[0],"","", "","","","","","",""]
            KMAMdata=[f"{i} - Mahd",len(KMAM),np.mean(KMAM)[0],np.median(KMAM),np.std(KMAM, ddof=1)[0],st.kruskal(KMAM,KMAH,KMAS)[0],st.kruskal(KMAM,KMAH,KMAS)[1], mt.multipletests(UlistMA, alpha=0.05, method='bonferroni')[0][0],mt.multipletests(UlistMA, alpha=0.05, method='bonferroni')[0][1],mt.multipletests(UlistMA, alpha=0.05, method='bonferroni')[0][2],"","","",""]
            KMAHdata=[f"{i} - Mahd",len(KMAH),np.mean(KMAH)[0],np.median(KMAH),np.std(KMAH, ddof=1)[0],"","", "","","","","","",""]
            KMASdata=[f"{i} - Mahd",len(KMAS),np.mean(KMAS)[0],np.median(KMAS),np.std(KMAS, ddof=1)[0],"","", "","","","","","",""]

            pd.DataFrame(data=[KBEMdata,KBEHdata,KBESdata,KMAMdata,KMAHdata,KMASdata], index=["B-M","B-H","B-S","M-M","M-H","M-S"], columns=["Parameter","n","mean","median","std","KW","p","M-H","M-S","H-S", "","U-MB-MM","p-MB-MM","U-HB-HM","p-HB-HM","U-SB-SM","p-SB-SM"]).to_excel(f"XLSX\{selection}\{i}.xlsx")
            
            #print(f"{i} Beweidung - {st.kruskal(KBeM,KBeH,KBeS)}")
            #print(f"{i} Mahd - {st.kruskal(KMaM,KMaH,KMaS)}")
            #print(f"{i} M - {st.kruskal(KBeM,KMaM)}")
            #print(f"{i} H - {st.kruskal(KBeH,KMaH)}")
            #print(f"{i} S - {st.kruskal(KBeS,KMaS)}")

            plt.figure()
            ax1=plt.subplot(1,7,1)
            plt.title(f"{i} - Mahd", loc="left")
            plt.boxplot(KMAM, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,7,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KMAH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,7,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KMAS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax4=plt.subplot(1,7,5, sharey=ax1)
            plt.title(f"{i} - Beweidung", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KBEM, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax5=plt.subplot(1,7,6, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KBEH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,7,7, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KBES, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
#            plt.show()
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()

        elif selection == "Ende":
            Stat=2
            return Stat
    Stat=1
    return Stat

def Shinozaki():
    S=pd.read_excel("SORTstatistics.xlsx", sheet_name="Shinozaki - Anlage", header=2, index_col=0)
#    print(S)
    plt.xticks(S.index, labels=[1,"",3,"",5,"",7,"",9,"",11,"",13,"",15,"",17,"",19,"",21])
    plt.plot(S["DS - S"], label="DS", marker="o")
    plt.plot(S["TH - S"], label="TH", marker="s")
    plt.plot(S["ZS - S"], label="ZS", marker="X")
    plt.legend()
    plt.title("Shinozaki-Kurve")
    plt.savefig(rf"Figures\Shinozaki.png", bbox_inches="tight")
#    plt.show()
    plt.close()
    Stat=2
    return Stat

def Grouped(selection):
    pathDS=r"Figures\Kopfdaten\DS"
    pathTH =r"Figures\Kopfdaten\TH"
    pathZS =r"Figures\Kopfdaten\ZS"
    picturesDS=listdir(pathDS)
    picturesTH=listdir(pathTH)
    picturesZS=listdir(pathZS)
    #print(picturesDS)

    if selection == "small":
        for i in range(len(picturesDS)):
            imDS=img.open(f"{pathDS}\{picturesDS[i]}")
        #    imDS.show()
        #    print(imDS.size)
            imTH=img.open(f"{pathTH}\{picturesTH[i]}")
            imZS=img.open(f"{pathZS}\{picturesZS[i]}")
            if imDS.size[0] == imTH.size[0] == imZS.size[0]:
                Gw=imDS.size[0]
                Gh=imDS.size[1]+imTH.size[1]+imZS.size[1]
                imGroup=img.new("RGB", size=(Gw,Gh), color=(255,255,255))
                imGroup.paste(imDS,(0,0))
                imGroup.paste(imTH, (0, imDS.size[1]))
                imGroup.paste(imZS, (0, imDS.size[1]+imTH.size[1]))
        #        imGroup.show()
                imGroup.save(rf"Figures\Grouped\{picturesDS[i]}.png")
        #        print("alle gleiche Größe")
            else: 
                Gw=max([imDS.size[0], imTH.size[0], imZS.size[0]])
        #        print(f"{Gw}\n DS: {imDS.size[0]} TH: {imTH.size[0]} ZS: {imZS.size[0]}")
                Gh=imDS.size[1]+imTH.size[1]+imZS.size[1]
                imGroup=img.new("RGB", size=(Gw,Gh), color=(255,255,255))
                imGroup.paste(imDS, (int((Gw-imDS.size[0])/2),0))
                imGroup.paste(imTH, (int((Gw-imTH.size[0])/2), imDS.size[1]))
                imGroup.paste(imZS, (int((Gw-imZS.size[0])/2), imDS.size[1]+imTH.size[1]))
        #        imGroup.show()
                imGroup.save(rf"Figures\Grouped\{picturesDS[i]}.png")
        #        print("Nope")
            imDS.close()
            imTH.close()
            imZS.close()

        Stat=1
        return Stat
    
    elif selection == "big":
        pathA =r"Figures\Kopfdaten\Anlage"
        picturesA=listdir(pathA)
        #print(picturesA)
        for i in range(len(picturesA)):
            imDS=img.open(f"{pathDS}\{picturesDS[i]}")
        #    imDS.show()
            imTH=img.open(f"{pathTH}\{picturesTH[i]}")
            imZS=img.open(f"{pathZS}\{picturesZS[i]}")
            imA=img.open(f"{pathA}\{picturesA[i]}")
#            print(f"DS: {imDS.size}\nTH: {imTH.size}\nZS: {imZS.size}\nA:{imA.size}")
            if imDS.size[0] == imTH.size[0] == imZS.size[0] == imA.size[0]:
                Gw=imDS.size[0]
                Gh=imDS.size[1]+imTH.size[1]+imZS.size[1]+imA.size[1]
                imGroup=img.new("RGB", size=(Gw,Gh), color=(255,255,255))
                imGroup.paste(imDS,(0,0))
                imGroup.paste(imTH, (0, imDS.size[1]))
                imGroup.paste(imZS, (0, imDS.size[1]+imTH.size[1]))
                imGroup.paste(imA, (0, imDS.size[1]+imTH.size[1]+imZS.size[1]))
        #        imGroup.show()
                imGroup.save(rf"Figures\Grouped\{picturesDS[i]}.png")
#                print("alle gleiche Größe")
            else: 
                Gw=max([imDS.size[0], imTH.size[0], imZS.size[0], imA.size[0]])
        #        print(f"{Gw}\n DS: {imDS.size[0]} TH: {imTH.size[0]} ZS: {imZS.size[0]}")
                Gh=imDS.size[1]+imTH.size[1]+imZS.size[1]+imA.size[1]
                imGroup=img.new("RGB", size=(Gw,Gh), color=(255,255,255))
                imGroup.paste(imDS, (int((Gw-imDS.size[0])/2),0))
                imGroup.paste(imTH, (int((Gw-imTH.size[0])/2), imDS.size[1]))
                imGroup.paste(imZS, (int((Gw-imZS.size[0])/2), imDS.size[1]+imTH.size[1]))
                imGroup.paste(imA, (int((Gw-imA.size[0])/2), imDS.size[1]+imTH.size[1]+imZS.size[1]))
        #        imGroup.show()
                imGroup.save(rf"Figures\BigGrouped\{picturesA[i]}.png")
#                print("Nope")
            imDS.close()
            imTH.close()
            imZS.close()
            imA.close()

        Stat=1
        return Stat
    
    elif selection == "Ende":
        Stat=2
        return Stat
    
def compare(selection):
    Stat = 0
    K = pd.read_excel("Kopfdaten.xlsx", sheet_name="nur relevante Kopfdaten", header=1, index_col=0)
    for i in param:
        KP=pd.DataFrame(K.loc[f"{i}"])
#        print(KP)
        if selection == "DS-TH-ZS":     
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
            KPDSZSutest=[st.mannwhitneyu(KPMDS,KPMZS)[0],st.mannwhitneyu(KPMDS,KPMZS)[1],st.mannwhitneyu(KPHDS,KPHZS)[0],st.mannwhitneyu(KPHDS,KPHZS)[1],st.mannwhitneyu(KPSDS,KPSZS)[0],st.mannwhitneyu(KPSDS,KPSZS)[1]]
            pd.DataFrame(data=[KPDSZSutest], index=[i], columns=["U: MDS-MZS","p: MDS-MZS","U: HDS-HZS","p: HDS-HZS","U: SDS-SZS","p: SDS-SZS"]).to_excel(f"XLSX\DSZS\{i}.xlsx")
            KPDSTHZSkwtest=[st.kruskal(KPMDS,KPMTH,KPMZS)[0],st.kruskal(KPMDS,KPMTH,KPMZS)[1],st.kruskal(KPHDS,KPHTH,KPHZS)[0],st.kruskal(KPHDS,KPHTH,KPHZS)[1],st.kruskal(KPSDS,KPSTH,KPSZS)[0],st.kruskal(KPSDS,KPSTH,KPSZS)[1]]
            #print(f"{i} - {KPMDS}")
            #print(f"{i} - {max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))}")
            pd.DataFrame(data=[KPDSTHZSkwtest], index=[i], columns=["KW: M-M","p: M-M","KW: H-H","p: H-H","KW: S-S","p: S-S"]).to_excel(f"XLSX\DSTHZS\{i}.xlsx")
            
            yMAX=max(max(KPMDS[0]),max(KPHDS[0]),max(KPSDS[0]),max(KPMTH[0]),max(KPHTH[0]), max(KPSTH[0]),max(KPMZS[0]),max(KPHZS[0]),max(KPSZS[0]))
            yMIN=min(min(KPMDS[0]),min(KPHDS[0]),min(KPSDS[0]),min(KPMTH[0]),min(KPHTH[0]), min(KPSTH[0]),min(KPMZS[0]),min(KPHZS[0]),min(KPSZS[0]))
           
            if i in vegparam:
                yADD=15-(math.ceil(yMAX)%10) 
                ySUB=5+(yMIN%5)
            elif i in staparam:
                yADD=1.2-(yMIN%1)
                ySUB=0.2+(yMIN%0.2)
            elif i in divparam:
                if i == "Hs":
                    yADD=1.2-(yMIN%1)
                    ySUB=0.2+(yMIN%0.2)
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
                else:
                    ax1.yaxis.set_major_locator(ticker.MultipleLocator(10))
                    ax1.yaxis.set_minor_locator(ticker.MultipleLocator(5))
            
            plt.title(f"{i} - TH", loc="left")
            plt.boxplot(KPMTH, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax2=plt.subplot(1,11,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHTH, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax3=plt.subplot(1,11,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSTH, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax4=plt.subplot(1,11,5, sharey=ax1)
            plt.title("ZS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMZS, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})  
            ax5=plt.subplot(1,11,6, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHZS, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax6=plt.subplot(1,11,7, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSZS, labels="S", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax7=plt.subplot(1,11,9, sharey=ax1)
            plt.title("DS", loc="left")
            plt.tick_params("y", labelleft=True)
            plt.boxplot(KPMDS, labels="M", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax8=plt.subplot(1,11,10, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPHDS, labels="H", showmeans=True,medianprops={"color":"blue"}, meanprops={"marker":"+"})
            ax9=plt.subplot(1,11,11, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPSDS, labels="S", showmeans=True, medianprops={"color":"blue"},meanprops={"marker":"+"})
            plt.savefig(rf"Figures\Compare\{i}.png", bbox_inches="")
            plt.close()  
        elif selection == "Ende": 
            Stat=2
            return Stat
    Stat=1
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

    dfDS=pd.DataFrame()
    dfTH=pd.DataFrame()
    dfZS=pd.DataFrame()
    dfAnlage=pd.DataFrame()
    dfAlle=pd.DataFrame()
    dfNutz=pd.DataFrame()
    dfNutzTyp=pd.DataFrame()
    dfDSZS=pd.DataFrame()
    dfDSTHZS=pd.DataFrame()

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
    elif selection == "Ende": 
        Stat=2
        return Stat

    for ix in param:
        for iy in param:
            if ix != iy:
                dfx=pd.read_excel(rf"{path}\{ix}.xlsx", index_col=0)
                dfy=pd.read_excel(rf"{path}\{iy}.xlsx", index_col=0)
                plt.scatter(dfx.loc["M"],dfy.loc["M"], label="M", marker="x", c="r")
                plt.scatter(dfx.loc["H"],dfy.loc["H"], label="H", marker="o", c="g")
                plt.scatter(dfx.loc["S"],dfy.loc["S"], label="S", marker="s", c="b")
                plt.xlabel(ix)
                plt.ylabel(iy)
                plt.legend()
                plt.savefig(rf"Figures\Korr\{selection}\f(x)={iy}({ix}).png", bbox_inches="tight")
                plt.close()
                #plt.show()
    Stat=1
    return Stat

if __name__ == "__main__":
    Boot = str(input("Starten: type any string | Ende \n"))
    if Boot != "Ende":
        while Boot != "Ende":
            Status = 0
            Task = str(input("Korr | DWD | KopfdatenMW auswerten: KMW | Kopfdaten auswerten: K | Shinozaki-Kurven: SK | Grouped: G | compare | merge | Ende \n"))
            if Task == "KMW":
                while Status != 2:
                    selection=str(input("DS | TH | ZS | Anlage | Ende \n"))
                    #Status=KopfdatenMW(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("KMW beendet")
            elif Task == "K":
                Status=0
                while Status != 2:
                    selection=str(input("(spear) | DS | TH | ZS | Anlage | (Alle)  | (Nutz) | (NutzTyp) | Ende \n"))
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
            elif Task == "compare":
                while Status != 2:
                        selection=str(input("DS-TH-ZS | Typ-Anl-Nutz | Ende \n"))
                        Status=compare(selection)
                        if Status == 1:
                            print(f"{selection} Done")
                        elif Status == 2:
                            print("compare beendet")
            elif Task == "merge":
                Status=0
                while Status !=2:
                    Status=merge()
                    if Status == 2:
                        print("merge beendet")
            elif Task == "G":
                while Status != 2:
                    selection=str(input("small | big | Ende \n"))
                    Status=Grouped(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("G beendet")
            elif Task =="DWD":
                Status=0
                while Status !=2:
                    Status=DWD()
                    if Status == 2:
                        print("DWD beendet")
            elif Task =="Korr":
                while Status !=2:
                    selection=str(input("DS | TH | ZS | Ende \n"))
                    Status=Korr(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("Korr beendet")
            elif Task == "Ende":
                Boot = str(input("Starten: type any string | Ende \n"))

# important/used parts -> put in one program
                # Shinozaki
                # Kopfdaten (ohne Spear)
                # Compare und merge
                # DWD
                # Korr
    

    
    
