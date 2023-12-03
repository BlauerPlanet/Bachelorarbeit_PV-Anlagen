# als Errorbars statt Boxplots geplottet, da keine Quantile in Daten
import matplotlib.pyplot as plt
import pandas as pd
#DSKM_mL= pd.read_csv("DSKM_mL.txt", sep=";",decimal=",")

paramDict = {"Dges [%]":0,"DKr [%]":4,"DMoos [%]":8,"DStreu [%]":12,"HKr [cm]":16,"mT":20,"mL":24,"mF":28,"mR":32,"mN":36,"mM":40,"mW":44,"mTr":48,"AntThero":52,"AntHemi":56,"E":60,"Hs":64,"AbS":68,"Artzahl":72}
param = ["Dges [%]","DKr [%]","DMoos [%]","DStreu [%]","HKr [cm]","mT","mL","mF","mR","mN","mM","mW","mTr","AntThero","AntHemi","E","Hs","AbS","Artzahl"]
l = []

def KopfdatenMW(selection):
    Stat = 0
    if selection == "DS":
        KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - DS", header=1, index_col=0)
    elif selection == "TH":
        KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - TH", header=1, index_col=0)
    elif selection == "ZS":
        KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - FTyp - ZS", header=1, index_col=0)
    elif selection == "Anlage":
        KM = pd.read_excel("SORTstatistics.xlsx", sheet_name="KopfdatenMW - Anlage", header=1, index_col=0)
    elif selection == "Ende":
        Stat=2
        return Stat
#    print(type(KM.index[1])==type(param[1]))
    for i in KM.index:
        l.append(KM.loc[f"{i}"])

    for ind in range(len(param)):
    #    print(l[paramDict[f"{param[ind]}"]])
        ind_df=pd.DataFrame(l[paramDict[f"{param[ind]}"]])
    #    print(ind_df)
        plt.errorbar(ind_df["Typ"],ind_df["Mittelwert"], ind_df["StdAbw"], fmt='ok', lw=3, label="Mittelwert und Standardabweichung")
        plt.errorbar(ind_df["Typ"],ind_df["Median"], fmt="ro", label="Median")
        plt.title(f"{ind_df.index[0]} - {selection}", loc="left", fontsize=12)
        plt.legend(bbox_to_anchor=(0,1,1,0), loc="lower right", fontsize=8)
#        plt.tight_layout()
#        plt.show()
        plt.savefig(rf"Figures\KopfdatenMW\{selection}\{ind_df.index[0]}.png", bbox_inches="tight")
        plt.close()
    Stat=1
    return Stat

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
            DSMList=["DSM","DSM.1","DSM.2","DSM.3","DSM.4","DSM.5","DSM.6"]
            DSHList=["DSH","DSH.1","DSH.2","DSH.3","DSH.4","DSH.5","DSH.6"] 
            DSSList=["DSS","DSS.1","DSS.2","DSS.3","DSS.4","DSS.5","DSS.6"]
            for indM in DSMList:
#                print(KP.loc[f"{ind}"])
                DSM.append(float(KP.loc[f"{indM}"]))
            for indH in DSHList:
                DSH.append(float(KP.loc[f"{indH}"]))
            for indS in DSSList:
                DSS.append(float(KP.loc[f"{indS}"]))

            KPM=pd.DataFrame(DSM)
            KPH=pd.DataFrame(DSH)
            KPS=pd.DataFrame(DSS)
#            print(KPM)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True)
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True)
            ax3=plt.subplot(1,3,3,sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True)
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
#            print(KPM)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True)
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True)
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True)
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
#            print(KPM)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - {selection}", loc="left")
            plt.boxplot(KPM, labels="M", showmeans=True)
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPH, labels="H", showmeans=True)
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPS, labels="S", showmeans=True)
#            plt.show()   
            plt.savefig(rf"Figures\Kopfdaten\{selection}\{i}.png", bbox_inches="tight")
            plt.close()       
        elif selection == "Anlage":
            DS=[]
            TH=[]
            ZS=[]
            DSList=['DSM', 'DSM.1', 'DSM.2', 'DSM.3', 'DSM.4', 'DSM.5', 'DSM.6', 'DSH', 'DSH.1', 'DSH.2', 'DSH.3', 'DSH.4', 'DSH.5', 'DSH.6', 'DSS', 'DSS.1', 'DSS.2', 'DSS.3', 'DSS.4', 'DSS.5', 'DSS.6']
            THList=['THM', 'THM.1', 'THM.2', 'THM.3', 'THM.4', 'THH', 'THH.1', 'THH.2', 'THH.3', 'THH.4', 'THS', 'THS.1', 'THS.2', 'THS.3', 'THS.4']
            ZSList=['ZSM', 'ZSM.1', 'ZSM.2', 'ZSM.3', 'ZSH', 'ZSH.1', 'ZSH.2', 'ZSH.3', 'ZSS', 'ZSS.1', 'ZSS.2', 'ZSS.3']
            for indM in DSList:
#                print(KP.loc[f"{ind}"])
                DS.append(float(KP.loc[f"{indM}"]))
            for indH in THList:
                TH.append(float(KP.loc[f"{indH}"]))
            for indS in ZSList:
                ZS.append(float(KP.loc[f"{indS}"]))

            KPDS=pd.DataFrame(DS)#, index=list(range(11,32)))
            KPTH=pd.DataFrame(TH)#, index=list(range(11,26)))
            KPZS=pd.DataFrame(ZS)#, index=list(range(11,23)))
#            print(KPTH)
            plt.figure()
            ax1=plt.subplot(1,3,1)
            plt.title(f"{i} - M, H & S ", loc="left")
            plt.boxplot(KPDS, labels="D", showmeans=True)
            ax2=plt.subplot(1,3,2, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPTH, labels="T", showmeans=True)
            ax3=plt.subplot(1,3,3, sharey=ax1)
            plt.tick_params("y", labelleft=False)
            plt.boxplot(KPZS, labels="Z", showmeans=True)
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

if __name__ == "__main__":
    Boot = str(input("Starten: type any string | Ende \n"))
    if Boot != "Ende":
        while Boot != "Ende":
            Status = 0
            Task = str(input("KopfdatenMW auswerten: KMW | Kopfdaten auswerten: K | Shinozaki-Kurven: SK | Ende \n"))
            if Task == "KMW":
                while Status != 2:
                    selection=str(input("DS | TH | ZS | Anlage | Ende \n"))
                    Status=KopfdatenMW(selection)
                    if Status == 1:
                        print(f"{selection} Done")
                    elif Status == 2:
                        print("KMW beendet")
            elif Task == "K":
                Status=0
                while Status != 2:
                    selection=str(input("DS | TH | ZS | Anlage | Ende \n"))
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
            elif Task == "Ende":
                Boot = str(input("Starten: type any string | Ende \n"))


    

    
    
