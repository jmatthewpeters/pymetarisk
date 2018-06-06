#----------------------------------------------------------------------------------------
# Python module wrapping access to .NET scripting for  MetaRisk.
#
#----------------------------------------------------------------------------------------

import clr
import sys
import os
import pathlib
import openpyxl

clr.AddReference("System, Version=2.0.5.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e, Retargetable=Yes")
clr.AddReference("System.Core, Version=2.0.5.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e, Retargetable=Yes")
clr.AddReference("GuyCarp.ServiceLayer")
clr.AddReference("GuyCarp.ServiceLayer.Caching")
clr.AddReference("XMR.Client.UserModel")
clr.AddReference("XMR.Client.Simulation")
clr.AddReference("XMR.Base")

from GuyCarp.ServiceLayer import *
from GuyCarp.ServiceLayer.Server.Management import *
from GuyCarp.ServiceLayer.Caching import *
from XMR.Client.UserModel import *
from XMR.Client.Simulation import *
from XMR.Base.API import *
from XMR.Base.Kernel.Runtime import *
from System import Array

def getdatafromexcel(filePath, sheetName, rangeName):

    normalizedFilePath = os.path.abspath(filePath)
    workbook = openpyxl.load_workbook(normalizedFilePath, False, True)
    ws = workbook[sheetName]
    inputData = ws[rangeName]

    retval = [[cell.value for cell in row] for row in inputData]
      
    workbook.close()

    return retval


class CapitalModel(object):
    def __init__(self, modelpath, filename, variations=[], losscauses=[], contracts=[], reserves=[], groups=[]):
        self.modelpath = modelpath
        self.filename = filename
        self.metariskmodel = UserModelRequestProcessor()
        self.variations = variations
        self.losscauses = losscauses
        self.contracts = contracts
        self.reserves = reserves
        self.modeldict = {}
        self.modeldict['modelpath'] = self.modelpath
        self.modeldict['filename'] = self.filename


    def newproject(self):
        '''creates a new projects
        '''
        self.metariskmodel.NewProject()
       

    def saveproject(self):
        '''saves the current project
        '''
        self.metariskmodel.SaveProject(self.filename)

    def openproject(self):
        '''opens an existing project
        this is not started when the object is initiated
        '''
        status = self.metariskmodel.OpenProject(self.modelpath / self.filename)
        return status
    
    def closeproject(self):
        '''closes the project
        '''
        pass

    def renameproject(self, newfilename):
        '''renames model file with new name 
        '''
        self.metariskmodel.SaveProject(newfilename)

    def addlosscause(self, variation, name, losstype):
        '''Add LossCause, Variation Name
        .Standard, .Bulk, .Attritional, .LossRatio, .Clash, .Assumed, .Reserve, .Tabular, .OEP
        '''
        data = []
        data.append(["Operation", "Component", "Variation", "Name"])
        data.append(["Add", "LossCause", variation, name])
        self.metariskmodel.ProcessUserAPI(listtoarray(data))
    
    def updatelosscause(self, name, variation,*kwargs):
        '''Operation, Component, Variation, Name, Minimum, Maximum, BetaMean
        Attritional: Distribution(Normal, Lognormal, Gamma), Mean(O or greater), CV(0 or greater)
        ScaleFactor (AnyValue)
        Clash: SeverityCorrelation (0 to 1), Claims(positive decimal values), RelativeProbability(0 to 1)
        Tabular: Time (0 or greater), Value (Any Value but not infinite)(can be multiplevalues)
        OEP: LossValue(>0), Probability(0 to 1) (multiple rows not blank)
        .PolicyProfile: Name, MinLimit, MaxLimit, Risks, CessionRate.1, CessionRate.2, SIR, MaxLoss, TotalValue, Participation, Premium
        AnnualWritten, ScaleFactor, UnearnedPremiumReserve, FixedExpense, AcquisitionExpense, OtherExpense, ULAE, InitialLag, NumberWritten, EarningTimeSpan
        LossCause.Inflation Variation Name LossInflation PremiumInflation
        '''
        pass

    def renamelosscause(self, name, newname, variation):
        '''renames loss cause
        '''
        pass

    def deletelosscause(self, name, variation):
        '''deletes loss cause
        '''
    
    def addcontract(self, variation, name):
        '''Add, Contract, Variation, Name
        Contract.Excess
        .ExcessWithReinstatements
        .ExcessWithRateOnPremium
        .ExcessWithSwingRatePremium
        .FHCF
        .PerRiskExcess
        .QuotaShare
        .SampleILW
        .StopLoss
        .StopLossRatio
        .SurplusShare
        .AggregateExcess
        .IndexedExcess
        .TabularQuotaShare
        '''
        pass

    def updatecontract(self, variation, name, *kwargs):
        '''variation, name, limit, retention
        StartTime, EndTime, ContractType (LossesOccurring or RisksAttaching), Currency
        ClauseName, 
        PerOccurrenceExcess (Limit, Retention, Currency, AllocateProportionally)
        BasePremium (Payments, InitialTimeLag, AnnualPremium, Currency)
        RateOnPremium(Payment, RateonSubjectPremium, DepositPremium, MinimumPremiumPct)
        '''
        pass
    
    def renamecontract(self, variation, name, newname):
        '''renames an existing contracts
        '''
        pass

    def deletecontract(self, variation, name):
        '''deletes a contract from the model
        '''
        pass
    
    def copycontract(self, name, variation, newvariation, newname=""):
        '''copies contract to new variation and optionally changes the name
        '''
        pass


    def addcoveragetocontract(self, contract, losscause):
        '''Add, Coverage, Variation, LossCause, Contract
        Coverage
        '''
        pass

    def addcomponenttogroup(self, name, variation, component, group):
        '''Operation, Component Variation, Name, Group
        Update, LossCause, Variation, Name, Group
        '''
        pass

    def addobjects(self, data):
        '''do we want to have a common process function to pass arrays to modelbuilder?
        we could collect all the operations together and apply them all at once.
        '''
        self.metariskmodel.ProcessUserAPI(listtoarray(data))

    def deleteobjects(self, data):
        '''general detail objects using modelbuilder
        '''
        pass

    def updateobjects(self, data):
        '''general modelbuider to udpate objects
        expects a data list
        '''
        pass

    def renameobjects(self, data):
        '''general method ot rename ojbcets using model builder
        '''
        pass

    def commitchanges(self, data):
        '''general commit changes 
        '''
        pass



def listtoarray(data):
    '''designed to take a 2d list and covert it to a 2d array
    '''
    retvalue = Array.CreateInstance(str, len(data), len(data[0]))
    for rowindex, row in enumerate(data):
        for colindex, listvalue in enumerate(row):
            retvalue[rowindex, colindex] = listvalue

    return retvalue