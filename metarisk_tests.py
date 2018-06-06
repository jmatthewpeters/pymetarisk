#
# MetaRisk tests
#

import metarisk
import pathlib

def main():
    testexcelfunction()

    
def test_createproject():
    lossCauseName = 'Pizza'
    projectname = 'test-project.xmr'
    ecm = metarisk.CapitalModel("./", projectname)
    newVariation = ecm.newproject()
    ecm.addlosscause("*", lossCauseName)
    ecm.saveproject()

def testexcelfunction():
    print(metarisk.getdatafromexcel(filePath=r'Y:\Python\MetaRisk\test_data.xlsx', sheetName='Sheet1', rangeName='A1:B2'))


if __name__ == "__main__":
    main()
