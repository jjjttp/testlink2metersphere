class ExcelTestCase(object):
    testsuitedec=""
    testcasecount=0
    def __settestsuitename__(self,testsuitename):
        self.testsuitename=testsuitename
    def __gettestsuitename__(self):
        return self.testsuitename
    def __settestcasename__(self,testcasename):
        self.testcasename=testcasename
    def __gettestcasename__(self):
        return self.testcasename
    def __settestsuitedec__(self,testsuitedec):
        self.testsuitedec=testsuitedec
    def __gettestsuitedec__(self):
        return self.testsuitedec
    def __settestcasecount__(self,testcasecount):
        self.testcasecount=testcasecount
    def __settestcasenob__(self,testcasenob):
        self.testcasenob=testcasenob
    def __settestcasesummary__(self,testcasesummary):
        self.testcasesummary=testcasesummary
    def __settestcaseprecon__(self,testcaseprecon):
        self.testcaseprecon=testcaseprecon
    def __settestcasestatus__(self,testcasestatus):
        self.testcasestatus=testcasestatus
    def __settestcasetype__(self,testcasetype):
        self.testcasetype=testcasetype
    def __setteststepdic__(self,teststepdic):
        self.teststepdic=teststepdic