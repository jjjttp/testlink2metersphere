class ExcelTestSuite(object):
    testsuitedec=""
    testcasecount=0
    def __settestsuitename__(self,testsuitename):
        self.testsuitename=testsuitename
    def __gettestsuitename__(self):
        return self.testsuitename
    def __settestsuitedec__(self,testsuitedec):
        self.testsuitedec=testsuitedec
    def __gettestsuitedec__(self):
        return self.testsuitedec
    def __settestcasecount__(self,testcasecount):
        self.testcasecount=testcasecount