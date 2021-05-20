class ExcelTestStep(object):
    testsuitedec=""
    testcasecount=0
    def __init__(self):
        self.teststepcon=[]
        self.teststepexpect=[]
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
    def __settestcasenob__(self,testcasenob):
        self.testcasenob=testcasenob
    def __setteststepcon__(self,teststepcon):
        self.teststepcon.append(teststepcon)
    def __setteststepexpect__(self,teststepexpect):
        self.teststepexpect.append(teststepexpect)
    def __setteststepnum__(self,teststepnum):
        self.teststepnum=teststepnum
