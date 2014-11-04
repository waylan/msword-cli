import os

# =========================================================
# Dummy Mock objects to patch the com objects for testing
# =========================================================

class MockDoc(object):
    ''' A Document object. '''
    def __init__(self, Name):
        self.Name = Name

    def PrintOut(self, **kwargs):
        pass

    def Save(self, **kwargs):
        pass

    def SaveAs(self, **kwargs):
        pass

    def ExportAsFixedFormat(self, **kwargs):
        pass

class MockDocs(list):
    ''' A Document Collection. '''
    @property
    def Count(self):
        return len(self)

    def Item(self, index):
        # 1 based index - not 0. 
        return self[index-1]

    def Save(self, **kwargs):
        pass

    def Open(self, **kwargs):
        pass

    def Add(self, **kwargs):
        pass


class MockApp(object):
    ''' An Application object. '''
    def __init__(self, docs):
        ''' 
        Build App Object

        Given a list of document names (as strings)
        build a mock Application object of those docs
        '''
        self.Documents = MockDocs()
        for doc in docs:
            self.Documents.append(MockDoc(doc))
        if self.Documents.Count:
            self.ActiveDocument = self.Documents[-1]

    def Quit(self):
        pass

    @property
    def Visible(self):
        return True

# =========================================================
# Utility functions
# =========================================================

def touch(filename, times=None):
    ''' Simulate a file touch operation. '''
    with open(filename, 'a'):
        os.utime(filename, times)