import unittest
import mock
from click.testing import CliRunner
import msword_cli
import os
from .util import MockApp

@mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx']))
class TestPrintCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_print_defaults(self, mock_app):
        ''' Test print defaults. '''
        result = self.runner.invoke(msword_cli.prnt)
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_multiple_copies(self, mock_app):
        ''' Test printing multiple copies. '''
        result = self.runner.invoke(msword_cli.prnt, ['--copies', '3'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=3,                                   # <= The notable kwarg
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_multiple_copies(self, mock_app):
        ''' Test printing no colloate. '''
        result = self.runner.invoke(msword_cli.prnt, ['--no-collate'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=False,                              # <= The notable kwarg
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_doc_with_markup(self, mock_app):
        ''' Test printing doc with markup. '''
        result = self.runner.invoke(msword_cli.prnt, ['--item', 'doc_with_markup'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentWithMarkup, # <= The notable kwarg
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_to_file(self, mock_app):
        ''' Test printing to file. '''
        filename = 'foo'
        result = self.runner.invoke(msword_cli.prnt, ['--to-file', filename, '--append'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=True,                           # <= The notable kwarg
            OutputFileName=os.path.abspath(filename),   # <= The notable kwarg
            Append=True,                                # <= The notable kwarg
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_odd_pages(self, mock_app):
        ''' Test printing odd pages. '''
        result = self.runner.invoke(msword_cli.prnt, ['--odd'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintOddPagesOnly,  # <= The notable kwarg
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_current_page(self, mock_app):
        ''' Test print current page. '''
        result = self.runner.invoke(msword_cli.prnt, ['--current-page'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintCurrentPage,   # <= The notable kwarg,
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_selection(self, mock_app):
        ''' Test print selection. '''
        result = self.runner.invoke(msword_cli.prnt, ['--selection'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintSelection,    # <= The notable kwarg
            PrintZoomColumn=1,
            PrintZoomRow=1
        )

    def test_print_page_range(self, mock_app):
        ''' Test print range of pages. '''
        range = '2-3, 6'
        result = self.runner.invoke(msword_cli.prnt, ['--pages', range])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintRangeOfPages, # <= The notable kwarg
            Pages=range,                            # <= The notable kwarg
            PrintZoomColumn=1,
            PrintZoomRow=1
        )


    def test_print_zoom(self, mock_app):
        ''' Test print multiple pages on one page. '''
        result = self.runner.invoke(msword_cli.prnt, ['--columns', '2', '--rows', '4'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.PrintOut.assert_called_with(
            Background=True, 
            Copies=1,
            Collate=True,
            Item=msword_cli.C.wdPrintDocumentContent,
            PrintToFile=False,
            PageType=msword_cli.C.wdPrintAllPages,
            Range=msword_cli.C.wdPrintAllDocument,
            PrintZoomColumn=2,                  # <= The notable kwarg
            PrintZoomRow=4                      # <= The notable kwarg
        )