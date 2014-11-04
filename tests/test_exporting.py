import unittest
import mock
from click.testing import CliRunner
import msword_cli
import os
from .util import MockApp

@mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx']))
class TestExportCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_export_defaults(self, mock_app):
        ''' Test export defaults. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, [filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_no_path(self, mock_app):
        ''' Test export with no path. '''
        result = self.runner.invoke(msword_cli.export)
        self.assertEqual(result.exit_code, 2)

    def test_export_dir_path(self, mock_app):
        ''' Test export with dir for path. '''
        mock_app.ActiveDocument.Name = 'foo.docx'
        result = self.runner.invoke(msword_cli.export, [os.getcwd()]) 
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath('foo.pdf'),              # <= Notable kwarg
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_no_ext(self, mock_app):
        ''' Test export with no file extension given. '''
        filename = 'foo'
        result = self.runner.invoke(msword_cli.export, [filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename + '.pdf'),      # <= Notable kwarg
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_xps(self, mock_app):
        ''' Test export to xps format. '''
        filename = 'foo'
        result = self.runner.invoke(msword_cli.export, ['--xps', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename + '.xps'),      # <= Notable kwarg
            ExportFormat=msword_cli.C.wdExportFormatXPS,            # <= Notable kwarg
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_for_screen(self, mock_app):
        ''' Test export for screens. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--for-screen', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForOnScreen,   # <= Notable kwarg
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_show(self, mock_app):
        ''' Test export show. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--show', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=True,                                   # <= Notable kwarg
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_range(self, mock_app):
        ''' Test export range of pages. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--pages', '2-3', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportFromTo,                      # <= Notable kwarg
            From=2,                                                 # <= Notable kwarg
            To=3,                                                   # <= Notable kwarg
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_bad_range(self, mock_app):
        ''' Test export range of pages. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--pages', '4-3', filename])
        self.assertEqual(result.exit_code, 2)

    def test_export_current_page(self, mock_app):
        ''' Test export current page. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--current-page', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportCurrentPage,                 # <= Notable kwarg
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_selection(self, mock_app):
        ''' Test export selection. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--selection', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportSelection,                   # <= Notable kwarg
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_with_markup(self, mock_app):
        ''' Test export with markup. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--with-markup', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentWithMarkup,           # <= Notable kwarg
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_with_properties(self, mock_app):
        ''' Test export with properties. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--with-props', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=True,                                   # <= Notable kwarg
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_without_irm(self, mock_app):
        ''' Test export without IRM. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--without-irm', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=False,                                          # <= Notable kwarg
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_with_heading_bookmarks(self, mock_app):
        ''' Test export with heading bookmarks. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--with-heading-bookmarks', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateHeadingBookmarks, # <= Notable kwarg
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_with_word_bookmarks(self, mock_app):
        ''' Test export with word bookmarks. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--with-word-bookmarks', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateWordBookmarks,   # <= Notable kwarg
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_without_structure_tags(self, mock_app):
        ''' Test export without structure tags. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--without-structure-tags', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=False,                                     # <= Notable kwarg
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

    def test_export_without_bitmaped_fonts(self, mock_app):
        ''' Test export without bitmaped fonts. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--without-bitmaped-fonts', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=False,                                   # <= Notable kwarg
            UseISO19005_1=False
        )

    def test_export_useiso19005_1(self, mock_app):
        ''' Test export useiso19005-1. '''
        filename = 'foo.pdf'
        result = self.runner.invoke(msword_cli.export, ['--useiso19005-1', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.ExportAsFixedFormat.assert_called_with(
            OutputFileName=os.path.abspath(filename),
            ExportFormat=msword_cli.C.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
            Range=msword_cli.C.wdExportAllDocument,
            Item=msword_cli.C.wdExportDocumentContent,
            IncludeDocProps=False,
            KeepIRM=True,
            CreateBookmarks=msword_cli.C.wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=True                                          # <= Notable kwarg
        )
