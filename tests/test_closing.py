import unittest
import mock
from click.testing import CliRunner
import msword_cli
from .util import MockApp

@mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx']))
class TestCloseCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_close_defaults(self, mock_app):
        ''' Test close defaults. '''
        result = self.runner.invoke(msword_cli.close)
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.Close.assert_called_with(msword_cli.C.wdPromptToSaveChanges)

    def test_force_close(self, mock_app):
        ''' Test force close. '''
        result = self.runner.invoke(msword_cli.close, ['--force'])
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.Close.assert_called_with(msword_cli.C.wdDoNotSaveChanges)

    def test_close_all(self, mock_app):
        ''' Test close all. '''
        result = self.runner.invoke(msword_cli.close, ['--all'])
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Close.assert_called_with(msword_cli.C.wdPromptToSaveChanges)

    def test_force_close_all(self, mock_app):
        ''' Test force close all. '''
        result = self.runner.invoke(msword_cli.close, ['--all', '--force'])
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Close.assert_called_with(msword_cli.C.wdDoNotSaveChanges)

    def test_close_no_quit(self, mock_app):
        ''' Test close but don't quit. '''
        mock_app.Documents.Count = 2
        result = self.runner.invoke(msword_cli.close)
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.Close.assert_called_with(msword_cli.C.wdPromptToSaveChanges)
        self.assertEqual(mock_app.Quit.called, False)

    def test_close_quit(self, mock_app):
        ''' Test close and quit. '''
        mock_app.Documents.Count = 0
        result = self.runner.invoke(msword_cli.close)
        self.assertEqual(result.exit_code, 0)
        mock_app.ActiveDocument.Close.assert_called_with(msword_cli.C.wdPromptToSaveChanges)
        self.assertEqual(mock_app.Quit.called, True)