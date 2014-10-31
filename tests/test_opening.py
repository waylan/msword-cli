import unittest
import mock
from click.testing import CliRunner
import msword_cli
from .util import MockApp, touch
import os


@mock.patch('msword_cli.WORD', spec_set=MockApp([]))
class TestOpenCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_open_defaults(self, mock_app):
        ''' Test open defaults. '''
        filename = 'foo.docx'
        with self.runner.isolated_filesystem():
            touch(filename)
            result = self.runner.invoke(msword_cli.open, [filename])
            self.assertEqual(result.exit_code, 0)
            mock_app.Documents.Open.assert_called_with(FileName=os.path.abspath(filename), Visible=True)

    def test_open_hide(self, mock_app):
        ''' Test open hide. '''
        filename = 'foo.docx'
        with self.runner.isolated_filesystem():
            touch(filename)
            result = self.runner.invoke(msword_cli.open, ['--hide', filename])
            self.assertEqual(result.exit_code, 0)
            mock_app.Documents.Open.assert_called_with(FileName=os.path.abspath(filename), Visible=False)


@mock.patch('msword_cli.WORD', spec_set=MockApp([]))
class TestNewCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_new_defaults(self, mock_app):
        ''' Test new defaults. '''
        result = self.runner.invoke(msword_cli.new)
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Add.assert_called_with(Visible=True)

    def test_new_hide(self, mock_app):
        ''' Test new hide. '''
        result = self.runner.invoke(msword_cli.new, ['--hide'])
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Add.assert_called_with(Visible=False)

    def test_new_template_cwd(self, mock_app):
        ''' Test new with template from cwd. '''
        filename = 'foo.dot'
        with self.runner.isolated_filesystem():
            touch(filename)
            result = self.runner.invoke(msword_cli.new, ['--template', filename])
            self.assertEqual(result.exit_code, 0)
            mock_app.Documents.Add.assert_called_with(os.path.abspath(filename), Visible=True)

    def test_new_template_abspath(self, mock_app):
        ''' Test new with template from absolute path. '''
        filename = 'foo.dot'
        with self.runner.isolated_filesystem():
            touch(filename)
            absname = os.path.abspath(filename)
            result = self.runner.invoke(msword_cli.new, ['--template', absname])
            self.assertEqual(result.exit_code, 0)
            mock_app.Documents.Add.assert_called_with(absname, Visible=True)

    def test_new_template_default_path(self, mock_app):
        ''' Test new with template from directory set in Word's File Options dialog. '''
        filename = 'normal.dotm'
        result = self.runner.invoke(msword_cli.new, ['--template', filename])
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Add.assert_called_with(os.path.join(msword_cli.TEMPLATE_DIR, filename), Visible=True)
