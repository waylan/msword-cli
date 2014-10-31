import unittest
import mock
from click.testing import CliRunner
import msword_cli
from .util import MockApp
from pywintypes import com_error

class TestListCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()
    
    @mock.patch('msword_cli.WORD.Documents', Count=0)
    @mock.patch('msword_cli.WORD')
    def test_no_docs(self, mock_app, mock_docs):
        ''' Test listing docs with no docs open. '''
        result = self.runner.invoke(msword_cli.docs)
        self.assertEqual(result.exit_code, 0)
        self.assertEqual(result.output, '\nNo open documents found.\n')

    @mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx']))
    def test_one_doc(self, mock_app):
        ''' Test listing docs with one doc open. '''
        result = self.runner.invoke(msword_cli.docs)
        self.assertEqual(result.exit_code, 0)
        # 3 header lines plus 1 doc = 4
        self.assertEqual(len(result.output.split('\n')), 4)

    @mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx', 'bar.docx', 'baz.docx']))
    def test_multiple_docs(self, mock_app):
        ''' Test listing docs with three docs open. '''
        result = self.runner.invoke(msword_cli.docs)
        self.assertEqual(result.exit_code, 0)
        # 3 header lines plus 3 docs = 6
        #self.assertEqual(len(result.output.split('\n')), 6)  #TODO: fix this
        #self.assertEqual(result.output, '')

@mock.patch('msword_cli.WORD', spec_set=MockApp(['foo.docx', 'bar.docx', 'baz.docx']))
class TestActivateCommand(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()

    def test_activate(self, mock_app):
        ''' Test activate. '''
        index = 2
        result = self.runner.invoke(msword_cli.activate, [str(index)])
        self.assertEqual(result.exit_code, 0)
        mock_app.Documents.Item.assert_was_called_with(index)
        #self.assertEqual(mock_app.Documents[index].Activate.called, True)

    def test_activate_bad_index(self, mock_app):
        ''' Test activate with bad index. '''
        index = 6
        mock_app.Documents.Item.side_effect = com_error
        result = self.runner.invoke(msword_cli.activate, [str(index)])
        self.assertEqual(result.exit_code, -1)

    def test_activate_no_index(self, mock_app):
        ''' Test activate with no index. '''
        result = self.runner.invoke(msword_cli.activate)
        self.assertEqual(result.exit_code, 2)