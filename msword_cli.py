from __future__ import unicode_literals
from win32com import client as com
from win32com.client import constants as C
from pywintypes import com_error
from collections import OrderedDict
import click
import os

VERSION = '0.1'

try:
    WORD = com.gencache.EnsureDispatch('Word.Application')
    TEMPLATE_DIR = WORD.Options.DefaultFilePath(C.wdUserTemplatesPath)
except com_error as e:
    raise click.ClickException(e.excepinfo[2])
except Exception as e:
    raise click.ClickException("Unable to load 'Word.Application'.")


class Template(click.Path):
    '''
    The Template type resolves relative paths first to the cwd and then
    to the template directory set in Word's File Options dialog. Upon
    resolving the realative path, it behaves as click.Path. Absolute
    paths receive no special treatment.
    '''
    def convert(self, value, param, ctx):
        if not os.path.isabs(value):
            # Resolve relative path
            if os.path.exists(os.path.abspath(value)):
                # Use existing file in cwd
                value = os.path.abspath(value)
            else:
                # Assume template dir
                value = os.path.join(TEMPLATE_DIR, value)
        # Pass on to click.Path for further validation
        return super(Template, self).convert(value, param, ctx)


def validate_range(ctx, param, value):
    '''
    Validate a simple range and return a tuple: (eg.'3-6' => (3, 6)).
    '''
    try:
        if value is not None:
            frm, to = map(int, value.split('-', 2))
            if frm > to:
                raise ValueError
            return (frm, to)
    except ValueError:
        raise click.BadParameter('Range must be in the format "x-y" where "x" and "y" are integers '
                                 'and "y" is greater than or equal to "x".')


def print_version(ctx, param, value):
    '''
    Click callback which prints version and exits.
    '''
    if not value or ctx.resilient_parsing:
        return
    click.echo('Version %s' % VERSION)
    ctx.exit()   


@click.group(chain=True)
@click.option('--version', is_flag=True, callback=print_version,
              expose_value=False, is_eager=True, 
              help='Print version info and exit.')
def cli():
    ''' 
    Command line interface for Microsoft Word. 
    
    Run 'msw <command> --help' to display help for a specific command. 
    '''
    pass


@cli.command('open')
@click.argument('path', type=click.Path(exists=True, resolve_path=True))
@click.option('--show/--hide', default=True, 
              help='Display or hide the document.')
def open(path, show):
    ''' 
    Open an existing document and activate. 
    '''
    click.echo('Opening document at "%s"' % path)
    try:
        Word.Documents.Open(FileName=path, Visible=show)
        if show and not WORD.Visible:
            # Only change state to visible if not visible
            # otherwise leave Word's visible state as-is
            WORD.Visible = show
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])


@cli.command('new')
@click.option('-t', '--template', type=Template(exists=True, dir_okay=False, resolve_path=True), 
              help='Use template at PATH.')
@click.option('--show/--hide', default=True, 
              help='Display or hide the document.')
def new(template, show):
    ''' 
    Create a new document and activate. 
    
    When the template path is relative, an attempt will be made to load 
    the template from the current working directory, and then from the 
    template directory set in Word's File Options dialog. If your 
    template file is in another location, you must specify an absolute 
    path.

    If no template path is provided, a new blank document will be created.
    '''
    try:
        if template:
            click.echo('Opening new document using template: "%s"' % template)
            doc = WORD.Documents.Add(template, Visible=show)
        else:
            click.echo('Opening new blank document.')
            doc = WORD.Documents.Add(Visible=show)
        if show and not WORD.Visible:
            # Only change state to visible if not visible
            # otherwise leave Word's visible state as-is
            WORD.Visible = show
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])


PRINT_OUT_ITEMS = OrderedDict([
    ('document_content',  C.wdPrintDocumentContent),
    ('doc_with_markup',   C.wdPrintDocumentWithMarkup),
    ('comments',          C.wdPrintComments),
    ('properties',        C.wdPrintProperties),
    ('markup',            C.wdPrintMarkup),
    ('styles',            C.wdPrintStyles),
    ('auto_text_entries', C.wdPrintAutoTextEntries),
    ('key_assignments',   C.wdPrintKeyAssignments),
    ('envelope',          C.wdPrintEnvelope)
])

@cli.command('print')
@click.option('-c', '--copies', default=1, help='The number of copies to be printed.')
@click.option('-p', '--pages', type=str, help='The page numbers and page ranges '
              'to be printed, separated by commas. For example, "2, 6-10" prints '
              'page 2 and pages 6 through 10. Ignored if \'--current-page\' or '
              '\'--selection\' is specified.')
@click.option('--even', 'pagetype', flag_value=C.wdPrintEvenPagesOnly,
              help='Print even-numbered pages only.')
@click.option('--odd', 'pagetype', flag_value=C.wdPrintOddPagesOnly, 
              help='Print odd-numbered pages only.')
@click.option('--current-page', 'range', flag_value=C.wdPrintCurrentPage, 
              help='Print the current page only.')
@click.option('--selection', 'range', flag_value=C.wdPrintSelection, 
              help='Print the current selection.')
@click.option('--no-collate', is_flag=True, help='Do not collate multiple copies.')
@click.option('--to-file', type=click.Path(dir_okay=False, resolve_path=True), 
              help='Print document to file at PATH.')
@click.option('--append', is_flag=True, help='Append the document to the file specified '
              'by the \'--to-file\' option rather than overwriting it.')
@click.option('--columns', type=click.Choice(['1', '2', '3', '4']), default='1',
              help='The number of pages to fit horizontally on one page. '
              'Use with the \'--rows\' option to print multiple pages on a single sheet.')
@click.option('--rows', type=click.Choice(['1', '2', '4']), default='1',
              help='The number of pages to fit vertically on one page. '
              'Use with the \'--columns\' argument to print multiple pages on a single sheet.')
@click.option('--item', type=click.Choice(PRINT_OUT_ITEMS.keys()), default='document_content', 
              help='The item to be printed. Defaults to \'document_content\'.')
def prnt(copies, pages, pagetype, range, item, no_collate, to_file, append, columns, rows):
    ''' 
    Print active document to default printer. 

    The options '--pages', '--current-page' and '--selection' are mutualy exclusive.
    The '--pages' option is ignored if '--current-page' or '--selection' is specified.
    The last option specified  of '--current-page' or '--selection' will be honored.
    If none of these options are specified, the entire document is exported.

    The options '--even' and '--odd' are mutualy exclusive. Only the last one specified 
    will be honored.  If neither is specified, both even and odd pages will be printed.
    '''
    click.echo('Printing %s copies of pages: %s' % (copies, pages or 'all'))
    options = {
        'Background':       True,
        'Copies':           copies,
        'Collate':          not no_collate,
        'Item':             PRINT_OUT_ITEMS[item],
        'PrintToFile':      False,
        'PageType':         C.wdPrintAllPages,
        'Range':            C.wdPrintAllDocument,
        'PrintZoomColumn':  int(columns),
        'PrintZoomRow':     int(rows)
    }
    if range:
        options['Range'] = range
    elif pages:
        options['Pages'] = pages
        options['Range'] = C.wdPrintRangeOfPages

    if pagetype:
        options['PageType'] = pagetype

    if to_file:
        options['PrintToFile'] = True
        options['OutputFileName'] = to_file
        if append:
            options['Append'] = True
    
    try:
        WORD.ActiveDocument.PrintOut(**options)
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])


@cli.command('export')
@click.option('--pdf', 'format', flag_value=C.wdExportFormatPDF, default=True,
              help="Export document into PDF format. The default.")
@click.option('--xps', 'format', flag_value=C.wdExportFormatXPS,
              help='Export document into XML Paper Specification (XPS) format.')
@click.option('--show', is_flag=True, help='Open the new file in the appropriate viewer after exporting.')
@click.option('--for-print', 'optimize', flag_value=C.wdExportOptimizeForPrint, default=True,
              help='Export for print, which is higher quailty and results in a larger file size. '
              'The default.')
@click.option('--for-screen', 'optimize', flag_value=C.wdExportOptimizeForOnScreen,
              help='Export for screen, which is a lower quality and results in a smaller file size.')
@click.option('--pages', type=str, callback=validate_range, 
              help='The range of pages to export. For example, "3-6" exports '
              'pages 3 through 6 and "2-2" exports page 2 only. Ignored if '
              '\'--current-page\' or \'--selection\' is specified.')
@click.option('--current-page', 'range', flag_value=C.wdExportCurrentPage, 
              help='Export the current page only.')
@click.option('--selection', 'range', flag_value=C.wdExportSelection, 
              help='Export the current selection.')
@click.option('--with-markup', 'markup', is_flag=True, help='Export the document with markup.')
@click.option('--with-props', 'properties', is_flag=True, 
              help='Include document properties in the newly exported file.')
@click.option('--without-irm', 'irm', is_flag=True, help='Exclude IRM permissions to an XPS document '
              'if the source document has IRM protections.')
@click.option('--with-heading-bookmarks', 'bookmarks', flag_value=C.wdExportCreateHeadingBookmarks,
              help='Create a bookmark in the exported document for each Microsoft Word heading, which '
              'includes only headings within the main document and text boxes not within headers, '
              'footers, endnotes, footnotes, or comments.')
@click.option('--with-word-bookmarks', 'bookmarks', flag_value=C.wdExportCreateWordBookmarks,
              help='Create a bookmark in the exported document for each Word bookmark, which includes '
              'all bookmarks except those contained within headers and footers.')
@click.option('--without-structure-tags', 'struct', is_flag=True,
              help='Exclude extra data to help screen readers, for example information about the flow '
              'and logical organization of the content.')
@click.option('--without-bitmaped-fonts', 'bitmap', is_flag=True, help='Exclude a bitmap of the text. '
              'The viewer\'s computer substitutes an appropriate font if the authored one is not '
              'available. Warning: always inlcude a bitmap when font licenses do not permit a font to '
              'be embedded in the PDF file.')
@click.option('--useiso19005-1', is_flag=True, help='Limit PDF usage to the PDF subset standardized '
              'as ISO 19005-1. If used, the resulting files are more reliably self-contained but may '
              'be larger or show more visual artifacts due to the restrictions of the format.') 
@click.argument('path', type=click.Path(dir_okay=True, resolve_path=True))
def export(path, format, show, optimize, pages, range, markup, properties, irm, bookmarks, struct, bitmap, useiso19005_1):
    '''
    Save active document as PDF or XPS format to PATH.

    If PATH points to a directory rather than a file, a new file will be created
    in that directory using the name of the document. For example:

        somedoc.docx => somedoc.pdf

    A relative path will resolve to the current working directory. You may want to
    use an absolute path to ensure the document is saved to the correct location.

    The options '--pdf' and '--xps' are mutualy exclusive. Only the last one specified 
    will be honored.  The default is '--pdf', which will be assumed if neither is specified.

    The options '--pages', '--current-page' and '--selection' are mutualy exclusive.
    The '--pages' option is ignored if '--current-page' or '--selection' is specified.
    The last option specified  of '--current-page' or '--selection' will be honored.
    If none of these options are specified, the entire document is exported.

    The options '--with-heading-bookmarks' and '--with-word-bookmarks' are mutualy exclusive.
    Only the last one specified will be honored.  If neither is specified, no bookmarks are exported.
    '''  
    try:
        if os.path.isdir(path):
            # No filename specified. Used document name
            name = os.path.splitext(WORD.ActiveDocument.Name)[0]
            path = os.path.join(path, name)
        if os.path.splitext(path)[1].lower() not in ['pdf', 'xps']:
            # No file extension specified. Use file format.
            if format == C.wdExportFormatPDF:
                ext = '.pdf'
            else:
                ext = '.xps'
            path += ext
        options = {
                'OutputFileName':     path,
                'ExportFormat':       format,
                'OpenAfterExport':    show,
                'OptimizeFor':        optimize,
                'Range':              range if range else C.wdExportFromTo if pages else C.wdExportAllDocument,
                'Item':               C.wdExportDocumentWithMarkup if markup else C.wdExportDocumentContent,
                'IncludeDocProps':    properties,
                'KeepIRM':            not irm,
                'CreateBookmarks':    bookmarks or C.wdExportCreateNoBookmarks,
                'DocStructureTags':   not struct,
                'BitmapMissingFonts': not bitmap,
                'UseISO19005_1':      useiso19005_1
        }

        if pages:
            options['From'] = pages[0]
            options['To'] = pages[1]

        WORD.ActiveDocument.ExportAsFixedFormat(**options)
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])

@cli.command('save')
@click.option('-a', '--all', is_flag=True,
              help='Save all open documents.')
@click.option('-f', '--force', is_flag=True,
              help='Do not prompt to save changes.')
@click.option('-p', '--path', type=click.Path(resolve_path=True),
              help='Save document to PATH.')
def save(all, force, path):
    ''' 
    Save document(s). 
    
    If a path is given, the active document is saved
    with a new name at PATH. The --all and --force 
    options are ignored when a path is provided.
    
    If no path is given, the active document is saved 
    to its current path.
    '''
    click.echo('save to "%s"' % path)
    try:
        if path:
            click.echo('Saving document to: "%s"' % path)
            WORD.ActiveDocument.SaveAs(path)
        else:
            if all:
                doc = WORD.Documents
            else:
                doc = WORD.ActiveDocument
            click.echo('Saving changes to existing document.')
            doc.Save(NoPrompt=force)
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])


@cli.command('close')
@click.option('-a', '--all', is_flag=True,
              help='Close all open documents.')
@click.option('-f', '--force', is_flag=True,
              help='Force close without saving changes.')
def close(all, force):
    ''' 
    Close document(s). 
    
    Unless the --force option is used, Word will prompt
    to save any changes.
    
    Will only quit Word if no other documents are open.
    '''
    try:
        if all:
            doc = WORD.Documents
        else:
            doc = WORD.ActiveDocument
        if force:
            click.echo('Force closing document...')
            doc.Close(C.wdDoNotSaveChanges)
        else:
            click.echo('Closing document...')
            doc.Close(C.wdPromptToSaveChanges)

        if not WORD.Documents.Count:
            # Only quit if no other documents are open
            WORD.Quit()
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])
    

@cli.command('activate')
@click.argument('index', type=int)
def activate(index):
    ''' 
    Activate a document. 
    '''
    click.echo('Activate document at index "%s"' % index)
    try:
        WORD.Documents.Item(index).Activate()
    except com_error as e:
        raise click.ClickException(e.excepinfo[2])


@cli.command('docs') # Or should it be 'list' or 'ls'???
def docs():
    ''' 
    List open documents. 
    '''
    if WORD.Documents.Count:
        click.echo('\nOpen Documents:\n')
        template = ' {{}} [{{: ={}}}] {{}}{{}}'.format(len(str(WORD.Documents.Count)))
        for i, doc in enumerate(WORD.Documents, start=1):
            active = '*' if WORD.ActiveDocument == doc else ' '
            saved = '*' if not doc.Saved else ''
            click.echo(template.format(active, i, doc.Name, saved))
    else:
        click.echo('\nNo open documents found.')
    
if __name__ == '__main__':
    cli()