import logging
import io
import itertools as it
import zipfile

import lxml
import openpyxl as ox
import path_helpers as ph

from ._version import get_versions
__version__ = get_versions()['version']
del get_versions

logger = logging.getLogger(__name__)


EXCEL_NAMESPACES = {k: getattr(ox.xml.constants, k)
                    for k in ('SHEET_MAIN_NS', 'REL_NS', 'PKG_REL_NS')}


def load_extension_lists(xlsx_path):
    '''
    Load extension list for each worksheet in an Excel spreadsheet.

    Extension lists include, e.g., worksheet data validation settings.

    Note that ``openpyxl`` does not currently `support reading existing data
    validation
    <http://openpyxl.readthedocs.io/en/default/validation.html#validating-cells>`_.

    As a workaround, this function makes it possible to load the extension list
    for each worksheet in a workbook so they may be restored to a workbook
    modified by ``openpyxl`` using  :func:`update_extension_lists`.

    .. versionadded:: 0.2

    See also
    --------
    :func:`update_extension_lists` :func:`update_data_validations`,
    :func:`load_extension_lists`,

    Parameters
    ----------
    xlsx_path : str
        Path to Excel ``xlsx`` file.

    Returns
    -------
    dict
        Mapping from each worksheet filepath in Excel ZIP file to
        corresponding extension list XML element (or ``None`` if the
        worksheet does not contain any extension list element).
    '''
    extension_lists = {}

    # Open Excel file.
    with zipfile.ZipFile(xlsx_path, mode='r') as input_:
        # Get mapping from each worksheet filename to corresponding `ZipInfo`
        # object.
        zip_info_by_filenames = {ph.path(v.filename): v
                                 for v in input_.filelist}

        # Extract extension list XML element (if present) from each worksheet.
        for filename_i, zip_info_i in zip_info_by_filenames.iteritems():
            if filename_i.parent != 'xl/worksheets':
                continue
            with io.BytesIO(input_.read(filename_i)) as data_i:
                worksheet_root_i = lxml.etree.parse(data_i)
            extension_lists_i = (worksheet_root_i
                                 .xpath('/SHEET_MAIN_NS:worksheet'
                                        '/SHEET_MAIN_NS:extLst',
                                        namespaces=EXCEL_NAMESPACES))
            if extension_lists_i:
                # This worksheet has an extension list.
                extension_lists[filename_i] = extension_lists_i[0]
            else:
                extension_lists[filename_i] = None
    return extension_lists


def update_extension_lists(xlsx_path, extension_lists):
    '''
    Update extension list for worksheets in an Excel spreadsheet.

    Extension lists include, e.g., worksheet data validation settings.

    Note that ``openpyxl`` does not currently `support reading existing data
    validation
    <http://openpyxl.readthedocs.io/en/default/validation.html#validating-cells>`_.

    As a workaround, this function makes it possible to restore extension
    lists saved using :func:`load_extension_lists` after modifying the
    workbook with ``openpyxl``.

    .. versionadded:: 0.2

    See also
    --------
    :func:`load_extension_lists` :func:`update_data_validations`,
    :func:`load_extension_lists`,

    Parameters
    ----------
    xlsx_path : str
        Path to Excel ``xlsx`` file.
    extension_lists : dict
        Mapping from each worksheet filepath in Excel ZIP file to
        corresponding extension list XML element.

    Returns
    -------
    bytes
        Modified Excel ``.xlsx`` file contents as a bytes string.
    '''
    with io.BytesIO() as output:
        with zipfile.ZipFile(output, mode='a',
                             compression=zipfile.ZIP_DEFLATED) as output_zip:
            # - Read existing file
            # - Append extension list from template file to worksheet XML.
            # - Copy all files except for `worksheet1` to in-memory zip file.
            with zipfile.ZipFile(xlsx_path, mode='r') as input_:
                zip_infos_by_filename = {ph.path(v.filename): v
                                         for v in input_.filelist}
                for filename_i, zip_info_i in zip_infos_by_filename.iteritems():
                    extension_list_i = extension_lists.get(filename_i)

                    if extension_list_i is None or (filename_i not in
                                                    extension_lists):
                        # Worksheet file has no extension list.  Use original
                        # worksheet contents.
                        contents_i = input_.read(filename_i)
                    else:
                        # Worksheet file has **extension list**.
                        # Load worksheet file XML contents from `xlsx_path`
                        # file.
                        with io.BytesIO(input_.read(filename_i)) as data:
                            root_i = lxml.etree.parse(data)
                        # Get root worksheet XML element.
                        worksheet_i = \
                            root_i.xpath('/SHEET_MAIN_NS:worksheet',
                                         namespaces=EXCEL_NAMESPACES)[0]
                        # Append the extension list to the worksheet element.
                        worksheet_i.append(extension_list_i)
                        # Use modified worksheet contents with extension list
                        # added.
                        contents_i = lxml.etree.tostring(root_i)
                    # Write worksheet contents to output zip.
                    output_zip.writestr(filename_i, contents_i,
                                        zip_info_i.compress_type)
            output_zip.close()
        return output.getvalue()


def load_data_validations(xlsx_path):
    '''
    Load data validations element for each worksheet in an Excel spreadsheet.

    ``openpyxl`` does not currently `support reading existing data
    validation
    <http://openpyxl.readthedocs.io/en/default/validation.html#validating-cells>`_.

    As a workaround, this function makes it possible to load the data
    validations element for each worksheet in a workbook so they may be
    restored to a workbook modified by ``openpyxl`` using
    :func:`update_data_validations`.

    .. versionadded:: 0.4

    See also
    --------
    :func:`update_data_validations`, :func:`load_extension_lists`,
    :func:`update_extension_lists`

    Parameters
    ----------
    xlsx_path : str
        Path to Excel ``xlsx`` file.

    Returns
    -------
    dict
        Mapping from each worksheet filepath in Excel ZIP file to
        corresponding ``dataValidations`` XML element (or ``None`` if the
        worksheet does not contain any data validations element).
    '''
    data_validations = {}

    # Open Excel file.
    with zipfile.ZipFile(xlsx_path, mode='r') as input_:
        # Get mapping from each worksheet filename to corresponding `ZipInfo`
        # object.
        zip_info_by_filenames = {ph.path(v.filename): v
                                 for v in input_.filelist}

        # Extract data validations XML element (if present) from each
        # worksheet.
        for filename_i, zip_info_i in zip_info_by_filenames.iteritems():
            if filename_i.parent != 'xl/worksheets':
                continue
            with io.BytesIO(input_.read(filename_i)) as data_i:
                worksheet_root_i = lxml.etree.parse(data_i)
            data_validations_i = (worksheet_root_i
                                  .xpath('/SHEET_MAIN_NS:worksheet'
                                         '/SHEET_MAIN_NS:dataValidations',
                                         namespaces=EXCEL_NAMESPACES))
            if data_validations_i:
                # This worksheet has a data validations element.
                data_validations[filename_i] = data_validations_i[0]
            else:
                data_validations[filename_i] = None
    return data_validations


def update_data_validations(xlsx_path, data_validations):
    '''
    Update data validations for worksheets in an Excel spreadsheet.

    ``openpyxl`` does not currently `support reading existing data
    validation
    <http://openpyxl.readthedocs.io/en/default/validation.html#validating-cells>`_.

    As a workaround, this function makes it possible to restore a data
    validations element saved using :func:`load_data_validations` after
    modifying the workbook with ``openpyxl``.

    .. versionadded:: 0.2

    See also
    --------
    :func:`load_data_validations`, :func:`load_extension_lists`,
    :func:`update_extension_lists`

    Parameters
    ----------
    xlsx_path : str
        Path to Excel ``xlsx`` file.
    data_validations : dict
        Mapping from each worksheet filepath in Excel ZIP file to
        corresponding data validations XML element.

    Returns
    -------
    bytes
        Modified Excel ``.xlsx`` file contents as a bytes string.
    '''
    with io.BytesIO() as output:
        with zipfile.ZipFile(output, mode='a',
                             compression=zipfile.ZIP_DEFLATED) as output_zip:
            # - Read existing file
            # - Append data validations element from `data_validations` to
            #   worksheet XML.
            # - Copy all files except for `worksheet1` to in-memory zip file.
            with zipfile.ZipFile(xlsx_path, mode='r') as input_:
                zip_infos_by_filename = {ph.path(v.filename): v
                                         for v in input_.filelist}
                for filename_i, zip_info_i in zip_infos_by_filename.iteritems():
                    data_validations_i = data_validations.get(filename_i)

                    if data_validations_i is None or (filename_i not in
                                                      data_validations):
                        # Worksheet file has no data validations element.  Use
                        # original worksheet contents.
                        contents_i = input_.read(filename_i)
                    else:
                        # Worksheet file has **data validations element**.
                        # Load worksheet file XML contents from `xlsx_path`
                        # file.
                        with io.BytesIO(input_.read(filename_i)) as data:
                            root_i = lxml.etree.parse(data)
                        # Get root worksheet XML element.
                        worksheet_i = \
                            root_i.xpath('/SHEET_MAIN_NS:worksheet',
                                         namespaces=EXCEL_NAMESPACES)[0]

                        existing_validations_i = \
                            worksheet_i.xpath('//SHEET_MAIN_NS:dataValidations',
                                              namespaces=EXCEL_NAMESPACES)

                        if existing_validations_i:
                            logger.debug('Replace existing data validation(s)')
                            worksheet_i.replace(existing_validations_i[0],
                                                data_validations_i)
                        else:
                            logger.debug('Append new data validation(s)')
                            # Append the data validations element to the
                            # worksheet element.
                            worksheet_i.append(data_validations_i)
                        # Use modified worksheet contents with data validations
                        # element added.
                        contents_i = lxml.etree.tostring(root_i)
                    # Write worksheet contents to output zip.
                    output_zip.writestr(filename_i, contents_i,
                                        zip_info_i.compress_type)
            output_zip.close()
        return output.getvalue()


def get_column_widths(worksheet, min_width=None):
    '''
    .. versionadded:: 0.2

    Parameters
    ----------
    worksheet : openpyxl.worksheet.worksheet.Worksheet
        Excel worksheet.
    min_width : int, optional
        Minimum column width in characters.

    Returns
    -------
    dict
        Mapping from letter of each column containing at least one non-blank
        cell to the corresponding column width to fit the widest entry in the
        column.
    '''
    def column_key(x):
        return x.column
    column_widths = {column_i: max(max(len(str(cell_ij.value))
                                       for cell_ij in cells_i),
                                   min_width or 0)
                     for column_i, cells_i in
                     it.groupby(sorted(worksheet.get_cell_collection(),
                                       key=column_key), key=column_key)}
    return column_widths


def get_defined_names_by_worksheet(workbook):
    '''
    .. versionadded:: 0.3

    Parameters
    ----------
    workbook : openpyxl.workbook.workbook.Workbook

    Returns
    -------
    dict
        Mapping from each worksheet name to the corresponding defined names
        (i.e., named ranges) in the worksheet.

        Each value in the top-level dictionary corresponds to a dictionary
        mapping each defined name to the corresponding range.

        For example:

            {'Foo sheet': {'Some foo range': '$D$11:$D$1048576',
                           'Some foo cell': '$B$6'},
             'Bar sheet': {'Some bar range': '$I$2:$I$3',
                           'Some bar cell': '$K$2'}}
    '''
    defined_name_tuples = \
        sorted([tuple(it.chain(*[(sheet_name_i, defined_name_i.name, range_i)
                                 for sheet_name_i, range_i in
                                 defined_name_i.destinations]))
                for defined_name_i in workbook.defined_names.definedName])

    return dict([(sheet_name_i,
                  dict([tuple_ij[1:] for tuple_ij in defined_names_group_i]))
                 for sheet_name_i, defined_names_group_i in
                 it.groupby(defined_name_tuples, lambda n: n[0])])


def extract_worksheet_xml(xlsx_path, worksheet_path):
    '''
    Extract worksheet XML element from an Excel spreadsheet.

    Useful, for example, to display worksheet contents:

    >>>> import lxml
    >>>> from openpyxl_helpers import extract_worksheet_xml
    >>>>
    >>>> template_root = extract_worksheet_xml(template_path, worksheet_path='xl/worksheets/sheet1.xml')
    >>>> print lxml.etree.tostring(template_root, pretty_print=True)

    .. versionadded:: 0.4

    Parameters
    ----------
    xlsx_path : str
        Path to Excel ``xlsx`` file.
    worksheet_path : str
        Path to worksheet, e.g., ``path`` attribute of an
        ``openpyxl.worksheet.worksheet.Worksheet`` instance.

    Returns
    -------
    lxml.etree._Element
        XML element for specified worksheet document.
    '''
    with zipfile.ZipFile(xlsx_path, mode='r') as input_zip:
        if worksheet_path.startswith('/'):
            worksheet_path = worksheet_path[1:]
        return lxml.etree.fromstring(input_zip.read(worksheet_path))
