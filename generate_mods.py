#!/usr/bin/env python
'''Pass the name of the spreadsheet to this script and
it will generate individual mods files for each record
in the mods_files directory, logging the output to dataset_mods.log.
Run './generate_mods.py --help' to see various options.

Notes: 
1. Requirements: xlrd, lxml, and bdrxml.
2. The spreadsheet can be any version of Excel, or a CSV file.
3. The first row of the dataset is for headers, the second row is for
    MODS mapping tags, and the rest of the rows are for the data.
4. The control row should have the full MODS path for the data, in the
    following format: <mods:name type="personal"><mods:namePart>
5. Unicode - all text strings from xlrd (for Excel files) are Unicode. For xlrd
    numbers, we convert those into Unicode, since we're just writing text out
    to files. The encoding of CSV files can be specified as an argument (if 
    it's not a valid encoding for Python, a LookupError will be raised). The
    encoding of the output files can also be specified as an argument (if
    there's an input character that can't be encoded in the output encoding, a
    UnicodeEncodeError will be raised).

'''
import csv
import io
import sys
import string
import logging
import logging.handlers
import datetime
import os
import codecs
import re
from optparse import OptionParser

from lxml import etree
import xlrd
from eulxml.xmlmap import load_xmlobject_from_file
from bdrxml import mods

#set up logging to console & log file
LOG_FILENAME = 'dataset_mods.log'
logger = logging.getLogger('simple')
logger.setLevel(logging.DEBUG)
fileHandler = logging.handlers.RotatingFileHandler(
                LOG_FILENAME, maxBytes=10000000, backupCount=5)
logFormat = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
fileHandler.setFormatter(logFormat)
logger.addHandler(fileHandler)
consoleHandler = logging.StreamHandler()
consoleHandler.setLevel(logging.INFO)
consFormat = logging.Formatter("%(levelname)s %(message)s")
consoleHandler.setFormatter(consFormat)
logger.addHandler(consoleHandler)

#directory for mods files
MODS_DIR = "mods_files"


class ModsRecord(object):

    def __init__(self, id, mods_id, field_data, data_files):
        self.id = id #this is what ties parent records to children
        self.mods_id = mods_id #this object's mods id (from a column or calculated)
        self.parent_mods_filename = u'%s.mods' % id
        self.mods_filename = u'%s.mods' % mods_id
        self._field_data = field_data
        self.data_files = data_files

    def field_data(self):
        #return list of {'mods_path': xxx, 'data': xxx}
        return self._field_data


class DataHandler(object):
    '''Handle interacting with the data.
    
    Use 1-based values for sheets or rows in public functions.
    There should be no data in str objects - they should all be unicode,
    which is what xlrd uses, and we convert all CSV data to unicode objects
    as well.
    '''
    def __init__(self, filename, inputEncoding='utf-8', sheet=1, ctrlRow=2, forceDates=False, obj_type='parent'):
        '''Open file and get data from correct sheet.
        
        First, try opening the file as an excel spreadsheet.
        If that fails, try opening it as a CSV file.
        Exit with error if CSV doesn't work.
        '''
        self.obj_type = obj_type
        #set the date override value
        self.forceDates = forceDates
        self.inputEncoding = inputEncoding
        self._ctrlRow = ctrlRow
        #open file
        try:
            self.book = xlrd.open_workbook(filename)
            self.dataset = self.book.sheet_by_index(int(sheet)-1)
            self.dataType = 'xlrd'
            logger.debug('Got "%s" dataset.' % self.dataset.name)
        except xlrd.XLRDError as xerr:
            logger.debug('Failed xlrd open: %s.' % repr(xerr))
            #now try using csv
            try:
                #need to open the file with whatever encoding it's in
                logger.debug('opening file with ' + self.inputEncoding + ' encoding.')
                csvFile = codecs.open(filename, 'r', self.inputEncoding)
                #read some test data to pass to sniffer for checking the dialect
                data = csvFile.read(4096) #data is unicode object
                csvFile.seek(0)
                #Sniffer needs data encoded in ascii (just drop non-ascii characters for now)
                dataAscii = data.encode('ascii', 'ignore')
                dialect = csv.Sniffer().sniff(dataAscii)
                #set doublequote to true because that's the default and the Sniffer doesn't
                #   seem to pick it up right
                dialect.doublequote = True
                self.dataType = 'csv'
                #CSV module doesn't handle unicode correctly, so temporarily
                #   encode data as UTF-8, which it can handle.
                csvReader = csv.reader(self._utf_8_encoder(csvFile), dialect)
                #self.csvData is a list of lists of the row data
                self.csvData = []
                for row in csvReader:
                    if len(row) > 0:
                        #convert all the data back to unicode since we're done w/ CSV module
                        row = [unicode(cell, 'utf-8') for cell in row]
                        self.csvData.append(row)
                logger.debug('Got CSV data')
                csvFile.close()
            except Exception as e:
                logger.error(str(e))
                logger.error('Could not recognize file format. Exiting.')
                csvFile.close()
                sys.exit(1)

    def get_mods_records(self):
        id_col = self._get_id_col()
        if id_col is None:
            raise Exception('no ID column')
        index = self._ctrlRow
        mods_records = []
        mods_ids = {}
        data_file_col = self._get_filename_col()
        for data_row in self._get_data_rows():
            index += 1
            rec_id = data_row[id_col].strip()
            if not rec_id:
                logger.warning('no id on row %s - skipping' % index)
                continue
            mods_id_col = self._get_mods_id_col()
            if mods_id_col is not None:
                mods_id = data_row[mods_id_col].strip()
            else:
                if rec_id in mods_ids:
                    mods_id = u'%s_%s' % (rec_id, mods_ids[rec_id])
                    mods_ids[rec_id] = mods_ids[rec_id] + 1
                else:
                    if self.obj_type == 'parent':
                        mods_id = rec_id
                        mods_ids[rec_id] = 1
                    else:
                        mods_id = u'%s_1' % rec_id
                        mods_ids[rec_id] = 2
            field_data = []
            cols_to_map = self.get_cols_to_map()
            for i, val in enumerate(data_row):
                if i in cols_to_map and len(val) > 0:
                    field_data.append({'mods_path': cols_to_map[i], 'data': val})
            data_files = []
            if data_file_col is not None:
                data_files = [df.strip() for df in data_row[data_file_col].split(u',')]
            mods_records.append(ModsRecord(rec_id, mods_id, field_data, data_files))
        return mods_records

    def _get_data_rows(self):
        '''data rows will be all the rows after the control row'''
        for i in xrange(self._ctrlRow+1, self._get_total_rows()+1): #xrange doesn't include the stop value
            yield self.get_row(i)

    def _get_control_row(self):
        '''Retrieve the row that controls MODS mapping locations.'''
        return self.get_row(self._ctrlRow)

    def _get_col_from_id_names(self, id_names):
        #try control row first
        for i, val in enumerate(self._get_control_row()):
            if val.lower() in id_names:
                return i
        #try first row if needed
        for i, val in enumerate(self.get_row(1)):
            if val.lower() in id_names:
                return i
        #return None if we didn't find anything
        return None

    def _get_mods_id_col(self):
        ID_NAMES = [u'mods id', '<mods:mods id="">']
        return self._get_col_from_id_names(ID_NAMES)

    def _get_id_col(self):
        '''Get index of column that contains id for tying children to parents'''
        ID_NAMES = [u'id', u'tracker item id', u'tracker id', u'record name', u'file id']
        return self._get_col_from_id_names(ID_NAMES)

    def _get_filename_col(self):
        '''Get index of column that contains data file name.'''
        ID_NAMES = [u'file name', u'filename', u'file_id']
        return self._get_col_from_id_names(ID_NAMES)

    def get_cols_to_map(self):
        '''Get a dict of columns & values in dataset that should be mapped to MODS
        (some will just be ignored).
        '''
        cols = {}
        ctrl_row = self._get_control_row()
        for i, val in enumerate(ctrl_row):
            #we'll assume it's to be mapped if we see the start of a MODS tag
            if val.startswith(u'<mods'):
                cols[i] = val
        return cols

    def get_row(self, index):
        '''Retrieve a list of unicode values (index is 1-based like excel)'''
        #subtract 1 from index so that it's 0-based like xlrd and csvData list
        index = index - 1
        if self.dataType == 'xlrd':
            row = self.dataset.row_values(index)
            #In a data column that's mapped to a date field, we could find a text
            #   string that looks like a date - we might want to reformat 
            #   that as well.
            if index > (self._ctrlRow-1):
                for i, v in enumerate(self._get_control_row()):
                    if 'date' in v:
                        if isinstance(row[i], basestring):
                            #we may have a text date, so see if we can understand it
                            # *process_text_date will return a text value of the
                            #   reformatted date if possible, else the original value
                            row[i] = process_text_date(row[i], self.forceDates)
            for i, v in enumerate(row):
                if isinstance(v, float):
                    #there are some interesting things that happen
                    # with numbers in Excel. Eg. what looks like an int in Excel
                    # is actually stored as a float (and xlrd handles as a float).
                    #http://stackoverflow.com/questions/2739989/reading-numeric-excel-data-as-text-using-xlrd-in-python
                    #if cell is XL_CELL_NUMBER
                    if self.dataset.cell_type(index, i) == 2 and int(v) == v:
                        #convert data into int & then unicode
                        #Note: if a number was displayed as xxxx.0 in Excel, we
                        #   would lose the .0 here
                        row[i] = unicode(int(v))
                    #Dates are also stored as floats in Excel, so we have to do
                    #   some extra processing to get a datetime object
                    #if we have an XL_CELL_DATE
                    elif self.dataset.cell_type(index, i) == 3:
                        #try to get an actual date out of it, instead of a float
                        #Note: we are losing Excel formatting information here,
                        #   and formatting the date as yyyy-mm-dd.
                        tup = xlrd.xldate_as_tuple(v, self.book.datemode)
                        d = datetime.datetime(*tup)
                        if tup[0] == 0 and tup[1] == 0 and tup[2] == 0:
                            #just time, no date
                            row[i] = unicode('{0:%H:%M:%S}'.format(d))
                        elif tup[3] == 0 and tup[4] == 0 and tup[5] == 0:
                            #just date, no time
                            row[i] = unicode('{0:%Y-%m-%d}'.format(d))
                        else:
                            #assume full date/time
                            row[i] = unicode('{0:%Y-%m-%d %H:%M:%S}'.format(d))
        elif self.dataType == 'csv':
            row = self.csvData[index]
            if index > (self._ctrlRow-1):
                for i, v in enumerate(self._get_control_row()):
                    if 'date' in v:
                        if isinstance(row[i], basestring):
                            #we may have a text date, so see if we can understand it
                            # *process_text_date will return a text value of the
                            #   reformatted date if possible, else the original value
                            row[i] = process_text_date(row[i], self.forceDates)
        #this final loop should be unnecessary, but it's a final check to
        #   make sure everything is unicode.
        for i, v in enumerate(row):
            if not isinstance(v, unicode):
                try:
                    row[i] = unicode(v, self.inputEncoding)
                #if v isn't a string, we might get this error, so try without
                #   the encoding
                except TypeError:
                    row[i] = unicode(v)
        #finally return the row
        return row

    def _utf_8_encoder(self, unicode_csv_data):
        '''From docs.python.org/2.6/library/csv.html
        
        CSV module doesn't handle unicode objects, but should handle UTF-8 data.'''
        for line in unicode_csv_data:
            yield line.encode('utf-8')

    def _get_total_rows(self):
        '''Get total number of rows in the dataset.'''
        totalRows = 0
        if self.dataType == 'xlrd':
            totalRows = self.dataset.nrows
        elif self.dataType == 'csv':
            totalRows = len(self.csvData)
        return totalRows


def process_text_date(strDate, forceDates=False):
    '''Take a text-based date and try to reformat it to yyyy-mm-dd if needed.
        
    Note: in xx/xx/xx or xx-xx-xx, we assume that year is last, not first.'''
    #do some checking on strDate - if it's not what we're looking for,
    #   just return strDate without changing anything
    if not isinstance(strDate, basestring):
        return strDate
    if len(strDate) == 0:
        return strDate
    #Some date formats we could understand:
    #dd/dd/dddd, dd/dd/dd, d/d/dd, ...
    mmddyy = re.compile('^\d?\d/\d?\d/\d\d$')
    mmddyyyy = re.compile('^\d?\d/\d?\d/\d\d\d\d$')
    #dd-dd-dddd, dd-dd-dd, d-d-dd, ...
    mmddyy2 = re.compile('^\d?\d-\d?\d-\d\d$')
    mmddyyyy2 = re.compile('^\d?\d-\d?\d-\d\d\d\d$')
    format = '' #flag to remember which format we used
    if mmddyy.search(strDate):
        try:
            #try mm/dd/yy first, since that should be more common in the US
            newDate = datetime.datetime.strptime(strDate, '%m/%d/%y')
            format = 'mmddyy'
        except ValueError:
            try:
                newDate = datetime.datetime.strptime(strDate, '%d/%m/%y')
                format = 'ddmmyy'
            except ValueError:
                logger.warning('Error creating date from ' + strDate)
                return strDate
    elif mmddyyyy.search(strDate):
        try:
            newDate = datetime.datetime.strptime(strDate, '%m/%d/%Y')
            format = 'mmddyyyy'
        except ValueError:
            try:
                newDate = datetime.datetime.strptime(strDate, '%d/%m/%Y')
                format = 'ddmmyyyy'
            except ValueError:
                logger.warning('Error creating date from ' + strDate)
                return strDate
    elif mmddyy2.search(strDate):
        try:
            #try mm-dd-yy first, since that should be more common
            newDate = datetime.datetime.strptime(strDate, '%m-%d-%y')
            format = 'mmddyy'
        except ValueError:
            try:
                newDate = datetime.datetime.strptime(strDate, '%d-%m-%y')
                format = 'ddmmyy'
            except ValueError:
                logger.warning('Error creating date from ' + strDate)
                return strDate
    elif mmddyyyy2.search(strDate):
        try:
            newDate = datetime.datetime.strptime(strDate, '%m-%d-%Y')
            format = 'mmddyyyy'
        except ValueError:
            try:
                newDate = datetime.datetime.strptime(strDate, '%d-%m-%Y')
                format = 'ddmmyyyy'
            except ValueError:
                logger.warning('Error creating date from ' + strDate)
                return strDate
    else:
        #logger.warning('Could not parse date string: ' + strDate)
        return strDate
    #at this point, we have newDate, but it could still have been ambiguous
    #day & month are both between 1 and 12 & not equal - ambiguous
    if newDate.day <= 12 and newDate.day != newDate.month: 
        if forceDates:
            logger.warning('Ambiguous day/month: ' + strDate + 
                        '. Using it anyway.')
            return newDate.strftime('%Y-%m-%d')
        else:
            logger.warning('Ambiguous day/month: ' + strDate)
            return strDate
    #year is only two digits - don't know the century, or if year was
    # interchanged with month or day
    elif format == 'mmddyy' or format == 'ddmmyy':
        if forceDates:
            logger.warning('Ambiguous year: ' + strDate +
                        '. Using it anyway.')
            return newDate.strftime('%Y-%m-%d')
        else:
            logger.warning('Ambiguous year: ' + strDate)
            return strDate
    else:
        return newDate.strftime('%Y-%m-%d')


class Mapper(object):
    '''Map data into a Mods object.
    Each instance of this class can only handle 1 MODS object.'''

    def __init__(self, encoding='utf-8', parent_mods=None):
        self.dataSeparator = u'||'
        self.encoding = encoding
        self._parent_mods = parent_mods
        #dict for keeping track of which fields we've cleared out the parent
        # info for. So we can have multiple columns in the spreadsheet w/ the same field.
        self._cleared_fields = {}
        if parent_mods:
            self._mods = parent_mods
        else:
            self._mods = mods.make_mods()

    def get_mods(self):
        return self._mods

    def add_data(self, mods_loc, data):
        '''Method to actually put the data in the correct place of MODS obj.'''
        #parse location info into elements/attributes
        loc = LocationParser(mods_loc)
        base_element = loc.get_base_element()
        location_sections = loc.get_sections()
        data_vals = [data.strip() for data in data.split(self.dataSeparator)]
        #strip any empty data sections so we don't have to worry about it below
        data_vals = [self._get_data_divs(data, loc.has_sectioned_data) for data in data_vals if data]
        #handle various MODS elements
        if base_element['element'] == u'mods:mods':
            if 'ID' in base_element['attributes']:
                self._mods.id = data_vals[0][0]
        elif base_element['element'] == u'mods:name':
            if not self._cleared_fields.get(u'names', None):
                self._mods.names = []
                self._cleared_fields[u'names'] = True
            self._add_name_data(base_element, location_sections, data_vals)
        elif base_element['element'] == u'mods:namePart':
            #grab the last name that was added
            name = self._mods.names[-1]
            np = mods.NamePart(text=data_vals[0][0])
            if u'type' in base_element[u'attributes']:
                np.type = base_element[u'attributes'][u'type']
            name.name_parts.append(np)
        elif base_element[u'element'] == u'mods:titleInfo':
            if not self._cleared_fields.get(u'title_info_list', None):
                self._mods.title_info_list = []
                self._cleared_fields[u'title_info_list'] = True
            self._add_title_data(base_element, location_sections, data_vals)
        elif base_element[u'element'] == u'mods:language':
            if not self._cleared_fields.get(u'languages', None):
                self._mods.languages = []
                self._cleared_fields[u'languages'] = True
            for data in data_vals:
                language = mods.Language()
                language_term = mods.LanguageTerm(text=data[0])
                if u'authority' in location_sections[0][0]['attributes']:
                    language_term.authority = location_sections[0][0]['attributes']['authority']
                if u'type' in location_sections[0][0]['attributes']:
                    language_term.type = location_sections[0][0][u'attributes'][u'type']
                language.terms.append(language_term)
                self._mods.languages.append(language)
        elif base_element[u'element'] == u'mods:genre':
            if not self._cleared_fields.get(u'genres', None):
                self._mods.genres = []
                self._cleared_fields[u'genres'] = True
            for data in data_vals:
                genre = mods.Genre(text=data[0])
                if 'authority' in base_element['attributes']:
                    genre.authority = base_element['attributes']['authority']
                self._mods.genres.append(genre)
        elif base_element['element'] == 'mods:originInfo':
            if not self._cleared_fields.get(u'origin_info', None):
                self._mods.origin_info = None
                self._cleared_fields[u'origin_info'] = True
                self._mods.create_origin_info()
            self._add_origin_info_data(base_element, location_sections, data_vals)
        elif base_element['element'] == 'mods:physicalDescription':
            if not self._cleared_fields.get(u'physical_description', None):
                self._mods.physical_description = None
                self._cleared_fields[u'physical_description'] = True
                #can only have one physical description currently
                self._mods.create_physical_description()
            data_divs = data_vals[0]
            for index, section in enumerate(location_sections):
                if section[0][u'element'] == 'mods:extent':
                    self._mods.physical_description.extent = data_divs[index]
                elif section[0][u'element'] == 'mods:digitalOrigin':
                    self._mods.physical_description.digital_origin = data_divs[index]
                elif section[0][u'element'] == 'mods:note':
                    self._mods.physical_description.note = data_divs[index]
        elif base_element['element'] == 'mods:typeOfResource':
            if not self._cleared_fields.get(u'typeOfResource', None):
                self._mods.resource_type = None
                self._cleared_fields[u'typeOfResource'] = True
            self._mods.resource_type = data_vals[0][0]
        elif base_element['element'] == 'mods:abstract':
            if not self._cleared_fields.get(u'abstract', None):
                self._mods.abstract = None
                self._cleared_fields[u'abstract'] = True
                #can only have one abstract currently
                self._mods.create_abstract()
            self._mods.abstract.text = data_vals[0][0]
        elif base_element['element'] == 'mods:note':
            if not self._cleared_fields.get(u'notes', None):
                self._mods.notes = []
                self._cleared_fields[u'notes'] = True
            for data in data_vals:
                note = mods.Note(text=data[0])
                if 'type' in base_element['attributes']:
                    note.type = base_element['attributes']['type']
                if 'displayLabel' in base_element['attributes']:
                    note.label = base_element['attributes']['displayLabel']
                self._mods.notes.append(note)
        elif base_element['element'] == 'mods:subject':
            if not self._cleared_fields.get(u'subjects', None):
                self._mods.subjects = []
                self._cleared_fields[u'subjects'] = True
            for data in data_vals:
                subject = mods.Subject()
                if 'authority' in base_element['attributes']:
                    subject.authority = base_element['attributes']['authority']
                data_divs = data
                for section, div in zip(location_sections, data_divs):
                    if section[0]['element'] == 'mods:topic':
                        topic = mods.Topic(text=div)
                        subject.topic_list.append(topic)
                    elif section[0]['element'] == 'mods:temporal':
                        temporal = mods.Temporal(text=div)
                        subject.temporal_list.append(temporal)
                    elif section[0]['element'] == 'mods:geographic':
                        subject.geographic = div
                    elif section[0]['element'] == 'mods:hierarchicalGeographic':
                        print(u'%s' % section)
                        hg = mods.HierarchicalGeographic()
                        if section[1]['element'] == 'mods:country':
                            if 'data' in section[1]:
                                hg.country = section[1]['data']
                                if section[2]['element'] == 'mods:state':
                                    hg.state = div
                            else:
                                hg.country = div
                        subject.hierarchical_geographic = hg
                self._mods.subjects.append(subject)
        elif base_element['element'] == 'mods:identifier':
            if not self._cleared_fields.get(u'identifiers', None):
                self._mods.identifiers = []
                self._cleared_fields[u'identifiers'] = True
            for data in data_vals:
                identifier = mods.Identifier(text=data[0])
                if 'type' in base_element['attributes']:
                    identifier.type = base_element['attributes']['type']
                if 'displayLabel' in base_element['attributes']:
                    identifier.label = base_element['attributes']['displayLabel']
                self._mods.identifiers.append(identifier)
        elif base_element['element'] == u'mods:location':
            if not self._cleared_fields.get(u'locations', None):
                self._mods.locations = []
                self._cleared_fields[u'locations'] = True
            for data in data_vals:
                loc = mods.Location()
                data_divs = data
                for section, div in zip(location_sections, data_divs):
                    if section[0]['element'] == u'mods:url':
                        if section[0]['data']:
                            loc.url = section[0]['data']
                        else:
                            loc.url = div
                    elif section[0]['element'] == u'mods:physicalLocation':
                        if section[0]['data']:
                            loc.physical = section[0]['data']
                        else:
                            loc.physical = div
                    elif section[0]['element'] == u'mods:holdingSimple':
                        hs = mods.HoldingSimple()
                        if section[1]['element'] == u'mods:copyInformation':
                            if section[2]['element'] == u'mods:note':
                                note = mods.Note(text=div)
                                ci = mods.CopyInformation()
                                ci.notes.append(note)
                                hs.copy_information.append(ci)
                                loc.holding_simple = hs
                self._mods.locations.append(loc)
        elif base_element['element'] == u'mods:relatedItem':
            if not self._cleared_fields.get(u'related', None):
                self._mods.related_items = []
                self._cleared_fields[u'related'] = True
            for data in data_vals:
                related_item = mods.RelatedItem()
                if u'type' in base_element[u'attributes']:
                    related_item.type = base_element[u'attributes'][u'type']
                if u'displayLabel' in base_element[u'attributes']:
                    related_item.label = base_element[u'attributes'][u'displayLabel']
                if location_sections[0][0][u'element'] == u'mods:titleInfo':
                    if location_sections[0][1][u'element'] == u'mods:title':
                        related_item.title = data[0]
                self._mods.related_items.append(related_item)
        else:
            logger.error('element not handled! %s' % base_element)
            raise Exception('element not handled!')

    def _add_title_data(self, base_element, location_sections, data_vals):
        for data_divs in data_vals:
            title = mods.TitleInfo()
            if u'type' in base_element['attributes']:
                title.type = base_element['attributes']['type']
            if u'displayLabel' in base_element['attributes']:
                title.label = base_element['attributes']['displayLabel']
            for section, div in zip(location_sections, data_divs):
                for element in section:
                    if element[u'element'] == u'mods:title':
                        title.title = div
                    elif element[u'element'] == u'mods:partName':
                        title.part_name = div
                    elif element[u'element'] == u'mods:partNumber':
                        title.part_number = div
                    elif element[u'element'] == u'mods:nonSort':
                        title.non_sort = div
            self._mods.title_info_list.append(title)

    def _get_data_divs(self, data, has_sectioned_data):
        data_divs = []
        if not has_sectioned_data:
            return [data]
        #split data into its divisions based on '#', but allow \ to escape the #
        while data:
            ind = data.find(u'#')
            if ind == -1:
                data_divs.append(data)
                data = ''
            else:
                while ind != -1 and data[ind-1] == u'\\':
                    #remove '\'
                    data = data[:ind-1] + data[ind:]
                    #find next '#' (being sure to advance past current '#')
                    ind = data.find(u'#', ind)
                if ind == -1:
                    data_divs.append(data)
                    data = u''
                else:
                    data_divs.append(data[:ind])
                    data = data[ind+1:]
        return data_divs


    def _add_name_data(self, base_element, location_sections, data_vals):
        '''Method to handle more complicated name data. '''
        for data in data_vals:
            name = mods.Name() #we're always going to be creating a name
            if u'type' in base_element[u'attributes']:
                name.type = base_element[u'attributes'][u'type']
            data_divs = data
            for index, section in enumerate(location_sections):
                try:
                    div = data_divs[index].strip()
                except:
                    div = None
                #make sure we have data for this section (except for mods:role, which could just have a constant)
                if not div and section[0][u'element'] != u'mods:role':
                    continue
                for element in section:
                    #handle base name
                    if element['element'] == u'mods:namePart' and u'type' not in element['attributes']:
                        np = mods.NamePart(text=div)
                        name.name_parts.append(np)
                    elif element[u'element'] == u'mods:namePart' and u'type' in element[u'attributes']:
                        np = mods.NamePart(text=div)
                        np.type = element[u'attributes'][u'type']
                        name.name_parts.append(np)
                    elif element['element'] == u'mods:roleTerm':
                        role_attrs = element['attributes']
                        if element[u'data']:
                            role = mods.Role(text=element['data'])
                        else:
                            if div:
                                role = mods.Role(text=div)
                            else:
                                continue
                        if u'type' in role_attrs:
                            role.type = role_attrs['type']
                        if u'authority' in role_attrs:
                            role.authority = role_attrs[u'authority']
                        name.roles.append(role)
            self._mods.names.append(name)

    def _add_origin_info_data(self, base_element, location_sections, data_vals):
        if u'displayLabel' in base_element['attributes']:
            self._mods.origin_info.label = base_element[u'attributes'][u'displayLabel']
        for data in data_vals:
            divs = data
            for index, section in enumerate(location_sections):
                if not divs[index]:
                    continue
                if section[0][u'element'] == u'mods:dateCreated':
                    date = mods.DateCreated(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.created.append(date)
                elif section[0][u'element'] == u'mods:dateIssued':
                    date = mods.DateIssued(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.issued.append(date)
                elif section[0][u'element'] == u'mods:dateCaptured':
                    date = mods.DateCaptured(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.captured.append(date)
                elif section[0][u'element'] == u'mods:dateValid':
                    date = mods.DateValid(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.valid.append(date)
                elif section[0][u'element'] == u'mods:dateModified':
                    date = mods.DateModified(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.modified.append(date)
                elif section[0][u'element'] == u'mods:copyrightDate':
                    date = mods.CopyrightDate(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.copyright.append(date)
                elif section[0][u'element'] == u'mods:dateOther':
                    date = mods.DateOther(date=divs[index])
                    date = self._set_date_attributes(date, section[0][u'attributes'])
                    self._mods.origin_info.other.append(date)
                elif section[0][u'element'] == u'mods:place':
                    place = mods.Place()
                    placeTerm = mods.PlaceTerm(text=divs[index])
                    place.place_terms.append(placeTerm)
                    self._mods.origin_info.places.append(place)
                elif section[0][u'element'] == u'mods:publisher':
                    self._mods.origin_info.publisher = divs[index]
                else:
                    print(u'unhandled originInfo element: %s' % section)
                    raise Exception('unhandled originInfo element: %s' % section)

    def _set_date_attributes(self, date, attributes):
        if u'encoding' in attributes:
            date.encoding = attributes[u'encoding']
        if u'point' in attributes:
            date.point = attributes[u'point']
        if u'keyDate' in attributes:
            date.key_date = attributes[u'keyDate']
        return date


class LocationParser(object):
    '''class for parsing dataset location instructions.
    eg. <mods:name type="personal"><mods:namePart>#<mods:namePart type="date">#<mods:namePart type="termsOfAddress">'''

    def __init__(self, data):
        self.has_sectioned_data = False
        self._data = data #raw data we receive
        self._section_separator = u'#'
        self._base_element = None #in the example, this will be set to {'element': 'mods:name', 'attributes': {'type': 'personal'}}
        self._sections = [] #list of the sections, which are divided by '#' (in the example, there are 3 sections)
            #each section consists of a list of elements
            #each element is a dict containing the element name, its attributes, and any data in that element
        self._parse()

    def get_base_element(self):
        return self._base_element

    def get_sections(self):
        return self._sections

    def _parse_base_element(self, data):
        #grab the first tag (including namespace) & parse into self._base_element
        startTagPos = data.find(u'<')
        endTagPos = data.find(u'>')
        if endTagPos > startTagPos:
            tag = data[startTagPos:endTagPos+1]
            #remove first tag from data for the rest of the parsing
            data = data[endTagPos+1:]
            #parse tag into elements & attributes
            space = tag.find(u' ')
            if space > 0:
                name = tag[1:space]
                attributes = self._parse_attributes(tag[space:-1])
            else:
                name = tag[1:-1]
                attributes = {}
            return ({u'element': name, u'attributes': attributes, u'data': None}, data)
        else:
            raise Exception('Error parsing "%s"!' % data.encode('utf-8'))

    def _parse(self):
        '''Get the first Mods field we're looking at in this string.'''
        #first strip off leading & trailing whitespace
        data = self._data.strip()
        #very basic data checking
        if data[0] != u'<':
            raise Exception('location data must start with "<"')
        #grab base element (eg. mods:originInfo, mods:name, ...)
        self._base_element, data = self._parse_base_element(data)
        if not data:
            return #we're done - there was just one base element
        #now pull out elements/attributes in order, for each section
        location_sections = data.split(self._section_separator)
        if len(location_sections) > 1:
            self.has_sectioned_data = True
        for section in location_sections:
            new_section = []
            while len(section) > 0:
                #grab the first tag (including namespace)
                startTagPos = section.find(u'<')
                endTagPos = section.find(u'>')
                if endTagPos > startTagPos:
                    tag = section[startTagPos:endTagPos+1]
                    #remove first tag from section for the next loop
                    section = section[endTagPos+1:]
                    if tag[:2] == u'</':
                        continue
                else:
                    raise Exception('Error parsing "%s"!' % section)
                #get element name and attributes to put in list
                space = tag.find(u' ')
                if space > 0:
                    name = tag[1:space]
                    attributes = self._parse_attributes(tag[space:-1])
                else:
                    name = tag[1:-1]
                    attributes = {}
                #there could be some text before the next tag
                text = None
                if section:
                    next_tag_start = section.find(u'<')
                    if next_tag_start == 0:
                        pass
                    elif next_tag_start == -1:
                        text = section
                        section = ''
                    else:
                        text = section[:next_tag_start]
                        section = section[next_tag_start:]
                if text:
                    new_section.append({'element': name, 'attributes': attributes, 'data': text})
                else:
                    new_section.append({'element': name, 'attributes': attributes, u'data': None})
            if new_section:
                self._sections.append(new_section)


    def _parse_attributes(self, data):
        data = data.strip()
        attributes = {}
        while len(data) > 0:
            equal = data.find('=')
            attr = data[:equal].strip()
            valStart = data.find('"', equal+1)
            valEnd = data.find('"', valStart+1)
            if valEnd > valStart:
                val = data[valStart+1:valEnd]
                attributes[attr] = val
                data = data[valEnd+1:].strip()
            else:
                logger.error('Error parsing attributes. data = "%s"' % data)
                raise Exception('Error parsing attributes!')
        return attributes


def get_mods_filename(parent_id, mods_id=None):
    #use a mods id value if available
    #otherwise, take the id and loop until we get a filename that doesn't exist yet
    if mods_id:
        base_filename = mods_id
    else:
        base_filename = parent_id
    filename = os.path.join(MODS_DIR, '%s.mods' % base_filename)
    ext = 1
    while os.path.exists(filename):
        filename = os.path.join(MODS_DIR, base_filename + u'_' 
            + str(ext) + u'.mods')
        ext += 1
    return filename


def process(dataHandler, copy_parent_to_children=False):
    '''Function to go through all the data and process it.'''
    #get dicts of columns that should be mapped & where they go in MODS
    index = 1
    for record in dataHandler.get_mods_records():
        filename = record.mods_filename
        if os.path.exists(os.path.join(MODS_DIR, filename)):
            raise Exception('%s already exists!' % filename)
        logger.info('Processing row %d to %s.' % (index, filename))
        if copy_parent_to_children:
            #load parent mods object if desired (& it exists)
            parent_filename = os.path.join(MODS_DIR, record.parent_mods_filename)
            parent_mods = None
            if os.path.exists(parent_filename):
                parent_mods = load_xmlobject_from_file(parent_filename, mods.Mods)
                mapper = Mapper(parent_mods=parent_mods)
        else:
            mapper = Mapper()
        for field in record.field_data():
            mapper.add_data(field['mods_path'], field['data'])
        mods_obj = mapper.get_mods()
        mods_data = unicode(mods_obj.serializeDocument(pretty=True), 'utf-8')
        with codecs.open(os.path.join(MODS_DIR, filename), 'w', 'utf-8') as f:
            f.write(mods_data)
        index = index + 1


if __name__ == '__main__':
    logger.info('Processing dataset to MODS files')
    #get options
    parser = OptionParser()
    parser.add_option('-t', '--type',
                    action='store', dest='type', default='parent',
                    help='type of records (parent or child, default is parent)')
    parser.add_option('--force-dates',
                    action='store_true', dest='force_dates', default=False,
                    help='force date conversion even if ambiguous')
    parser.add_option('--copy-parent-to-children',
                    action='store_true', dest='copy_parent_to_children', default=False,
                    help='copy parent data into children')
    parser.add_option('-s', '--sheet',
                    action='store', dest='sheet', default=1,
                    help='specify the sheet number (starting at 1) in an Excel spreadsheet')
    parser.add_option('-r', '--ctrl_row',
                    action='store', dest='row', default=2,
                    help='specify the control row number (starting at 1) in an Excel spreadsheet')
    parser.add_option('-i', '--input-encoding',
                    action='store', dest='in_enc', default='utf-8',
                    help='specify the input encoding for CSV files (default is UTF-8)')
    (options, args) = parser.parse_args()
    #make sure we have a directory to put the mods files in
    try:
        os.makedirs(MODS_DIR)
    except OSError as err:
        if os.path.isdir(MODS_DIR):
            pass
        else:
            #dir creation error - re-raise it
            raise
    #set up data handler & process data
    dataHandler = DataHandler(args[0], options.in_enc, int(options.sheet), int(options.row), options.force_dates, options.type)
    process(dataHandler, options.copy_parent_to_children)
    sys.exit()

