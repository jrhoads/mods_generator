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

class DataHandler(object):
    '''Handle interacting with the data.
    
    Use 1-based values for sheets or rows in public functions.
    There should be no data in str objects - they should all be unicode,
    which is what xlrd uses, and we convert all CSV data to unicode objects
    as well.
    '''
    def __init__(self, filename, inputEncoding='utf-8', sheet=1, ctrlRow=2, forceDates=False):
        '''Open file and get data from correct sheet.
        
        First, try opening the file as an excel spreadsheet.
        If that fails, try opening it as a CSV file.
        Exit with error if CSV doesn't work.
        '''
        logger.info('Opening "%s"' % filename)
        #set the date override value
        self.forceDates = forceDates
        self.inputEncoding = inputEncoding
        self.ctrlRow = ctrlRow
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

    def get_data_rows(self):
        '''data rows will be all the rows after the control row'''
        for i in xrange(self.ctrlRow+1, self._get_total_rows()+1): #xrange doesn't include the stop value
            yield self.get_row(i)

    def get_control_row(self):
        '''Retrieve the row that controls MODS mapping locations.'''
        return self.get_row(self.ctrlRow)

    def get_id_col(self):
        '''Get index of column that contains id (or filename).'''
        ID_NAMES = [u'record name', u'filename', u'id', u'file names']
        #try control row first
        for i, val in enumerate(self.get_control_row()):
            if val.lower() in ID_NAMES:
                return i
        #try first row if needed
        for i, val in enumerate(self.get_row(1)):
            if val.lower() in ID_NAMES:
                return i
        #return None if we didn't find anything
        return None

    def get_cols_to_map(self):
        '''Get a dict of columns & values in dataset that should be mapped to MODS
        (some will just be ignored).
        '''
        cols = {}
        ctrl_row = self.get_control_row()
        for i, val in enumerate(ctrl_row):
            #we'll assume it's to be mapped if we see the start of a MODS tag
            if u'<mods' in val:
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
            if index > (self.ctrlRow-1):
                for i, v in enumerate(self.get_control_row()):
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
    logger.debug('process_text_date: ' + strDate)
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
        logger.warning('Could not parse date string: ' + strDate)
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
        #parse location info into tags
        loc = LocationParser(mods_loc)
        tags = loc.get_tags()
        elements = loc.get_elements() #list of element names
        data_vals = [data.strip() for data in data.split(self.dataSeparator)]
        #handle various MODS elements
        if elements[0]['element'] == u'mods:name':
            if not self._cleared_fields.get(u'names', None):
                self._mods.names = []
                self._cleared_fields[u'names'] = True
            self._add_name_data(tags, elements, data_vals)
            return
        elif elements[0]['element'] == u'mods:namePart':
            #grab the last name that was added
            name = self._mods.names[-1]
            np = mods.NamePart(text=data)
            if u'type' in elements[0][u'attributes']:
                np.type = elements[0][u'attributes'][u'type']
            name.name_parts.append(np)
        elif elements[0][u'element'] == u'mods:titleInfo':
            if not self._cleared_fields.get(u'title_info_list', None):
                self._mods.title_info_list = []
                self._cleared_fields[u'title_info_list'] = True
            self._add_title_data(tags, elements, data_vals)
            return
        elif elements[0][u'element'] == u'mods:genre':
            if not self._cleared_fields.get(u'genres', None):
                self._mods.genres = []
                self._cleared_fields[u'genres'] = True
            for data in data_vals:
                genre = mods.Genre(text=data)
                if 'authority' in elements[0]['attributes']:
                    genre.authority = elements[0]['attributes']['authority']
                self._mods.genres.append(genre)
        elif elements[0]['element'] == 'mods:originInfo':
            if not self._cleared_fields.get(u'origin_info', None):
                self._mods.origin_info = None
                self._cleared_fields[u'origin_info'] = True
                self._mods.create_origin_info()
            self._add_origin_info_data(tags, elements, data_vals)
            return
        elif elements[0]['element'] == 'mods:physicalDescription':
            if not self._cleared_fields.get(u'physical_description', None):
                self._mods.physical_description = None
                self._cleared_fields[u'physical_description'] = True
                #can only have one physical description currently
                self._mods.create_physical_description()
            if elements[1]['element'] == 'mods:extent':
                self._mods.physical_description.extent = data_vals[0]
        elif elements[0]['element'] == 'mods:typeOfResource':
            if not self._cleared_fields.get(u'typeOfResource', None):
                self._mods.resource_type = None
                self._cleared_fields[u'typeOfResource'] = True
            self._mods.resource_type = data_vals[0]
        elif elements[0]['element'] == 'mods:abstract':
            if not self._cleared_fields.get(u'abstract', None):
                self._mods.abstract = None
                self._cleared_fields[u'abstract'] = True
                #can only have one abstract currently
                self._mods.create_abstract()
            self._mods.abstract.text = data_vals[0]
        elif elements[0]['element'] == 'mods:note':
            if not self._cleared_fields.get(u'notes', None):
                self._mods.notes = []
                self._cleared_fields[u'notes'] = True
            for data in data_vals:
                note = mods.Note(text=data)
                if 'type' in elements[0]['attributes']:
                    note.type = elements[0]['attributes']['type']
                if 'displayLabel' in elements[0]['attributes']:
                    note.label = elements[0]['attributes']['displayLabel']
                self._mods.notes.append(note)
        elif elements[0]['element'] == 'mods:subject':
            if not self._cleared_fields.get(u'subjects', None):
                self._mods.subjects = []
                self._cleared_fields[u'subjects'] = True
            for data in data_vals:
                subject = mods.Subject()
                if elements[1]['element'] == 'mods:topic':
                    subject.topic = data
                elif elements[1]['element'] == 'mods:geographic':
                    subject.geographic = data
                elif elements[1]['element'] == 'mods:hierarchicalGeographic':
                    hg = mods.HierarchicalGeographic()
                    if elements[2]['element'] == 'mods:country':
                        if 'data' in elements[2]:
                            hg.country = elements[2]['data']
                            if elements[3]['element'] == 'mods:state':
                                hg.state = data
                        else:
                            hg.country = data
                    subject.hierarchical_geographic = hg
                self._mods.subjects.append(subject)
        elif elements[0]['element'] == 'mods:identifier':
            if not self._cleared_fields.get(u'identifiers', None):
                self._mods.identifiers = []
                self._cleared_fields[u'identifiers'] = True
            for data in data_vals:
                identifier = mods.Identifier(text=data)
                if 'type' in elements[0]['attributes']:
                    identifier.type = elements[0]['attributes']['type']
                if 'displayLabel' in elements[0]['attributes']:
                    identifier.label = elements[0]['attributes']['displayLabel']
                self._mods.identifiers.append(identifier)
        elif elements[0]['element'] == 'mods:location':
            if not self._cleared_fields.get(u'locations', None):
                self._mods.locations = []
                self._cleared_fields[u'locations'] = True
            if elements[1]['element'] == 'mods:physicalLocation':
                for data in data_vals:
                    loc = mods.Location(physical=data)
                    self._mods.locations.append(loc)

    def _add_title_data(self, tags, elements, data_vals):
        for data in data_vals:
            title = mods.TitleInfo()
            divs = data.split(u'#')
            for element in elements[1:]:
                if element[u'element'] == u'mods:title':
                    title.title = divs[0]
                elif element[u'element'] == u'mods:partName':
                    if divs[1]:
                        title.part_name = divs[1]
                elif element[u'element'] == u'mods:partNumber':
                    if divs[2]:
                        title.part_number = divs[2]
            self._mods.title_info_list.append(title)

    def _add_name_data(self, tags, elements, data_vals):
        '''Method to handle more complicated name data. '''
        for data in data_vals:
            #elements[0] is mods:name
            name = mods.Name()
            if 'type' in elements[0]['attributes']:
                name.type = elements[0]['attributes']['type']
            #we might have a name and then a role, so split the data val
            # eg. John Smith#creator
            divs = data.split(u'#')
            role = None
            role_attrs = {}
            for element in elements[1:]:
                if element['element'] == 'mods:namePart':
                    np = mods.NamePart(text=divs[0])
                    name.name_parts.append(np)
                elif element['element'] == 'mods:role':
                    pass
                elif element['element'] == 'mods:roleTerm':
                    #add role subelement if present
                    if len(divs) > 1:
                        role = mods.Role(text=divs[1])
                        role_attrs = element['attributes']
                    elif 'data' in element:
                        role = mods.Role(text=element['data'])
                        role_attrs = element['attributes']
            if len(divs) > 1 and not role:
                role = mods.Role(text=divs[1])
            if role:
                if 'type' in role_attrs:
                    role.type = role_attrs['type']
                name.roles.append(role)
            self._mods.names.append(name)

    def _add_origin_info_data(self, tags, elements, data_vals):
        if 'displayLabel' in elements[0]['attributes']:
            self._mods.origin_info.label = elements[0]['attributes']['displayLabel']
        for data in data_vals:
            divs = data.split(u'#')
            for index, el in enumerate(elements[1:]):
                if el['element'] == 'mods:dateCreated':
                    date = mods.DateCreated(date=divs[index])
                    date = self._set_date_attributes(date, el['attributes'])
                    self._mods.origin_info.created.append(date)
                elif el['element'] == 'mods:dateOther':
                    date = mods.DateOther(date=divs[index])
                    date = self._set_date_attributes(date, el['attributes'])
                    self._mods.origin_info.other.append(date)
                else:
                    print('unhandled originInfo element: %s' % elements)
                    return

    def _set_date_attributes(self, date, attributes):
        if 'encoding' in attributes:
            date.encoding = attributes['encoding']
        if 'point' in attributes:
            date.point = attributes['point']
        if 'keyDate' in attributes:
            date.key_date = attributes['keyDate']
        return date


class LocationParser(object):
    '''Small class for parsing dataset location instructions.'''
    def __init__(self, data):
        self.data = data
        self.tags = [] #list of full tags
        self.elements = [] #list of element dicts
        self._parse()

    def get_tags(self):
        return self.tags

    def get_elements(self):
        return self.elements

    def get_attributes(self):
        return self.attributes

    def _parse(self):
        '''Get the first Mods field we're looking at in this string.'''
        #first strip off leading & trailing whitespace
        data = self.data.strip()
        #now pull out tags in order
        while len(data) > 0:
            #grab the first tag (including namespace)
            startTagPos = data.find(u'<')
            endTagPos = data.find(u'>')
            if endTagPos > startTagPos:
                tag = data[startTagPos:endTagPos+1]
                #remove first tag from data for the next loop
                data = data[endTagPos+1:]
                if tag[:2] == u'</':
                    continue
                self.tags.append(tag)
            else:
                raise Exception('Error parsing "%s"!' % data)
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
            if data:
                next_tag_start = data.find(u'<')
                if next_tag_start == 0:
                    pass
                elif next_tag_start == -1:
                    text = data
                    data = ''
                else:
                    text = data[:next_tag_start]
                    data = data[next_tag_start:]
            if text:
                self.elements.append({'element': name, 'attributes': attributes, 'data': text})
            else:
                self.elements.append({'element': name, 'attributes': attributes})
                    

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


def process(dataHandler):
    '''Function to go through all the data and process it.'''
    #get dicts of columns that should be mapped & where they go in MODS
    cols_to_map = dataHandler.get_cols_to_map()
    id_col = dataHandler.get_id_col()
    if id_col is None:
        logger.error('Could not get id column!')
        sys.exit(1)
    index = 1
    for row in dataHandler.get_data_rows():
        filename = row[id_col].strip()
        if len(filename) == 0:
            logger.warning('No filename defined for row %d. Skipping.' % (index))
            continue
        filename = os.path.join(MODS_DIR, filename + u'.mods')
        #load parent mods object if it exists
        parent_mods = None
        if os.path.exists(filename):
            parent_mods = load_xmlobject_from_file(filename, mods.Mods)
        ext = 1
        while os.path.exists(filename):
            filename = os.path.join(MODS_DIR, row[id_col].strip() + u'_' 
                + str(ext) + u'.mods')
            ext += 1
        logger.info('Processing row %d to %s.' % (index, filename))
        mapper = Mapper(parent_mods=parent_mods)
        #for each column that should be mapped, pass the mapping
        # info and this row's data to the mapper to create the MODS
        for i, val in enumerate(row):
            if i in cols_to_map and len(val) > 0:
                mapper.add_data(cols_to_map[i], val)
        mods_obj = mapper.get_mods()
        mods_data = unicode(mods_obj.serializeDocument(pretty=True), 'utf-8')
        with codecs.open(filename, 'w', 'utf-8') as f:
            f.write(mods_data)
        index = index + 1


if __name__ == '__main__':
    logger.info('Processing dataset to MODS files')
    #get options
    parser = OptionParser()
    parser.add_option('--force-dates',
                    action='store_true', dest='force_dates', default=False,
                    help='force date conversion even if ambiguous')
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
    dataHandler = DataHandler(args[0], options.in_enc, int(options.sheet), int(options.row), options.force_dates)
    process(dataHandler)
    sys.exit()

