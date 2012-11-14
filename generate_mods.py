#!/usr/bin/env python
'''Pass the name of the spreadsheet to this script and
it will generate individual mods files for each record
in the mods_files directory, logging the output to dataset_mods.log.
Run './generate_mods.py --help' to see various options.

Notes: 
1. The python xlrd and lxml modules are required.
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
    def __init__(self, filename, forceDates=False, sheet=1, inputEncoding='utf-8'):
        '''Open file and get data from correct sheet.
        
        First, try opening the file as an excel spreadsheet.
        If that fails, try opening it as a CSV file.
        Exit with error if CSV doesn't work.
        '''
        logger.info('Opening "%s"' % filename)
        #set the date override value
        self.forceDates = forceDates
        self.inputEncoding = inputEncoding
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
                csvReader = csv.reader(self.utf_8_encoder(csvFile), dialect)
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

    def utf_8_encoder(self, unicode_csv_data):
        '''From docs.python.org/2.6/library/csv.html
        
        CSV module doesn't handle unicode objects, but should handle UTF-8 data.'''
        for line in unicode_csv_data:
            yield line.encode('utf-8')

    def get_control_row(self):
        '''Retrieve the row that controls MODS mapping locations.'''
        #assume that the second row contains the mapping data.
        return self.get_row(2)

    def get_filename_col(self):
        '''Get index of column that contains filenames.'''
        #try control row first
        for i, val in enumerate(self.get_control_row()):
            if val in [u'record name', u'Filename']:
                return i
        #try first row if needed
        for i, val in enumerate(self.get_row(1)):
            if val in [u'record name', u'Filename']:
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
            if index > 1:
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

    def get_total_rows(self):
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
    '''Map data into a correct XML string representing the MODS.
    
    Each instance of this class can only handle 1 MODS object.
    Note: we should be generating valid MODS, but we're not using a MODS (or 
        XML) class.
    Note: the get_mods function will parse the xml & return a 'pretty' version
        of it.
    '''
    def __init__(self, encoding='utf-8'):
        '''self._mods is just a simple string object.'''
        #assume that the encoding value string is good utf-8 (or ascii) data
        self._mods = u'<?xml version="1.0" encoding="' + unicode(encoding, 'utf-8') + u'"?>'
        self._mods += u'<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/ http://www.loc.gov/standards/mods/v3/mods-3-4.xsd">'
        self.dataSeparator = u'|'
        self.encoding = encoding

    def get_mods(self):
        '''Returns a 'pretty' version of the self._mods unicode object with
        the mods tag closed.
        '''
        mods = self._mods + u'</mods:mods>'
        #pass a string to fromstring, not unicode
        modsStr = mods.encode(self.encoding)
        modsTree = etree.fromstring(modsStr)
        modsPrettyStr = etree.tostring(modsTree, pretty_print=True,
                xml_declaration=True, encoding=self.encoding)
        return unicode(modsPrettyStr, self.encoding)

    def add_data(self, mods_loc, data):
        '''Method to actually put the data in the correct place of MODS obj.'''
        logger.debug('Putting "%s" into %s' % (data, mods_loc))
        #sanitize data
        data = data.replace(u'&', u'&amp;')
        data = data.replace(u'<', u'&lt;')
        data = data.replace(u'>', u'&gt;')
        #parse location info into tags
        loc = LocationParser(mods_loc)
        tags = loc.get_tags()
        elements = loc.get_elements() #list of element names
        #names take some different processing, so use a different function
        if elements[0] == 'mods:name':
            self._add_name_data(tags, elements, data)
            return
        #check for multiple values that should go in separate tags
        # eg. xyz || abc =>
        #   <tag1><tag2>xyz</tag2></tag1>
        #   <tag1><tag2>abc</tag2></tag1>
        for d in  data.split(self.dataSeparator):
            #open tags
            for tag in tags[:]:
                self._mods += tag
            #add data
            self._mods += d.strip()
            #close tags
            for el in reversed(elements[:]):
                self._mods += u'</' + el + u'>'

    def _add_name_data(self, tags, elements, data):
        '''Method to handle more complicated name data. '''
        logger.debug('_add_name_data: %s' % data)
        #split into array of | sections (different people)
        names = data.split(self.dataSeparator)
        for name in names:
            #open first tag we have
            self._mods += tags[0] #mods:name
            #we might have a name and then a role
            # eg. John Smith#creator
            divs = name.split(u'#')
            #add namePart subelement
            self._mods += tags[1]
            self._mods += divs[0].strip()
            self._mods += u'</mods:namePart>'
            #add role subelement if present
            if len(divs) > 1:
                self._mods += u'<mods:role>'
                self._mods += u'<mods:roleTerm>'
                self._mods += divs[1].strip() + u'</mods:roleTerm>'
                self._mods += u'</mods:role>'
            #close <mods:name> tag
            self._mods += u'</mods:name>'

class LocationParser(object):
    '''Small class for parsing dataset location instructions.'''
    def __init__(self, data):
        self.data = data
        self.tags = [] #list of full tags
        self.elements = [] #list of element names
        #self.attributes = [] #list of attributes for each element
        self.parse()
    def get_tags(self):
        return self.tags
    def get_elements(self):
        return self.elements
    #def get_attributes(self):
        #return self.attributes
    def parse(self):
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
                self.tags.append(tag)
            else:
                raise Exception('Error parsing "%s"!' % data)
            #remove first tag from data
            data = data[endTagPos+1:]
            #get element name and attributes to put in list
            space = tag.find(u' ')
            if space > 0:
                name = tag[1:space]
                #attributes = self.parse_attributes(tag[space:-1])
            else:
                name = tag[1:-1]
                #attributes = {}
            self.elements.append(name)
            #self.attributes.append(attributes)
    #def parse_attributes(self, data):
        #data = data.strip()
        #parse attributes into dict
        #attributes = {}
        #while len(data) > 0:
            #equal = data.find('=')
            #attr = data[:equal].strip()
            #valStart = data.find('"', equal+1)
            #valEnd = data.find('"', valStart+1)
            #if valEnd > valStart:
                #val = data[valStart+1:valEnd]
                #attributes[attr] = val
                #data = data[valEnd+1:].strip()
            #else:
                #logger.error('Error parsing attributes. data = "%s"' % data)
                #raise Exception('Error parsing attributes!')
        #return attributes


def process(dataHandler, outputEncoding):
    '''Function to go through all the data and process it.'''
    #get dicts of columns that should be mapped & where they go in MODS
    colsToMap = dataHandler.get_cols_to_map()
    totalRows = dataHandler.get_total_rows()
    filenameCol = dataHandler.get_filename_col()
    if filenameCol is None:
        logger.error('Could not get filename column!')
        sys.exit(1)
    #assume that data starts on third row
    for i in xrange(3, totalRows):
        #get list of unicode values from this record of the dataset
        row = dataHandler.get_row(i)
        filename = row[filenameCol].strip()
        if len(filename) == 0:
            logger.warning('No filename defined for row %d. Skipping.' % (i+1))
            continue
        filename = os.path.join(MODS_DIR, filename + u'.mods')
        ext = 1
        while os.path.exists(filename):
            filename = os.path.join(MODS_DIR, row[filenameCol].strip() + u'_' 
                + str(ext) + u'.mods')
            ext += 1
        logger.info('Processing row %d to %s.' % (i+1, filename))
        mapper = Mapper(outputEncoding)
        #for each column that should be mapped, pass the mapping
        # info and this row's data to the mapper to create the MODS
        for i, val in enumerate(row):
            if i in colsToMap and len(val) > 0:
                mapper.add_data(colsToMap[i], val)
        mods = mapper.get_mods()
        logger.debug(mods)
        #write data to a file (with the desired output encoding)
        with codecs.open(filename, 'w', outputEncoding) as f:
            f.write(mods)

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
    parser.add_option('-i', '--input-encoding',
                    action='store', dest='in_enc', default='utf-8',
                    help='specify the input encoding for CSV files (default is UTF-8)')
    #Had a problem when I tried testing with ebcdic (cp500). Not sure which ones work & which don't.
    parser.add_option('-o', '--output-encoding',
                    action='store', dest='out_enc', default='utf-8',
                    help='specify the encoding of the output MODS files (default is UTF-8)')
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
    dataHandler = DataHandler(args[0], options.force_dates, int(options.sheet), options.in_enc)
    process(dataHandler, options.out_enc)
    sys.exit()

