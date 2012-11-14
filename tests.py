#!/usr/bin/env python
# -*- coding: utf-8 -*-

#Tests for generate_mods.py

import unittest
import os

from generate_mods import LocationParser, DataHandler, Mapper, process_text_date

class TestLocationParser(unittest.TestCase):

    def setUp(self):
        pass

    def test_single_tag(self):
        loc = '<mods:identifier type="local" displayLabel="PN_DB_id">'
        locParser = LocationParser(loc)
        self.assertEqual(len(locParser.get_elements()), 1)
        self.assertEqual(len(locParser.get_tags()), 1)
        self.assertEqual(locParser.get_elements()[0], 'mods:identifier')
        self.assertEqual(locParser.get_tags()[0], loc)

    def test_multi_tag(self):
        loc = '<mods:titleInfo><mods:title>'
        locParser = LocationParser(loc)
        self.assertEqual(len(locParser.get_elements()), 2)
        self.assertEqual(len(locParser.get_tags()), 2)
        self.assertEqual(locParser.get_elements()[0], 'mods:titleInfo')
        self.assertEqual(locParser.get_elements()[1], 'mods:title')
        self.assertEqual(locParser.get_tags()[0], '<mods:titleInfo>')
        self.assertEqual(locParser.get_tags()[1], '<mods:title>')

    def test_invalid_loc(self):
        loc = 'asdf1234'
        try:
            locParser = LocationParser(loc)
        except Exception:
            #return successfully if Exception was raised
            return
        #if we got here, no Exception was raised, so fail the test
        self.fail('Did not raise Exception on bad input!')

class TestDataHandler(unittest.TestCase):
    '''added some non-ascii characters to the files to make sure
    they're handled properly (é is u00e9, ă is u0103)'''

    def setUp(self):
        pass

    def test_xls(self):
        dh = DataHandler(os.path.join('test_files', 'data.xls'))
        ctrlRow = dh.get_control_row()
        self.assertEqual(dh.get_filename_col(), 2)
        self.assertEqual(dh.get_total_rows(), 4)
        self.assertEqual(ctrlRow[3], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        row3 = dh.get_row(3)
        self.assertEqual(row3[7], u'Test 1')
        for cell in row3:
            self.assertTrue(isinstance(cell, unicode))
        #test that process_text_date is working right
        self.assertEqual(row3[11], u'2005-10-21')
        self.assertEqual(row3[22], u'2005-10-10')
        #test that we can get the second sheet correctly
        dh = DataHandler(os.path.join('test_files', 'data.xls'), sheet=2)
        row3 = dh.get_row(3)
        self.assertEqual(row3[11], u'2008-10-21')
        self.assertEqual(dh.get_total_rows(), 3)

    def test_xlsx(self):
        dh = DataHandler(os.path.join('test_files', 'data.xlsx'))
        #get list of unicode objects
        ctrlRow = dh.get_control_row()
        self.assertEqual(dh.get_filename_col(), 2)
        self.assertEqual(dh.get_total_rows(), 4)
        self.assertEqual(ctrlRow[3], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        row3 = dh.get_row(3)
        for cell in row3:
            self.assertTrue(isinstance(cell, unicode))
        self.assertEqual(row3[11], u'2005-10-21')
        self.assertEqual(row3[22], u'2005-10-10')

    def test_csv(self):
        dh = DataHandler(os.path.join('test_files', 'data.csv'))
        ctrlRow = dh.get_control_row()
        for cell in ctrlRow:
            self.assertTrue(isinstance(cell, unicode))
        self.assertEqual(dh.get_filename_col(), 2)
        self.assertEqual(dh.get_total_rows(), 4)
        self.assertEqual(ctrlRow[3], u'<mods:identifier type="local" displayLabel="Originăl noé.">')

    def test_csv_small(self):
        dh = DataHandler(os.path.join('test_files', 'data-small.csv'))
        ctrlRow = dh.get_control_row()
        for cell in ctrlRow:
            self.assertTrue(isinstance(cell, unicode))
        self.assertEqual(dh.get_filename_col(), 2)
        self.assertEqual(dh.get_total_rows(), 3)
        self.assertEqual(ctrlRow[3], u'<mods:identifier type="local" displayLabel="Originăl noé.">')

class TestOther(unittest.TestCase):
    '''Test non-class functions.'''

    def test_process_text_date(self):
        '''Tests to make sure we're handling dates properly.'''
        #dates with slashes
        self.assertEqual(process_text_date('5/14/2000'), '2000-05-14')
        self.assertEqual(process_text_date('14/5/2000'), '2000-05-14')
        #dates with dashes
        self.assertEqual(process_text_date('3-17-2013'), '2013-03-17')
        self.assertEqual(process_text_date('17-3-2013'), '2013-03-17')
        #ambiguous dates or invalid dates should stay as they are
        self.assertEqual(process_text_date('5/4/99'), '5/4/99') #4/5 or 5/4
        self.assertEqual(process_text_date('5/14/00'), '5/14/00')
        self.assertEqual(process_text_date('05/14/00'), '05/14/00')
        self.assertEqual(process_text_date('14/5/00'), '14/5/00')
        self.assertEqual(process_text_date('14/05/00'), '14/05/00')
        self.assertEqual(process_text_date('3-17-13'), '3-17-13')
        self.assertEqual(process_text_date('03-17-13'), '03-17-13')
        self.assertEqual(process_text_date('17-3-13'), '17-3-13')
        self.assertEqual(process_text_date('17-03-13'), '17-03-13')
        self.assertEqual(process_text_date('3-3-03'), '3-3-03')
        self.assertEqual(process_text_date('03-17-03'), '03-17-03')
        self.assertEqual(process_text_date('05/14/01'), '05/14/01')
        self.assertEqual(process_text_date('05-14-12'), '05-14-12')
        self.assertEqual(process_text_date(1), 1)
        self.assertEqual(process_text_date(''), '')
        self.assertEqual(process_text_date(None), None)

        #test override as well
        self.assertEqual(process_text_date('5/4/99', True), '1999-05-04')
        self.assertEqual(process_text_date('5/17/99', True), '1999-05-17')

class TestMapper(unittest.TestCase):
    '''Test Mapper class.'''

    EMPTY_MODS = u'''<?xml version='1.0' encoding='utf-8'?>
<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/ http://www.loc.gov/standards/mods/v3/mods-3-4.xsd"/>
'''
    UTF16_MODS = u'''<?xml version='1.0' encoding='utf-16'?>
<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/ http://www.loc.gov/standards/mods/v3/mods-3-4.xsd">
  <mods:originInfo displayLabel="Date Ądded to Colléction">
    <mods:dateOther>2010-01-31</mods:dateOther>
  </mods:originInfo>
</mods:mods>
'''
    FULL_MODS = u'''<?xml version='1.0' encoding='utf-8'?>
<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/ http://www.loc.gov/standards/mods/v3/mods-3-4.xsd">
  <mods:identifier type="local" displayLabel="Original no.">1591</mods:identifier>
  <mods:identifier type="local" displayLabel="PN_DB_id">321</mods:identifier>
  <mods:titleInfo>
    <mods:title>é. #1 Test</mods:title>
  </mods:titleInfo>
  <mods:genre authority="aat">Programming Tests</mods:genre>
  <mods:originInfo displayLabel="Date Ądded to Colléction">
    <mods:dateOther>2010-01-31</mods:dateOther>
  </mods:originInfo>
  <mods:subject>
    <mods:topic>PROGRĄMMING</mods:topic>
  </mods:subject>
  <mods:subject>
    <mods:topic>Testing</mods:topic>
  </mods:subject>
  <mods:subject>
    <mods:topic>Recursion</mods:topic>
  </mods:subject>
  <mods:subject>
    <mods:geographic>United States</mods:geographic>
  </mods:subject>
  <mods:name type="personal">
    <mods:namePart>Smith</mods:namePart>
    <mods:role>
      <mods:roleTerm>creator</mods:roleTerm>
    </mods:role>
  </mods:name>
  <mods:name type="personal">
    <mods:namePart>Jones, T.</mods:namePart>
  </mods:name>
  <mods:originInfo>
    <mods:dateCreated>7/13/1899</mods:dateCreated>
  </mods:originInfo>
  <mods:note>Note 1&amp;2</mods:note>
  <mods:note>3&lt;4</mods:note>
  <mods:location>
    <mods:physicalLocation>zzz</mods:physicalLocation>
  </mods:location>
  <mods:note>another note</mods:note>
</mods:mods>
'''

    def test_mods_output(self):
        m1 = Mapper()
        mods = m1.get_mods()
        self.assertTrue(isinstance(mods, unicode))
        self.assertEqual(mods, self.EMPTY_MODS)
        m2 = Mapper('utf-16')
        m2.add_data(u'<mods:originInfo displayLabel="Date Ądded to Colléction"><mods:dateOther>', u'2010-01-31')
        mods = m2.get_mods()
        self.assertTrue(isinstance(mods, unicode))
        self.assertEqual(mods, self.UTF16_MODS)
        #add all data as unicode, since that's how it should be coming from DataHandler
        m = Mapper()
        m.add_data(u'<mods:identifier type="local" displayLabel="Original no.">', u'1591')
        m.add_data(u'<mods:identifier type="local" displayLabel="PN_DB_id">', u'321')
        m.add_data(u'<mods:titleInfo><mods:title>', u'é. #1 Test')
        m.add_data(u'<mods:genre authority="aat">', u'Programming Tests')
        m.add_data(u'<mods:originInfo displayLabel="Date Ądded to Colléction"><mods:dateOther>', u'2010-01-31')
        m.add_data(u'<mods:subject><mods:topic>', u'PROGRĄMMING | Testing')
        m.add_data(u'<mods:subject><mods:topic>', u'Recursion')
        m.add_data(u'<mods:subject><mods:geographic>', u'United States')
        m.add_data(u'<mods:name type="personal"><mods:namePart>', u'Smith#creator | Jones, T.')
        m.add_data(u'<mods:originInfo><mods:dateCreated>', u'7/13/1899')
        m.add_data(u'<mods:note>', u'Note 1&2')
        m.add_data(u'<mods:note>', u'3<4')
        m.add_data(u'<mods:location><mods:physicalLocation>', u'zzz')
        m.add_data(u'<mods:note>', u'another note')
        mods = m.get_mods()
        self.assertTrue(isinstance(mods, unicode))
        #this does assume that the attributes will always be written out in the same order
        self.assertEqual(mods, self.FULL_MODS)

if __name__ == '__main__':
    runner = unittest.TextTestRunner(verbosity=2)
    unittest.main(testRunner=runner)
