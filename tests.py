#!/usr/bin/env python
# -*- coding: utf-8 -*-
import unittest
import os

from generate_mods import LocationParser, DataHandler, Mapper, process_text_date
from bdrxml.mods import Mods

class TestLocationParser(unittest.TestCase):

    def setUp(self):
        pass

    def test_single_tag(self):
        loc = u'<mods:identifier type="local" displayLabel="PN_DB_id">'
        locParser = LocationParser(loc)
        base_element = locParser.get_base_element()
        self.assertEqual(base_element[u'element'], u'mods:identifier')
        self.assertEqual(base_element[u'attributes'], {u'type': u'local', u'displayLabel': u'PN_DB_id'})
        self.assertFalse(base_element[u'data'])
        sections = locParser.get_sections()
        self.assertFalse(sections)

    def test_multi_tag(self):
        loc = u'<mods:titleInfo><mods:title>'
        locParser = LocationParser(loc)
        base_element = locParser.get_base_element()
        self.assertEqual(base_element[u'element'], u'mods:titleInfo')
        self.assertEqual(base_element[u'attributes'], {})
        self.assertFalse(base_element[u'data'])
        sections = locParser.get_sections()
        self.assertEqual(len(sections), 1)
        first_section = sections[0]
        self.assertEqual(len(first_section), 1)
        self.assertEqual(first_section[0][u'element'], u'mods:title')
        self.assertEqual(first_section[0][u'attributes'], {})
        self.assertFalse(first_section[0][u'data'])

    def test_name_tag(self):
        loc = u'<mods:name type="personal"><mods:namePart>#<mods:role><mods:roleTerm type="text">winner'
        locParser = LocationParser(loc)
        base_element = locParser.get_base_element()
        self.assertEqual(base_element[u'element'], u'mods:name')
        self.assertEqual(base_element[u'attributes'], {u'type': u'personal'})
        self.assertFalse(base_element[u'data'])
        sections = locParser.get_sections()
        self.assertEqual(len(sections), 2)
        first_section = sections[0]
        self.assertEqual(len(first_section), 1)
        self.assertEqual(first_section[0][u'element'], u'mods:namePart')
        self.assertEqual(first_section[0][u'attributes'], {})
        self.assertFalse(first_section[0][u'data'])
        second_section = sections[1]
        self.assertEqual(len(second_section), 2)
        self.assertEqual(second_section[0][u'element'], u'mods:role')
        self.assertEqual(second_section[0][u'attributes'], {})
        self.assertFalse(second_section[0][u'data'], {})
        self.assertEqual(second_section[1][u'element'], u'mods:roleTerm')
        self.assertEqual(second_section[1][u'attributes'], {u'type': u'text'})
        self.assertEqual(second_section[1][u'data'], u'winner')

    def test_another_tag(self):
        loc = '<mods:subject><mods:hierarchicalGeographic><mods:country>United States</mods:country><mods:state>'
        locParser = LocationParser(loc)
        base_element = locParser.get_base_element()
        self.assertEqual(base_element[u'element'], u'mods:subject')
        self.assertEqual(base_element[u'attributes'], {})
        self.assertFalse(base_element[u'data'])
        sections = locParser.get_sections()
        self.assertEqual(len(sections), 1)
        first_section = sections[0]
        self.assertEqual(len(first_section), 3)
        self.assertEqual(first_section[0]['element'], 'mods:hierarchicalGeographic')
        self.assertEqual(first_section[1]['element'], 'mods:country')
        self.assertEqual(first_section[1]['data'], 'United States')
        self.assertEqual(first_section[2]['element'], 'mods:state')

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
        mods_records = dh.get_mods_records()
        self.assertEqual(len(mods_records), 2)
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['mods_path'], unicode))
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['data'], unicode))
        self.assertEqual(mods_records[0].field_data()[0]['mods_path'], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        self.assertEqual(mods_records[0].field_data()[0]['data'], u'123')
        self.assertEqual(mods_records[0].field_data()[2]['mods_path'], u'<mods:titleInfo><mods:title>')
        self.assertEqual(mods_records[0].field_data()[2]['data'], u'Test 1')
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(mods_records[1].id, u'test2')
        self.assertEqual(mods_records[0].mods_id, u'test1')
        self.assertEqual(mods_records[1].mods_id, u'test2')
        #test that process_text_date is working right
        self.assertEqual(mods_records[0].field_data()[4]['data'], u'2005-10-21')
        #test that we can get the second sheet correctly
        dh = DataHandler(os.path.join('test_files', 'data.xls'), sheet=2)
        mods_records = dh.get_mods_records()
        self.assertEqual(len(mods_records), 1)
        self.assertEqual(mods_records[0].mods_id, u'mods0001')
        self.assertEqual(mods_records[0].field_data()[5]['data'], u'2008-10-21')

    def test_xlsx(self):
        dh = DataHandler(os.path.join('test_files', 'data.xlsx'), obj_type='child')
        mods_records = dh.get_mods_records()
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(len(mods_records), 2)
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['mods_path'], unicode))
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['data'], unicode))
        self.assertEqual(mods_records[0].field_data()[0]['mods_path'], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        self.assertEqual(mods_records[0].field_data()[0]['data'], u'123')
        self.assertEqual(mods_records[0].field_data()[2]['mods_path'], u'<mods:titleInfo><mods:title>')
        self.assertEqual(mods_records[0].field_data()[2]['data'], u'Test 1')
        self.assertEqual(mods_records[0].field_data()[4]['data'], u'2005-10-21')
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(mods_records[0].mods_id, u'test1_1') #_1 because it's a child
        self.assertEqual(mods_records[1].id, u'test1')
        self.assertEqual(mods_records[1].mods_id, u'test1_2')

    def test_csv(self):
        dh = DataHandler(os.path.join('test_files', 'data.csv'))
        mods_records = dh.get_mods_records()
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(len(mods_records), 2)
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['mods_path'], unicode))
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['data'], unicode))
        self.assertEqual(mods_records[0].field_data()[0]['mods_path'], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        self.assertEqual(mods_records[0].field_data()[0]['data'], u'123')
        self.assertEqual(mods_records[0].field_data()[2]['mods_path'], u'<mods:titleInfo><mods:title>')
        self.assertEqual(mods_records[0].field_data()[2]['data'], u'Test 1')
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(mods_records[1].id, u'test2')
        self.assertEqual(mods_records[0].field_data()[4]['data'], u'2005-10-21')

    def test_csv_small(self):
        dh = DataHandler(os.path.join('test_files', 'data-small.csv'))
        mods_records = dh.get_mods_records()
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(len(mods_records), 1)
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['mods_path'], unicode))
        self.assertTrue(isinstance(mods_records[0].field_data()[0]['data'], unicode))
        self.assertEqual(mods_records[0].field_data()[0]['mods_path'], u'<mods:identifier type="local" displayLabel="Originăl noé.">')
        self.assertEqual(mods_records[0].field_data()[0]['data'], u'123')
        self.assertEqual(mods_records[0].field_data()[2]['mods_path'], u'<mods:titleInfo><mods:title>')
        self.assertEqual(mods_records[0].field_data()[2]['data'], u'Test 1')
        self.assertEqual(mods_records[0].id, u'test1')
        self.assertEqual(mods_records[0].field_data()[4]['data'], u'2005-10-21')


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

    EMPTY_MODS = u'''<?xml version='1.0' encoding='UTF-8'?>
<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd"/>
'''
    FULL_MODS = u'''<?xml version='1.0' encoding='UTF-8'?>
<mods:mods xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd" ID="mods000">
  <mods:physicalDescription>
    <mods:extent>1 video file</mods:extent>
    <mods:digitalOrigin>reformatted digital</mods:digitalOrigin>
    <mods:note>note 1</mods:note>
  </mods:physicalDescription>
  <mods:titleInfo>
    <mods:title>é. 1 Test</mods:title>
    <mods:partName>part #1</mods:partName>
    <mods:partNumber>1</mods:partNumber>
  </mods:titleInfo>
  <mods:titleInfo type="alternative" displayLabel="display">
    <mods:title>Alt Title</mods:title>
    <mods:nonSort>The</mods:nonSort>
  </mods:titleInfo>
  <mods:identifier type="local" displayLabel="Original no.">1591</mods:identifier>
  <mods:identifier type="local" displayLabel="PN_DB_id">321</mods:identifier>
  <mods:genre authority="aat">Programming Tests</mods:genre>
  <mods:originInfo displayLabel="Date Ądded to Colléction">
    <mods:publisher>Publisher</mods:publisher>
    <mods:place>
      <mods:placeTerm>USA</mods:placeTerm>
    </mods:place>
    <mods:dateOther encoding="w3cdtf" keyDate="yes">2010-01-31</mods:dateOther>
    <mods:dateCreated encoding="w3cdtf" point="end">7/13/1899</mods:dateCreated>
    <mods:dateCreated encoding="w3cdtf">1972-10-1973-07-07</mods:dateCreated>
    <mods:dateCreated encoding="w3cdtf" point="start" keyDate="yes">1972-10</mods:dateCreated>
    <mods:dateCreated encoding="w3cdtf" point="end">1973-07-07</mods:dateCreated>
    <mods:dateIssued encoding="w3cdtf">1974-01-01</mods:dateIssued>
    <mods:dateCaptured encoding="w3cdtf">1975-01-01</mods:dateCaptured>
    <mods:dateValid encoding="w3cdtf">1976-01-01</mods:dateValid>
    <mods:dateModified encoding="w3cdtf">1977-01-01</mods:dateModified>
    <mods:copyrightDate>1978-01-##</mods:copyrightDate>
  </mods:originInfo>
  <mods:subject>
    <mods:topic>PROGRĄMMING</mods:topic>
  </mods:subject>
  <mods:subject>
    <mods:topic>Testing</mods:topic>
  </mods:subject>
  <mods:subject>
    <mods:topic>Software</mods:topic>
    <mods:topic>Testing</mods:topic>
  </mods:subject>
  <mods:subject authority="local">
    <mods:topic>Recursion</mods:topic>
  </mods:subject>
  <mods:subject authority="local">
    <mods:temporal>1990s</mods:temporal>
  </mods:subject>
  <mods:subject>
    <mods:geographic>United States</mods:geographic>
  </mods:subject>
  <mods:subject>
    <mods:hierarchicalGeographic>
      <mods:country>United States</mods:country>
      <mods:state>Pennsylvania</mods:state>
    </mods:hierarchicalGeographic>
  </mods:subject>
  <mods:name type="personal">
    <mods:namePart>Smith</mods:namePart>
    <mods:role>
      <mods:roleTerm>creator</mods:roleTerm>
    </mods:role>
  </mods:name>
  <mods:name type="personal">
    <mods:namePart>Jones, T.</mods:namePart>
    <mods:namePart type="date">1799-1889</mods:namePart>
  </mods:name>
  <mods:name type="personal">
    <mods:namePart>Bob</mods:namePart>
    <mods:role>
      <mods:roleTerm type="text">winner</mods:roleTerm>
    </mods:role>
  </mods:name>
  <mods:name type="personal">
    <mods:namePart>Fob, Bob</mods:namePart>
  </mods:name>
  <mods:name type="personal">
    <mods:namePart>Smith, Ted</mods:namePart>
    <mods:namePart type="date">1900-2013</mods:namePart>
    <mods:namePart type="termsOfAddress">Sir</mods:namePart>
  </mods:name>
  <mods:note displayLabel="note label">Note 1&amp;2</mods:note>
  <mods:note>3&lt;4</mods:note>
  <mods:note>another note</mods:note>
  <mods:location>
    <mods:physicalLocation>zzz</mods:physicalLocation>
    <mods:url>http://www.example.com</mods:url>
    <mods:holdingSimple>
      <mods:copyInformation>
        <mods:note>Note 1</mods:note>
      </mods:copyInformation>
    </mods:holdingSimple>
  </mods:location>
  <mods:typeOfResource>video</mods:typeOfResource>
  <mods:language>
    <mods:languageTerm authority="iso639-2b" type="code">eng</mods:languageTerm>
  </mods:language>
  <mods:relatedItem type="related item" displayLabel="display">
    <mods:titleInfo>
      <mods:title>Some related item display title</mods:title>
    </mods:titleInfo>
  </mods:relatedItem>
</mods:mods>
'''

    def test_mods_output(self):
        self.maxDiff = None
        m1 = Mapper()
        mods = m1.get_mods()
        self.assertTrue(isinstance(mods, Mods))
        self.assertEqual(unicode(mods.serializeDocument(pretty=True), 'utf-8'), self.EMPTY_MODS)
        #put some data in here, so we can pass this as a parent_mods to the next test
        # these next two should be deleted and not displayed twice
        m1.add_data(u'<mods:identifier type="local" displayLabel="Original no.">', u'1591')
        m1.add_data(u'<mods:subject><mods:topic>', u'Recursion')
        # this one isn't added again, and should still be in the output
        m1.add_data(u'<mods:physicalDescription><mods:extent>#<mods:digitalOrigin>#<mods:note>', u'1 video file#reformatted digital#note 1')
        #add all data as unicode, since that's how it should be coming from DataHandler
        m = Mapper(parent_mods=m1.get_mods())
        m.add_data(u'<mods:mods ID="">', u'mods000')
        m.add_data(u'<mods:titleInfo><mods:title>#<mods:partName>#<mods:partNumber>', u'é. 1 Test#part \#1#1')
        m.add_data(u'<mods:titleInfo type="alternative" displayLabel="display"><mods:title>#<mods:nonSort>', u'Alt Title#The')
        m.add_data(u'<mods:identifier type="local" displayLabel="Original no.">', u'1591')
        m.add_data(u'<mods:identifier type="local" displayLabel="PN_DB_id">', u'321')
        m.add_data(u'<mods:genre authority="aat">', u'Programming Tests')
        m.add_data(u'<mods:originInfo><mods:publisher>', u'Publisher')
        m.add_data(u'<mods:originInfo><mods:place><mods:placeTerm>', u'USA')
        m.add_data(u'<mods:originInfo displayLabel="Date Ądded to Colléction"><mods:dateOther encoding="w3cdtf" keyDate="yes">', u'2010-01-31')
        m.add_data(u'<mods:subject><mods:topic>', u'PROGRĄMMING || Testing')
        m.add_data(u'<mods:subject><mods:topic>#<mods:topic>', u'Software#Testing')
        m.add_data(u'<mods:subject authority="local"><mods:topic>', u'Recursion || ')
        m.add_data(u'<mods:subject authority="local"><mods:temporal>', u'1990s')
        m.add_data(u'<mods:subject><mods:geographic>', u'United States')
        m.add_data(u'<mods:subject><mods:hierarchicalGeographic><mods:country>United States</mods:country><mods:state>', u'Pennsylvania')
        m.add_data(u'<mods:name type="personal"><mods:namePart>#<mods:role><mods:roleTerm>', u'Smith#creator || Jones, T.')
        m.add_data(u'<mods:namePart type="date">', u'1799-1889')
        m.add_data(u'<mods:name type="personal"><mods:namePart>#<mods:role><mods:roleTerm type="text">winner', u'Bob')
        m.add_data(u'<mods:name type="personal"><mods:namePart>#<mods:namePart type="date">#<mods:namePart type="termsOfAddress">', u'Fob, Bob || Smith, Ted#1900-2013#Sir')
        m.add_data(u'<mods:originInfo><mods:dateCreated encoding="w3cdtf" point="end">', u'7/13/1899')
        m.add_data(u'<mods:originInfo><mods:dateCreated encoding="w3cdtf">#<mods:dateCreated encoding="w3cdtf" point="start" keyDate="yes">#<mods:dateCreated encoding="w3cdtf" point="end">', u'1972-10-1973-07-07#1972-10#1973-07-07')
        m.add_data(u'<mods:note displayLabel="note label">', u'Note 1&2')
        m.add_data(u'<mods:note>', u'3<4')
        m.add_data(u'<mods:location><mods:physicalLocation>zzz#<mods:url>#<mods:holdingSimple><mods:copyInformation><mods:note>', u'#http://www.example.com#Note 1')
        m.add_data(u'<mods:note>', u'another note')
        m.add_data(u'<mods:typeOfResource>', u'video')
        m.add_data(u'<mods:language><mods:languageTerm authority="iso639-2b" type="code">', u'eng')
        m.add_data(u'<mods:relatedItem type="related item" displayLabel="display"><mods:titleInfo><mods:title>', u'Some related item display title')
        m.add_data(u'<mods:originInfo><mods:dateIssued encoding="w3cdtf">', u'1974-01-01')
        m.add_data(u'<mods:originInfo><mods:dateCaptured encoding="w3cdtf">', u'1975-01-01')
        m.add_data(u'<mods:originInfo><mods:dateValid encoding="w3cdtf">', u'1976-01-01')
        m.add_data(u'<mods:originInfo><mods:dateModified encoding="w3cdtf">', u'1977-01-01')
        m.add_data(u'<mods:originInfo><mods:copyrightDate>', u'1978-01-##')
        mods = m.get_mods()
        mods_data = unicode(mods.serializeDocument(pretty=True), 'utf-8')
        self.assertTrue(isinstance(mods, Mods))
        self.assertEqual(mods.title_info_list[0].title, u'é. 1 Test')
        self.assertEqual(mods.title_info_list[0].part_number, u'1')
        self.assertEqual(mods.title_info_list[0].part_name, u'part #1')
        #this does assume that the attributes will always be written out in the same order
        self.assertEqual(mods_data, self.FULL_MODS)

    def test_get_data_divs(self):
        m = Mapper()
        self.assertEqual(m._get_data_divs(u'part1#part2#part3', False), [u'part1#part2#part3'])
        self.assertEqual(m._get_data_divs(u'part1#part2#part3', True), [u'part1', u'part2', u'part3'])
        self.assertEqual(m._get_data_divs(u'part\#1#part2#part\#3', True), [u'part#1', u'part2', u'part#3'])
        self.assertEqual(m._get_data_divs(u'part\#1 and \#1a#part2#part\#3', True), [u'part#1 and #1a', u'part2', u'part#3'])


if __name__ == '__main__':
    runner = unittest.TextTestRunner(verbosity=2)
    unittest.main(testRunner=runner)

