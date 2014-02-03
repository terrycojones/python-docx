#!/usr/bin/env python

import lxml
from unittest import TestCase

from docx import (
    appproperties, contenttypes, coreproperties, getdocumenttext, heading,
    makeelement, newdocument, nsprefixes, opendocx, pagebreak, paragraph,
    picture, relationshiplist, replace, savedocx, search, table, websettings,
    wordrelationships)

TEST_FILE = 'ShortTest.docx'
IMAGE1_FILE = 'image1.png'


class TestDocx(TestCase):
    '''
    Test the docx module.
    '''

    def _simpledoc(self, noimagecopy=False):
        '''Make a docx (document, relationships) for use in other docx tests'''
        relationships = relationshiplist()
        imagefiledict = {}
        document = newdocument()
        body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
        body.append(heading('Heading 1', 1))
        body.append(heading('Heading 2', 2))
        body.append(paragraph('Paragraph 1'))
        for point in ['List Item 1', 'List Item 2', 'List Item 3']:
            body.append(paragraph(point, style='ListNumber'))
        body.append(pagebreak(type='page'))
        body.append(paragraph('Paragraph 2'))
        body.append(table([['A1', 'A2', 'A3'],
                           ['B1', 'B2', 'B3'],
                           ['C1', 'C2', 'C3']]))
        body.append(pagebreak(type='section', orient='portrait'))
        if noimagecopy:
            relationships, picpara, imagefiledict = picture(
                relationships, IMAGE1_FILE, 'This is a test description',
                imagefiledict=imagefiledict)
        else:
            relationships, picpara = picture(
                relationships, IMAGE1_FILE, 'This is a test description')
        body.append(picpara)
        body.append(pagebreak(type='section', orient='landscape'))
        body.append(paragraph('Paragraph 3'))
        if noimagecopy:
            return (document, body, relationships, imagefiledict)
        else:
            return (document, body, relationships)

    # --- Test Functions ---
    def testsearchandreplace(self):
        '''Ensure search and replace functions work'''
        document, body, relationships = self._simpledoc()
        body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
        self.assertTrue(search(body, 'ing 1'))
        self.assertTrue(search(body, 'ing 2'))
        self.assertTrue(search(body, 'graph 3'))
        self.assertTrue(search(body, 'ist Item'))
        self.assertTrue(search(body, 'A1'))
        if search(body, 'Paragraph 2'):
            body = replace(body, 'Paragraph 2', 'Whacko 55')
        self.assertTrue(search(body, 'Whacko 55'))

    def testtextextraction(self):
        '''Ensure text can be pulled out of a document'''
        document = opendocx(TEST_FILE)
        paratextlist = getdocumenttext(document)
        self.assertTrue(len(paratextlist) > 0)

    def testunsupportedpagebreak(self):
        '''Ensure unsupported page break types are trapped'''
        self.assertRaises(ValueError, pagebreak, type='unsup')

    def testnewdocument(self):
        '''Test that a new document can be created'''
        document, body, relationships = self._simpledoc()
        coreprops = coreproperties('Python docx testnewdocument',
                                   'An example of making docx from Python',
                                   'Alan Brooks',
                                   ['python', 'Office Open XML', 'Word'])
        savedocx(document, coreprops, appproperties(), contenttypes(),
                 websettings(), wordrelationships(relationships), TEST_FILE)

    def testnewdocument_noimagecopy(self):
        '''Test that a new document can be created'''
        document, body, relationships, imagefiledict = self._simpledoc(
            noimagecopy=True)
        coreprops = coreproperties('Python docx testnewdocument',
                                   'An example of making docx from Python',
                                   'Alan Brooks',
                                   ['python', 'Office Open XML', 'Word'])
        savedocx(document, coreprops, appproperties(), contenttypes(),
                 websettings(), wordrelationships(relationships), TEST_FILE,
                 imagefiledict=imagefiledict)

    def testopendocx(self):
        '''Ensure an etree element is returned'''
        self.assertTrue(isinstance(opendocx(TEST_FILE), lxml.etree._Element))

    def testmakeelement(self):
        '''Ensure custom elements get created'''
        testelement = makeelement('testname',
                                  attributes={'testattribute': 'testvalue'},
                                  tagtext='testtagtext')
        self.assertEqual('{http://schemas.openxmlformats.org/wordprocessingml/'
                         '2006/main}testname', testelement.tag)
        self.assertEqual(
            {
                '{http://schemas.openxmlformats.org/wordprocessingml/'
                '2006/main}testattribute': 'testvalue'
            },
            testelement.attrib)
        self.assertEqual('testtagtext', testelement.text)

    def testparagraph(self):
        '''Ensure paragraph creates p elements'''
        testpara = paragraph('paratext', style='BodyText')
        self.assertEqual(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p',
            testpara.tag)

    def testtable(self):
        '''Ensure tables make sense'''
        namespace = ('http://schemas.openxmlformats.org/wordprocessingml/'
                     '2006/main')
        testtable = table([['A1', 'A2'], ['B1', 'B2'], ['C1', 'C2']])
        self.assertEqual(
            'B2',
            testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',
                            namespaces={'ns0': namespace})[0].text)

if __name__ == '__main__':
    import nose
    nose.main()
