#!/usr/bin/env python2.6
'''
Test docx module
'''
import os
import lxml
import nose
from docx import *

TEST_FILE = 'Short python-docx test.docx'

# --- Setup & Support Functions ---
def setup_module():
    '''Set up test fixtures'''
    testnewdocument()

def teardown_module():
    '''Tear down test fixtures'''
    if TEST_FILE in os.listdir('.'):
        os.remove(TEST_FILE)

def simpledoc():
    '''Make a docx (document, relationships) for use in other docx tests'''
    relationships = relationshiplist()
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    docbody.append(heading('Heading 1',1)  )   
    docbody.append(heading('Heading 2',2))
    docbody.append(paragraph('Paragraph 1'))
    for point in ['List Item 1','List Item 2','List Item 3']:
        docbody.append(paragraph(point,style='ListNumber'))
    docbody.append(pagebreak(type='page'))
    docbody.append(paragraph('Paragraph 2')) 
    docbody.append(table([['A1','A2','A3'],['B1','B2','B3'],['C1','C2','C3']]))
    docbody.append(pagebreak(type='section', orient='portrait'))
    relationships,picpara = picture(relationships,'image1.png','This is a test description')
    docbody.append(picpara)
    docbody.append(pagebreak(type='section', orient='landscape'))
    docbody.append(paragraph('Paragraph 3'))
    return (document, docbody, relationships)


# --- Test Functions ---
def testsearchandreplace():
    '''Ensure search and replace functions work'''
    document, docbody, relationships = simpledoc()
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    if search(docbody, 'Paragraph 2'): 
        docbody = replace(docbody,'Paragraph 2','Whacko 55') 
    assert search(docbody, 'Whacko 55')
    assert False # replace works, search fails for some cases
    
def testtextextraction():
    '''Ensure text can be pulled out of a document'''
    document = opendocx(TEST_FILE)
    paratextlist = getdocumenttext(document)
    assert len(paratextlist) > 0

def testunsupportedpagebreak():
    '''Ensure unsupported page break types are trapped'''
    document = newdocument()
    docbody = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    try:
        docbody.append(pagebreak(type='unsup'))
    except ValueError:
        return # passed
    assert False # failed
    
def testnewdocument():
    '''Test that a new document can be created'''
    document, docbody, relationships = simpledoc()
    properties = docproperties('Python docx testnewdocument','A short example of making docx from Python','Alan Brooks',['python','Office Open XML','Word'])
    savedocx(document, properties, contenttypes(), websettings(), wordrelationships(relationships), TEST_FILE)

def testopendocx():
    '''Ensure an etree element is returned'''
    if isinstance(opendocx(TEST_FILE),lxml.etree._Element):
        pass
    else:
        assert False

def testmakeelement():
    '''Ensure custom elements get created'''
    testelement = makeelement('testname',attributes={'testattribute':'testvalue'},tagtext='testtagtext')
    assert testelement.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testname'
    assert testelement.attrib == {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}testattribute': 'testvalue'}
    assert testelement.text == 'testtagtext'

def testparagraph():
    '''Ensure paragraph creates p elements'''
    testpara = paragraph('paratext',style='BodyText')
    assert testpara.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'
    pass
    
def testtable():
    '''Ensure tables make sense'''
    testtable = table([['A1','A2'],['B1','B2'],['C1','C2']])
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    assert testtable.xpath('/ns0:tbl/ns0:tr[2]/ns0:tc[2]/ns0:p/ns0:r/ns0:t',namespaces={'ns0':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text == 'B2'

if __name__=='__main__':
    nose.main()