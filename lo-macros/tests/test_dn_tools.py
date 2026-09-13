import sys
sys.path.insert(0, "..")
from vistatype import dn_tools

xml_sample = '''<dtbook><level1><p>Some text<pagenum id="p12">12</pagenum> more text.</p>
<p>Continuing<pagenum id="p13" value="A13"/> here.</p></level1></dtbook>'''
tagged = dn_tools.add_pg_tags_to_xml(xml_sample)
print(tagged)
assert "<pagenum id=\"p12\">$pg12</pagenum>" in tagged
assert 'value="$pgA13"' in tagged

text_sample = '''"Oh, that's all very fine to SAY, Tom Sawyer, but how in the nation are
these fellows going to be ransomed if we don't know how to do it to them?
--that's the thing I want to get at. Now, what do you reckon it is?"

"Well, I don't know.  But per'aps if we keep them till they're ransomed,
it means that we keep them till they're dead."

"Now, that's something LIKE. That'll answer. Why couldn't you said that
before? We'll keep them till they're ransomed to death; and a bothersome
Lot they'll be, too--eating up everything, and always trying to get loose."
'''
fixed = dn_tools.fix_text_file_paragraphs(text_sample)
print("---FIXED---")
print(fixed)
paragraphs = [p for p in fixed.split("\n") if p]
assert len(paragraphs) == 3
assert paragraphs[0].startswith('"Oh, that\'s all very fine to SAY, Tom Sawyer,')
assert "ransomed if we don't know" in paragraphs[0]
print("DN TOOLS TEST PASSED")
