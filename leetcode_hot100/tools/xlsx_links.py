"""Add native OOXML hyperlinks for a capability missing in artifact-tool.

Workbook values, calculations, and formatting are authored by artifact-tool.
This utility touches only worksheet hyperlink elements and their relationships.
"""
import json,sys,zipfile,posixpath,os
from pathlib import Path
from xml.etree import ElementTree as ET

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
DOCREL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('',NS)
ET.register_namespace('r',DOCREL)

def patch_links(workbook,links,hide_original=False):
    workbook=Path(workbook).resolve()
    with zipfile.ZipFile(workbook) as z:
        contents={i.filename:z.read(i.filename) for i in z.infolist()}
    wb=ET.fromstring(contents['xl/workbook.xml'])
    rels=ET.fromstring(contents['xl/_rels/workbook.xml.rels'])
    targets={e.attrib['Id']:e.attrib['Target'] for e in rels}
    byname={e.attrib['name']:targets[e.attrib[f'{{{DOCREL}}}id']] for e in wb.find(f'{{{NS}}}sheets')}
    for sheet in {x['sheet'] for x in links}:
        target=byname[sheet]
        part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
        relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
        xml=ET.fromstring(contents[part])
        hyperlinks=xml.find(f'{{{NS}}}hyperlinks')
        if hyperlinks is None:
            hyperlinks=ET.Element(f'{{{NS}}}hyperlinks')
            # Schema position: after validations, before print/page/drawing/table elements.
            later={'printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
            at=next((i for i,e in enumerate(xml) if e.tag.rsplit('}',1)[-1] in later),len(xml))
            xml.insert(at,hyperlinks)
        relxml=ET.fromstring(contents[relpart]) if relpart in contents else ET.Element(f'{{{REL}}}Relationships')
        used={e.attrib['Id'] for e in relxml}
        for n,l in enumerate((l for l in links if l['sheet']==sheet),1):
            attrs={'ref':l['cell'],'display':l['label']}
            if l.get('tooltip'):attrs['tooltip']=l['tooltip']
            if l['target'].startswith('#'):
                attrs['location']=l['target'][1:]
            else:
                rid=f'codexHyperlink{n}'
                while rid in used: rid+='x'
                used.add(rid);attrs[f'{{{DOCREL}}}id']=rid
                ET.SubElement(relxml,f'{{{REL}}}Relationship',{'Id':rid,'Type':DOCREL+'/hyperlink','Target':l['target'],'TargetMode':'External'})
            ET.SubElement(hyperlinks,f'{{{NS}}}hyperlink',attrs)
        if hide_original:
            cols=xml.find(f'{{{NS}}}cols')
            if cols is not None:
                for col in cols:
                    if int(col.attrib['min'])>=6 and int(col.attrib['max'])<=7:col.set('hidden','1')
        contents[part]=ET.tostring(xml,encoding='utf-8',xml_declaration=True)
        contents[relpart]=ET.tostring(relxml,encoding='utf-8',xml_declaration=True)
    temporary=workbook.with_suffix('.links-patched.xlsx')
    with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in contents.items():z.writestr(name,data)
    os.replace(temporary,workbook)
    print(f'Added {len(links)} native hyperlinks.')

if __name__=='__main__':
    patch_links(sys.argv[1],json.loads(Path(sys.argv[2]).read_text(encoding='utf-8')),len(sys.argv)>3)
