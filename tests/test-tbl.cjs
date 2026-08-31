const { XMLParser } = require('fast-xml-parser');
const xml = `
<p:graphicFrame>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="2834640"/>
        </a:tblGrid>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
`;
const options = {
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name) => name === 'a:tbl',
};
const parser = new XMLParser(options);
const doc = parser.parse(xml);
const gf = doc['p:graphicFrame'];
const graphic = gf['a:graphic'];
const graphicData = graphic['a:graphicData'];
const tbl = graphicData['a:tbl'];
console.log("tbl is array?", Array.isArray(tbl));
console.log("tblGrid property:", tbl['a:tblGrid'] !== undefined);
console.log("tbl[0] tblGrid property:", tbl[0]['a:tblGrid'] !== undefined);
