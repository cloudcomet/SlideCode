const { XMLParser } = require("fast-xml-parser");
const xml = `<p:graphicFrame>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tr h="896112"><a:tc><a:txBody><a:p><a:r><a:t>TEST1</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => ['a:tr', 'a:tc', 'a:tbl', 'p:graphicFrame', 'a:r', 'a:br', 'a:p', 'a:gridCol', 'p:sp', 'p:pic'].includes(name)
});
const doc = parser.parse(xml);
console.log(JSON.stringify(doc, null, 2));
