const pptxgen = require("pptxgenjs");
const { convert } = require('pptx-to-pdf');
const fs = require('fs');

async function run() {
const pptx = new pptxgen();

pptx.layout = "LAYOUT_WIDE";

const THEME = {
  colors: {
    primary: "1E2761",
    dark: "141A3D",
    accent: "C9962E",
    light: "EAF0FB",
    border: "CADCFC",
    muted: "5B6B8C",
    white: "FFFFFF",
    success: "3B8F3B",
    danger: "B23A48"
  },

  fonts: {
    heading: "Cambria",
    body: "Calibri"
  }
};

function makeData(size) {

  return Array.from(
    { length: size },
    (_, index) => ({
      id: index + 1,
      title: `Research Item ${index + 1}`,
      year: 2020 + (index % 7),
      score: Math.round(
        50 +
        Math.sin(index / 4) * 30 +
        Math.random() * 20
      ),
      tags: [
        "AI",
        "RAG",
        "Agentic",
        "Multimodal"
      ].slice(0, (index % 4) + 1),

      metadata: {
        source: `Dataset-${index % 10}`,
        nested: {
          level1: {
            level2: {
              value: Math.random()
            }
          }
        }
      }
    })
  );
}

function addCard(slide, item, index) {

  const x =
    0.6 + (index % 3) * 4.15;

  const y =
    1.7 + Math.floor(index / 3) * 1.65;

  slide.addShape(
    pptx.ShapeType.roundRect,
    {
      x,
      y,
      w: 3.8,
      h: 1.35,
      rectRadius: 0.06,
      fill: {
        color:
          item.score >= 80
            ? THEME.colors.light
            : "F2F2F2"
      },
      line: {
        color: THEME.colors.border,
        width: 0.8
      }
    }
  );

  slide.addText(
    item.title,
    {
      x: x + 0.2,
      y: y + 0.15,
      w: 3.3,
      h: 0.3,
      fontFace: THEME.fonts.heading,
      fontSize: 12,
      bold: true,
      color: THEME.colors.primary
    }
  );

  slide.addText(
    [
      {
        text: `Score: ${item.score}`,
        options: {
          bold: true,
          color:
            item.score >= 80
              ? THEME.colors.success
              : THEME.colors.danger
        }
      },
      {
        text: `\nYear: ${item.year}`,
        options: {
          color: THEME.colors.muted
        }
      },
      {
        text: `\nTags: ${item.tags.join(", ")}`,
        options: {
          color: THEME.colors.muted
        }
      }
    ],
    {
      x: x + 0.2,
      y: y + 0.55,
      w: 3.3,
      h: 0.65,
      fontSize: 9
    }
  );
}

const data = makeData(500);

const average =
  data.reduce(
    (sum, item) =>
      sum + item.score,
    0
  ) / data.length;

for (let page = 0; page < 10; page++) {

  const slide = pptx.addSlide();

  slide.background = {
    color:
      page % 2 === 0
        ? THEME.colors.white
        : "F7F9FC"
  };

  slide.addText(
    `Comprehensive Stress Test — Page ${page + 1}`,
    {
      x: 0.6,
      y: 0.4,
      w: 11,
      h: 0.5,
      fontFace: THEME.fonts.heading,
      fontSize: 25,
      bold: true,
      color: THEME.colors.primary
    }
  );

  slide.addText(
    `Dataset size: ${data.length} | Average score: ${average.toFixed(2)}`,
    {
      x: 0.6,
      y: 0.95,
      w: 11,
      h: 0.3,
      fontSize: 10,
      color: THEME.colors.muted
    }
  );

  const pageData =
    data.slice(
      page * 12,
      page * 12 + 12
    );

  pageData.forEach(
    (item, index) =>
      addCard(
        slide,
        item,
        index
      )
  );

  slide.addChart(
    pptx.ChartType.line,
    [
      {
        name: "Score",
        labels: pageData.map(
          item => String(item.id)
        ),
        values: pageData.map(
          item => item.score
        )
      }
    ],
    {
      x: 0.7,
      y: 5.25,
      w: 5.8,
      h: 1.4,
      showLegend: false
    }
  );

  slide.addText(
    [
      {
        text: "Status: ",
        options: {
          bold: true
        }
      },
      {
        text:
          average >= 75
            ? "PASS"
            : "REVIEW",
        options: {
          bold: true,
          color:
            average >= 75
              ? THEME.colors.success
              : THEME.colors.danger
        }
      }
    ],
    {
      x: 7,
      y: 5.5,
      w: 4,
      h: 0.5,
      fontSize: 18
    }
  );
}

const buffer = await pptx.write({ outputType: 'nodebuffer' });
try {
   const pdf = await convert(buffer);
   fs.writeFileSync('test20.pdf', pdf);
   console.log("PDF generated");
} catch (e) {
   console.error("Convert error:", e);
}
}
run();
