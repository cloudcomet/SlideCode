const pptxgen = require("pptxgenjs");

// ---------- Palette ("Midnight Executive" classic academic/professional) ----------
const NAVY = "1E2761";
const NAVY_DARK = "141A3D";
const ICE = "CADCFC";
const ICE_LIGHT = "EAF0FB";
const WHITE = "FFFFFF";
const GOLD = "C9962E";
const TEXT_DARK = "1E2761";
const TEXT_MUTED = "5B6B8C";
const TEXT_MUTED_LIGHT = "AAB8DA";

const HEADER_FONT = "Cambria";
const BODY_FONT = "Calibri";

// Official UN SDG colors
const SDG = {
  3: "4C9F38",
  4: "C5192D",
  9: "FD6925",
  10: "DD1367",
  16: "00689D",
};

let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5

const PW = 13.333, PH = 7.5;

function footer(slide, num) {
  slide.addText("MULTIMODAL AGENTIC RAG  |  B.E. SEMINAR", {
    x: 0.6, y: 7.12, w: 6, h: 0.3, fontSize: 8.5, color: TEXT_MUTED,
    fontFace: BODY_FONT, isTextBox: true, charSpacing: 1,
  });
  slide.addText(String(num), {
    x: 12.4, y: 7.12, w: 0.4, h: 0.3, fontSize: 9, color: TEXT_MUTED,
    fontFace: BODY_FONT, align: "right", isTextBox: true,
  });
}

function eyebrow(slide, text, opts = {}) {
  slide.addText(text.toUpperCase(), Object.assign({
    x: 0.6, y: 0.42, w: 10, h: 0.32, fontSize: 12.5, bold: true,
    color: GOLD, fontFace: BODY_FONT, charSpacing: 2, isTextBox: true,
  }, opts));
}

function title(slide, text, opts = {}) {
  slide.addText(text, Object.assign({
    x: 0.6, y: 0.75, w: 12.1, h: 0.85, fontSize: 30, bold: true,
    color: TEXT_DARK, fontFace: HEADER_FONT, isTextBox: true,
  }, opts));
}

function iconCircle(slide, x, y, d, label, bg = NAVY, fg = WHITE, fs = 16) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  slide.addText(label, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: fs, bold: true, color: fg, fontFace: BODY_FONT, isTextBox: true, margin: 0 });
}

function card(slide, x, y, w, h, fillColor = ICE_LIGHT) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fillColor }, line: { type: "none" } });
}

// ============================================================ SLIDE 5: RAG LANDSCAPE (TYPES TABLE)
{
  let s = pres.addSlide();
  eyebrow(s, "Landscape of Approaches");
  title(s, "Types of RAG — From Naive to Multimodal Agentic");

  const rows = [
    ["Type", "Core Idea", "Example"],
    ["Naive RAG", "Single retrieve \u2192 generate pass, flat vector similarity", "Lewis et al., 2020"],
    ["Advanced / Modular RAG", "Adds query rewriting, reranking, filtering stages", "Improves precision"],
    ["Graph RAG", "Knowledge-graph traversal instead of pure vector search", "Microsoft GraphRAG, LightRAG"],
    ["Agentic RAG", "Agent decides when/what/whether to retrieve, can re-plan", "Singh et al. survey"],
    ["Multimodal RAG", "Retrieval spans text + image + table, not just text chunks", "MuRAG, UniRAG"],
    ["Multimodal Agentic RAG", "Agents plan & reason across modalities (this seminar)", "HM-RAG, MultiRAG, CollEX"],
    ["Vectorless RAG", "Skips embeddings; agent/LLM searches raw text or tools directly", "Emerging 2025-26 trend"],
    ["Cache-Augmented Gen. (CAG)", "Preloads knowledge into context/KV-cache ahead of query time", "Small, static corpora"],
  ];

  const tblRows = rows.map((r, i) => r.map((c, j) => ({
    text: c,
    options: {
      bold: i === 0,
      color: i === 0 ? WHITE : (j === 0 ? TEXT_DARK : TEXT_MUTED),
      fill: i === 0 ? NAVY : (i % 2 === 0 ? ICE_LIGHT : WHITE),
      fontSize: i === 0 ? 11.5 : 10.5,
      fontFace: BODY_FONT,
      valign: "middle",
      align: j === 0 ? "left" : "left",
    },
  })));

  s.addTable(tblRows, {
    x: 0.6, y: 1.75, w: 12.15, h: 5.1,
    colW: [3.1, 6.55, 2.5],
    border: { type: "solid", color: ICE, pt: 0.5 },
    autoPage: false,
    rowH: 0.55,
  });

  footer(s, 5);
}

const litPapers = [
  { p: "HM-RAG", v: "ACM Multimedia, 2025", f: "Hierarchical multi-agent multimodal retrieval", k: "Orchestrator + modality-specific sub-agents improve cross-modal retrieval precision over single-agent RAG." },
  { p: "MultiRAG", v: "Springer LNCS (AI 2025), 2025", f: "Multi-source, multi-modal RAG for scientific research", k: "Dynamic modal weighting preserves figure-text links, boosting accuracy on hybrid queries." },
  { p: "Patho-AgenticRAG", v: "Preprint, 2025", f: "Domain-specialized agentic RAG for pathology VLMs", k: "RL-based task decomposition & multi-turn search reduce diagnostic hallucination in clinical settings." },
  { p: "CollEX", v: "Preprint, 2025", f: "Interactive agentic RAG for scientific collections", k: "User-in-the-loop exploration across images, code, databases, audio & video broadens usable evidence." },
  { p: "Agentic RAG (IJCTT)", v: "Journal, 2025", f: "AI-driven information retrieval advancement", k: "Frames agentic RAG's advantage over static RAG for dynamic, real-world retrieval tasks." },
  { p: "UniRAG", v: "IEEE Access, 2024", f: "Unified retrieval + generation for multimodal QA", k: "Unifying pre-trained LMs for retrieval & generation improves multimodal QA accuracy." },
  { p: "Multimodal RAG + KG", v: "Emerging Sci. J., 2025", f: "Knowledge-graph-augmented multimodal RAG", k: "Linking multimodal entities via a knowledge graph improves retrieval relevance and cuts noise." },
  { p: "Hallucination Survey", v: "ACM TOIS (Q1), 2025", f: "Taxonomy of hallucination causes & mitigation in LLMs", k: "Positions retrieval-based grounding as a primary lever for hallucination mitigation." },
  { p: "LLM Architectures Survey", v: "IEEE Access, 2024", f: "Benchmarking LLM architecture trends", k: "Identifies efficiency-accuracy trade-offs relevant to choosing an agent's backbone model." },
  { p: "Agentic RAG Survey", v: "Preprint (Singh et al.), 2025", f: "Formal taxonomy of agentic RAG", k: "Defines reflection, planning, tool-use & multi-agent collaboration as core agentic RAG patterns." },
  { p: "Ask in Any Modality", v: "Preprint, 2025/26", f: "Comprehensive survey of multimodal RAG", k: "Maps modality-specific retrieval techniques and their respective performance gaps." },
  { p: "Survey of Multimodal RAG", v: "Preprint (Mei et al.), 2025", f: "Broad multimodal RAG survey", k: "Classifies multimodal RAG pipelines by retrieval strategy and fusion method." },
  { p: "RAG & Understanding in Vision", v: "Preprint, 2025", f: "Vision-centric RAG survey and outlook", k: "Positions vision-grounded retrieval as a needed complement to text-only RAG." },
  { p: "Scaling Beyond Context", v: "Preprint, 2025", f: "Multimodal RAG for document understanding", k: "Shows RAG-based scaling can outperform long-context-only approaches for document-heavy tasks." },
  { p: "mRAG", v: "Preprint, 2025", f: "Design space of multimodal RAG", k: "Elucidates architectural choices (retrieval granularity, fusion timing) shaping multimodal RAG performance." },
];

function litReviewSlide(subset, startNum, part, footerNum) {
  let s = pres.addSlide();
  eyebrow(s, "Literature Review");
  title(s, `15 Papers Reviewed \u2014 Summary & Key Findings (${part})`);

  const rows = [["#", "Paper / Venue", "Primary Focus", "Key Finding"]];
  subset.forEach((it, i) => {
    rows.push([String(startNum + i), "", it.f, it.k]);
  });

  const tblRows = rows.map((r, i) => r.map((c, j) => ({
    text: c,
    options: {
      bold: i === 0,
      color: i === 0 ? WHITE : (j === 1 ? TEXT_DARK : TEXT_MUTED),
      fill: i === 0 ? NAVY : (i % 2 === 0 ? ICE_LIGHT : WHITE),
      fontSize: i === 0 ? 11.5 : 10,
      fontFace: BODY_FONT,
      valign: "middle",
      align: j === 0 ? "center" : "left",
    },
  })));
  // rebuild column 1 (Paper/Venue) as rich text for each data row: bold name + muted italic venue
  for (let i = 1; i < tblRows.length; i++) {
    tblRows[i][1] = {
      text: [
        { text: subset[i - 1].p + "\n", options: { bold: true, color: TEXT_DARK, fontSize: 10.5 } },
        { text: subset[i - 1].v, options: { italic: true, color: TEXT_MUTED, fontSize: 8.5 } },
      ],
      options: { fill: tblRows[i][0].options.fill, valign: "middle" },
    };
  }

  s.addTable(tblRows, {
    x: 0.6, y: 1.5, w: 12.15, h: 5.45,
    colW: [0.55, 3.1, 3.55, 4.95],
    border: { type: "solid", color: ICE, pt: 0.5 },
    autoPage: false,
    rowH: [0.55, 0.98, 0.98, 0.98, 0.98, 0.98],
  });

  footer(s, footerNum);
}

litReviewSlide(litPapers.slice(0, 5), 1, "1/3", 9);
litReviewSlide(litPapers.slice(5, 10), 6, "2/3", 10);
litReviewSlide(litPapers.slice(10, 15), 11, "3/3", 11);

pres.writeFile({ fileName: "full-test.pptx" }).then(() => {
  const { convert } = require("pptx-to-pdf");
  const fs = require("fs");
  const buffer = fs.readFileSync("full-test.pptx");
  convert(buffer).then(pdfBuf => {
     fs.writeFileSync("full-test.pdf", pdfBuf);
     const pdftotext = require("child_process").execSync;
     console.log("PDF text length:", pdftotext("pdftotext full-test.pdf -").toString().length);
  });
});
