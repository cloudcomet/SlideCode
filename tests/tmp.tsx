import { useState, useRef, useEffect } from "react";
import Editor from "@monaco-editor/react";
import { Download, FileIcon, Loader2, Play, Moon, Sun, ChevronLeft, ChevronRight } from "lucide-react";
import PptxGenJS from "pptxgenjs";
import { Document, Page, pdfjs } from "react-pdf";
import "react-pdf/dist/Page/AnnotationLayer.css";
import "react-pdf/dist/Page/TextLayer.css";

// Setup PDF worker
pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

const DEFAULT_CODE = `// Write JavaScript to generate your PowerPoint presentation.
// A global 'pptx' instance (PptxGenJS) is already created for you.

// 1. Add a slide
const slide = pptx.addSlide();

// 2. Add text
slide.addText("Welcome to Code-to-PPTX", {
  x: 1.5,
  y: 1.5,
  w: 7,
  h: 1,
  fontSize: 32,
  bold: true,
  color: "363636",
  align: "center",
});

// 3. Add a shape
slide.addShape(pptx.ShapeType.rect, {
  x: 1.5,
  y: 2.8,
  w: 7,
  h: 1.5,
  fill: { color: "F1F1F1" },
  line: { color: "CCCCCC", width: 1 },
});

// 4. Add text inside the shape
slide.addText("Generate downloadable PPTX and PDF files instantly from JavaScript code.", {
  x: 1.5,
  y: 2.8,
  w: 7,
  h: 1.5,
  fontSize: 14,
  color: "666666",
  align: "center",
});

// The presentation is automatically generated and exported based on this object.
`;

export default function App() {
  const [code, setCode] = useState(DEFAULT_CODE);
  const [isGenerating, setIsGenerating] = useState(false);
  const [status, setStatus] = useState("Ready");
  const [error, setError] = useState<string | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [pptxBlobUrl, setPptxBlobUrl] = useState<string | null>(null);
  const [numPages, setNumPages] = useState<number | null>(null);
  const [pageNumber, setPageNumber] = useState(1);
  const [containerWidth, setContainerWidth] = useState<number>(0);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Clean up blob URLs to prevent memory leaks safely
  useEffect(() => {
    const url = previewUrl;
    return () => {
      if (url) {
        // Delay revoking the URL to ensure react-pdf worker has fully unmounted
        setTimeout(() => URL.revokeObjectURL(url), 10000);
      }
    };
  }, [previewUrl]);

  useEffect(() => {
    const url = pptxBlobUrl;
    return () => {
      if (url) {
        setTimeout(() => URL.revokeObjectURL(url), 10000);
      }
    };
  }, [pptxBlobUrl]);

  // Resize observer to make PDF responsive
  useEffect(() => {
    if (!containerRef.current) return;
    const resizeObserver = new ResizeObserver((entries) => {
      for (let entry of entries) {
        setContainerWidth(entry.contentRect.width);
      }
    });
    resizeObserver.observe(containerRef.current);
    return () => resizeObserver.disconnect();
  }, []);

  // Helper to execute code and get arraybuffer
  const generatePptxBuffer = async () => {
    try {
      const pptx = new PptxGenJS();
      
      // Make PptxGenJS available globally for advanced uses if needed
      (window as any).PptxGenJS = PptxGenJS;
      
      // Provide a mock require so code copied from Node/CommonJS examples works
      const customRequire = (moduleName: string) => {
        if (moduleName === 'pptxgenjs' || moduleName === 'PptxGenJS') {
          return ProxyPptxGenJS;
        }
        if ((window as any).require) {
          return (window as any).require(moduleName);
        }
        throw new Error(`Module '${moduleName}' not found.`);
      };
      
      let targetPptx = pptx;
      
      // Proxy the constructor to catch if they create a new instance
      const ProxyPptxGenJS = new Proxy(PptxGenJS, {
        construct(target: any, args: any[]) {
          const instance = new target(...args);
          targetPptx = instance;
          
          // Disable writeFile so it doesn't trigger a browser download before we convert to PDF
          instance.writeFile = function() {
             console.log("writeFile intercepted - download will be handled by the app.");
             return Promise.resolve("intercepted");
          };
          
          return instance;
        }
      });
      
      // Support basic import syntax
      let processedCode = code
        .replace(/import\s+([^{}\s]+)\s+from\s+['"](pptxgenjs|PptxGenJS)['"];?/g, "const $1 = require('$2');")
        .replace(/import\s+\*\s+as\s+([^{}\s]+)\s+from\s+['"](pptxgenjs|PptxGenJS)['"];?/g, "const $1 = require('$2');");
      
      // Execute the user's code inside a block to prevent variable redeclaration errors 
      // (like 'let pptx' shadowing the parameter)
      const exec = new Function("pptx", "require", "PptxGenJS", "{\n" + processedCode + "\n}");
      
      exec(targetPptx, customRequire, ProxyPptxGenJS);
      
      // Return the generated array buffer from whichever instance they used
      return await targetPptx.write({ outputType: "arraybuffer" }) as ArrayBuffer;
    } catch (err: any) {
      console.error(err);
      throw new Error(err.message || "Failed to execute your JavaScript code.");
    }
  };

  const handleRun = async () => {
    setIsGenerating(true);
    setStatus("Generating PPTX...");
    setError(null);
    setPreviewUrl(null);
    setPptxBlobUrl(null);
    setNumPages(null);
    
    try {
      const buffer = await generatePptxBuffer();
      const pptxBlob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
      setPptxBlobUrl(URL.createObjectURL(pptxBlob));
      
      setStatus("Generating PDF preview...");
      
      const formData = new FormData();
      formData.append("pptx", pptxBlob, "presentation.pptx");
      
      const response = await fetch("/api/convert-pdf", {
        method: "POST",
        body: formData,
      });
      
      if (!response.ok) {
        throw new Error("Failed to convert PPTX to PDF on the server.");
      }
      
      const pdfBlob = await response.blob();
      setPreviewUrl(URL.createObjectURL(pdfBlob));
      setStatus("Ready");
    } catch (err: any) {
      setError(err.message);
      setStatus("Error");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDownloadPptx = () => {
    if (!pptxBlobUrl) return;
    const a = document.createElement("a");
    a.href = pptxBlobUrl;
    a.download = "presentation.pptx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleDownloadPdf = () => {
    if (!previewUrl) return;
    const a = document.createElement("a");
    a.href = previewUrl;
    a.download = "presentation.pdf";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className={isDarkMode ? "dark" : ""}>
      <div className="flex flex-col h-screen bg-neutral-50 text-neutral-900 dark:bg-neutral-950 dark:text-neutral-100 font-sans transition-colors duration-200">
        {/* Header */}
        <header className="px-6 py-4 bg-white dark:bg-neutral-900 border-b border-neutral-200 dark:border-neutral-800 flex items-center justify-between transition-colors duration-200">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <div className="bg-indigo-600 p-2 rounded-lg text-white">
                <Play className="w-5 h-5" />
              </div>
              <h1 className="text-xl font-bold tracking-tight">CodeSlide</h1>
            </div>
            
            <div className="h-6 w-px bg-neutral-200 dark:bg-neutral-700 mx-2 transition-colors duration-200"></div>
            
            <button
              onClick={handleRun}
              disabled={isGenerating}
              className="flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white py-2 px-4 rounded-lg font-medium transition-colors shadow-sm text-sm"
            >
              {isGenerating ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Play className="w-4 h-4" fill="currentColor" />
              )}
              Run & Preview
            </button>
          </div>
          
          <div className="flex items-center gap-3">
            {error && (
              <span className="text-sm font-medium text-red-600 dark:text-red-400 truncate max-w-xs mr-2">
                {error}
              </span>
            )}
            {!error && status !== "Ready" && (
              <span className="text-sm font-medium text-neutral-500 dark:text-neutral-400 mr-2 flex items-center gap-2">
                <Loader2 className="w-3 h-3 animate-spin" />
                {status}
              </span>
            )}
            
            <button
              onClick={() => setIsDarkMode(!isDarkMode)}
              className="flex items-center justify-center p-2 rounded-lg bg-white dark:bg-neutral-800 hover:bg-neutral-50 dark:hover:bg-neutral-700 border border-neutral-200 dark:border-neutral-700 text-neutral-700 dark:text-neutral-300 transition-colors shadow-sm mr-2"
              title="Toggle Theme"
            >
              {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
            </button>
            
            <button
              onClick={handleDownloadPptx}
              disabled={!pptxBlobUrl || isGenerating}
              className="flex items-center justify-center gap-2 bg-white dark:bg-neutral-800 hover:bg-neutral-50 dark:hover:bg-neutral-700 border border-neutral-200 dark:border-neutral-700 disabled:opacity-50 text-neutral-700 dark:text-neutral-300 py-2 px-4 rounded-lg font-medium transition-colors shadow-sm text-sm"
              title="Download PPTX"
            >
              <FileIcon className="w-4 h-4" />
              <span className="hidden sm:inline">PPTX</span>
            </button>
  
            <button
              onClick={handleDownloadPdf}
              disabled={!previewUrl || isGenerating}
              className="flex items-center justify-center gap-2 bg-white dark:bg-neutral-800 hover:bg-neutral-50 dark:hover:bg-neutral-700 border border-neutral-200 dark:border-neutral-700 disabled:opacity-50 text-neutral-700 dark:text-neutral-300 py-2 px-4 rounded-lg font-medium transition-colors shadow-sm text-sm"
              title="Download PDF"
            >
              <Download className="w-4 h-4" />
              <span className="hidden sm:inline">PDF</span>
            </button>
          </div>
        </header>

      {/* Main Content */}
      <main className="flex-1 flex overflow-hidden transition-colors duration-200">
        {/* Left Side: Editor */}
        <div className="w-1/2 flex flex-col border-r border-neutral-200 dark:border-neutral-800 bg-white dark:bg-neutral-900 transition-colors duration-200">
          <div className="bg-neutral-50 dark:bg-neutral-800 px-4 py-2 text-xs font-semibold text-neutral-500 dark:text-neutral-400 uppercase tracking-wider border-b border-neutral-200 dark:border-neutral-700 flex justify-between items-center transition-colors duration-200">
            <span>JavaScript Editor</span>
            <a href="https://gitbrent.github.io/PptxGenJS/" target="_blank" rel="noreferrer" className="text-indigo-600 dark:text-indigo-400 hover:text-indigo-800 dark:hover:text-indigo-300 normal-case tracking-normal transition-colors">
              API Docs ↗
            </a>
          </div>
          <div className="flex-1 pt-4">
            <Editor
              height="100%"
              defaultLanguage="javascript"
              theme={isDarkMode ? "vs-dark" : "light"}
              value={code}
              onChange={(val) => setCode(val || "")}
              options={{
                minimap: { enabled: false },
                fontSize: 14,
                wordWrap: "on",
                scrollBeyondLastLine: false,
                lineNumbersMinChars: 3,
              }}
            />
          </div>
        </div>

        {/* Right Side: Preview */}
        <div className="w-1/2 flex flex-col bg-neutral-100 dark:bg-neutral-950 transition-colors duration-200">
          <div className="bg-neutral-50 dark:bg-neutral-800 px-4 py-2 text-xs font-semibold text-neutral-500 dark:text-neutral-400 uppercase tracking-wider border-b border-neutral-200 dark:border-neutral-700 transition-colors duration-200">
            Live Preview
          </div>
          
          <div className="flex-1 p-4 flex flex-col relative" ref={containerRef}>
            {previewUrl && !isGenerating ? (
              <div className="flex-1 bg-neutral-200/50 dark:bg-neutral-900/50 rounded-lg shadow-inner border border-neutral-200 dark:border-neutral-800 overflow-hidden relative flex flex-col transition-colors duration-200">
                <div className="flex-1 overflow-auto w-full flex justify-center py-4 custom-scrollbar">
                  <Document
                    file={previewUrl}
                    onLoadSuccess={({ numPages }) => {
                      setNumPages(numPages);
                      setPageNumber(1);
                    }}
                    loading={<div className="flex items-center justify-center p-8"><Loader2 className="w-8 h-8 animate-spin text-indigo-600 dark:text-indigo-400" /></div>}
                    className="max-w-full"
                  >
                    <Page 
                      pageNumber={pageNumber} 
                      width={containerWidth ? containerWidth - 40 : undefined} 
                      renderTextLayer={false}
                      renderAnnotationLayer={false}
                      className="shadow-md bg-white shrink-0"
                    />
                  </Document>
                </div>
                {numPages && numPages > 1 && (
                  <div className="bg-white dark:bg-neutral-900 border-t border-neutral-200 dark:border-neutral-800 w-full py-3 px-4 flex items-center justify-center gap-4 transition-colors duration-200">
                    <button 
                      onClick={() => setPageNumber(Math.max(1, pageNumber - 1))}
                      disabled={pageNumber <= 1}
                      className="p-1 rounded bg-neutral-100 dark:bg-neutral-800 hover:bg-neutral-200 dark:hover:bg-neutral-700 disabled:opacity-50 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                    <span className="text-sm font-medium dark:text-neutral-300">Page {pageNumber} of {numPages}</span>
                    <button 
                      onClick={() => setPageNumber(Math.max(1, Math.min(numPages, pageNumber + 1)))}
                      disabled={pageNumber >= numPages}
                      className="p-1 rounded bg-neutral-100 dark:bg-neutral-800 hover:bg-neutral-200 dark:hover:bg-neutral-700 disabled:opacity-50 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronRight className="w-5 h-5" />
                    </button>
                  </div>
                )}
              </div>
            ) : (
              <div className="flex-1 border-2 border-dashed border-neutral-300 dark:border-neutral-700 rounded-xl flex flex-col items-center justify-center text-neutral-400 dark:text-neutral-500 p-8 text-center bg-neutral-50/50 dark:bg-neutral-900/50 transition-colors duration-200">
                <Play className="w-12 h-12 mb-4 text-neutral-300 dark:text-neutral-600" />
                <p className="text-lg font-medium text-neutral-600 dark:text-neutral-400 mb-2">No Preview Available</p>
                <p className="max-w-xs text-sm">
                  Click the <strong>Run & Preview</strong> button in the header to generate and view your presentation.
                </p>
              </div>
            )}
          </div>
        </div>
      </main>
      </div>
    </div>
  );
}
