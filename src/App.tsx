import { useState, useRef, useEffect } from "react";
import Editor from "@monaco-editor/react";
import { Download, FileIcon, Loader2, Play, Moon, Sun, ChevronLeft, ChevronRight, ChevronFirst, ChevronLast, Sparkles, Cloud, User, LogOut } from "lucide-react";
import { useGoogleLogin, googleLogout } from '@react-oauth/google';
import PptxGenJS from "pptxgenjs";
import { Document, Page, pdfjs } from "react-pdf";
import Markdown from "react-markdown";
import "react-pdf/dist/Page/AnnotationLayer.css";
import "react-pdf/dist/Page/TextLayer.css";
// @ts-ignore - Vite specific import
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

// Setup PDF worker
pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const pdfOptions = {
  cMapUrl: `https://unpkg.com/pdfjs-dist@${pdfjs.version}/cmaps/`,
  cMapPacked: true,
  standardFontDataUrl: `https://unpkg.com/pdfjs-dist@${pdfjs.version}/standard_fonts/`,
};

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
  y: 2.75,
  w: 7,
  h: 0.5,
  fill: { color: "0088CC" },
});

// Note: Do not call pptx.writeFile() here.
// The app automatically extracts and builds the presentation.
`;

interface UserProfile {
  email: string;
  name: string;
  picture: string;
  access_token: string;
}

export default function App() {
  const [code, setCode] = useState(() => {
    const saved = localStorage.getItem("slidecode-saved-code");
    return saved !== null ? saved : DEFAULT_CODE;
  });
  const [isGenerating, setIsGenerating] = useState(false);
  const [status, setStatus] = useState<string>("Ready");
  const [error, setError] = useState<string | null>(null);
  const [pdfError, setPdfError] = useState<string | null>(null);
  
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [pptxBlobUrl, setPptxBlobUrl] = useState<string | null>(null);
  
  const [numPages, setNumPages] = useState<number | null>(null);
  const [pageNumber, setPageNumber] = useState(1);
  const [containerWidth, setContainerWidth] = useState<number>(0);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [explanation, setExplanation] = useState<string | null>(null);
  const [isExplaining, setIsExplaining] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [driveUrl, setDriveUrl] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  const [user, setUser] = useState<UserProfile | null>(null);
  const [isProfileOpen, setIsProfileOpen] = useState(false);
  const pendingUploadRef = useRef(false);
  const profileRef = useRef<HTMLDivElement>(null);

  // Auto-save code to local storage
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      localStorage.setItem("slidecode-saved-code", code);
    }, 2000);
    return () => clearTimeout(timeoutId);
  }, [code]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (profileRef.current && !profileRef.current.contains(event.target as Node)) {
        setIsProfileOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const login = useGoogleLogin({
    onSuccess: async (tokenResponse) => {
      try {
        const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
          headers: { Authorization: `Bearer ${tokenResponse.access_token}` }
        });
        const userInfo = await res.json();
        const newUser = {
          email: userInfo.email,
          name: userInfo.name,
          picture: userInfo.picture,
          access_token: tokenResponse.access_token
        };
        setUser(newUser);
        
        if (pendingUploadRef.current) {
          pendingUploadRef.current = false;
          await uploadToDrive(tokenResponse.access_token);
        }
      } catch (err) {
        console.error("Failed to fetch user info", err);
      }
    },
    scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email',
    onError: (error) => {
      setError("Google Login failed");
      pendingUploadRef.current = false;
    }
  });

  const handleLogout = () => {
    googleLogout();
    setUser(null);
    setDriveUrl(null);
    setIsProfileOpen(false);
  };

  const uploadToDrive = async (accessToken: string) => {
    if (!pptxBlobUrl) return;
    setIsUploading(true);
    setError(null);
    try {
      const response = await fetch(pptxBlobUrl);
      const blob = await response.blob();
      
      const metadata = {
        name: 'presentation.pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      };
      
      const form = new FormData();
      form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
      form.append('file', blob);
      
      const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        body: form
      });
      
      if (!res.ok) {
        const errorData = await res.json();
        throw new Error("Drive Upload Error: " + (errorData.error?.message || res.statusText));
      }
      
      const data = await res.json();
      setDriveUrl(data.webViewLink);
    } catch (e: any) {
      setError("Failed to upload to Drive: " + e.message);
    } finally {
      setIsUploading(false);
    }
  };

  // Remove these useEffects that revoke blobs, as Strict Mode causes them to revoke prematurely.
  // We will handle revocation when generating a new presentation instead.

  // Resize observer to make PDF responsive
  useEffect(() => {
    if (!containerRef.current) return;
    const resizeObserver = new ResizeObserver((entries) => {
      window.requestAnimationFrame(() => {
        if (!Array.isArray(entries) || !entries.length) return;
        setContainerWidth(entries[0].contentRect.width);
      });
    });
    resizeObserver.observe(containerRef.current);
    return () => resizeObserver.disconnect();
  }, []);

  // Helper to execute code and get arraybuffer
  const generatePptxBuffer = async (): Promise<ArrayBuffer> => {
    return new Promise(async (resolve, reject) => {
      try {
        const pptx = new PptxGenJS();
        let targetPptx = pptx;
        let resolved = false;

        const onComplete = async () => {
          if (resolved) return;
          resolved = true;
          try {
            if (!(targetPptx as any).slides || (targetPptx as any).slides.length === 0) {
              reject(new Error("No slides to render. Please add at least one slide using pptx.addSlide()."));
              return;
            }
            const buffer = await targetPptx.write({ outputType: 'arraybuffer' }) as ArrayBuffer;
            resolve(buffer);
          } catch (e) {
            reject(e);
          }
        };

        const interceptOutput = async () => {
          console.warn("pptx file output intercepted. Generating preview...");
          await onComplete();
          return Promise.resolve("");
        };
        
        targetPptx.writeFile = interceptOutput;
        (targetPptx as any).save = interceptOutput;

        // Proxy the constructor to catch if they create a new instance
        const ProxyPptxGenJS = new Proxy(PptxGenJS, {
          construct(target: any, args: any[]) {
            const instance = new target(...args);
            instance.writeFile = interceptOutput;
            (instance as any).save = interceptOutput;
            targetPptx = instance;
            return instance;
          }
        });

        // Make PptxGenJS available globally for advanced uses if needed
        (window as any).PptxGenJS = ProxyPptxGenJS;
        
        // We must provide a mock require for libraries that the script might assume exist
        const mockRequire = (moduleName: string) => {
          if (moduleName === 'pptxgenjs') return ProxyPptxGenJS;
          if ((window as any).require) {
            return (window as any).require(moduleName);
          }
          throw new Error(`Module '${moduleName}' not found.`);
        };
        
        // Execute user code
        // We use AsyncFunction to allow top-level await if needed
        const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
        const executeFn = new AsyncFunction('require', 'PptxGenJS', code);
        
        // Wait for script execution (using global pptx fallback to prevent SyntaxError if user declares const pptx)
        (window as any).pptx = pptx;
        await executeFn(mockRequire, ProxyPptxGenJS);
        delete (window as any).pptx;

        // If the script finished successfully but didn't call writeFile, 
        // give it a brief moment for dangling microtasks to settle, then generate.
        if (!resolved) {
          setTimeout(() => {
            if (!resolved) {
              onComplete();
            }
          }, 100);
        }
      } catch (err: any) {
        reject(new Error("Syntax or Execution Error: " + err.message));
      }
    });
  };

  const handleRun = async () => {
    setIsGenerating(true);
    setStatus("Generating PPTX...");
    setError(null);

    if (previewUrl) URL.revokeObjectURL(previewUrl);
    if (pptxBlobUrl) URL.revokeObjectURL(pptxBlobUrl);

    setPreviewUrl(null);
    setPptxBlobUrl(null);
    setPdfError(null);
    setNumPages(null);
    setExplanation(null);
    setDriveUrl(null);
    
    try {
      const buffer = await generatePptxBuffer();
      const pptxBlob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
      const currentPptxUrl = URL.createObjectURL(pptxBlob);
      setPptxBlobUrl(currentPptxUrl);
      
      setStatus("Generating PDF preview...");
      
      const formData = new FormData();
      formData.append("pptx", pptxBlob, "presentation.pptx");
      
      try {
        const response = await fetch("/api/convert-pdf", {
          method: "POST",
          body: formData,
        });
        
        if (!response.ok) {
          let errStr = "Failed to convert PPTX to PDF on the server.";
          try {
            const errData = await response.json();
            if (errData.error) errStr = errData.error;
          } catch (e) {}
          throw new Error(errStr);
        }
        
        const arrayBuffer = await response.arrayBuffer();
        const bytes = new Uint8Array(arrayBuffer);
        if (bytes.length < 5 || String.fromCharCode(...bytes.slice(0, 5)) !== "%PDF-") {
           const text = new TextDecoder().decode(bytes.slice(0, 150));
           throw new Error("Server returned invalid PDF data. Received: " + text);
        }

        const newPdfBlob = new Blob([arrayBuffer], { type: "application/pdf" });
        setPreviewUrl(URL.createObjectURL(newPdfBlob));
        setStatus("Ready");
      } catch (pdfErr: any) {
        setPdfError(pdfErr.message);
        setStatus("PDF Error");
      }
    } catch (err: any) {
      setError(err.message);
      setStatus("Error");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleRetryPdf = async () => {
    if (!pptxBlobUrl) return;
    setIsGenerating(true);
    setStatus("Retrying PDF preview...");
    setPdfError(null);
    
    try {
      const response = await fetch(pptxBlobUrl);
      const pptxBlob = await response.blob();
      
      const formData = new FormData();
      formData.append("pptx", pptxBlob, "presentation.pptx");
      
      const convertResponse = await fetch("/api/convert-pdf", {
        method: "POST",
        body: formData,
      });
      
      if (!convertResponse.ok) {
        let errStr = "Failed to convert PPTX to PDF on the server.";
        try {
          const errData = await convertResponse.json();
          if (errData.error) errStr = errData.error;
        } catch (e) {}
        throw new Error(errStr);
      }
      
      const arrayBuffer = await convertResponse.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      if (bytes.length < 5 || String.fromCharCode(...bytes.slice(0, 5)) !== "%PDF-") {
         const text = new TextDecoder().decode(bytes.slice(0, 150));
         throw new Error("Server returned invalid PDF data. Received: " + text);
      }

      const newPdfBlob = new Blob([arrayBuffer], { type: "application/pdf" });
      setPreviewUrl(URL.createObjectURL(newPdfBlob));
      setStatus("Ready");
    } catch (pdfErr: any) {
      setPdfError(pdfErr.message);
      setStatus("PDF Error");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleExplainError = async () => {
    if (!error) return;
    setIsExplaining(true);
    setExplanation(null);
    try {
      const response = await fetch("/api/explain-error", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ code, error }),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || data.explanation || "Failed to get explanation.");
      }
      setExplanation(data.explanation);
    } catch (err: any) {
      setExplanation("Error getting explanation: " + err.message);
    } finally {
      setIsExplaining(false);
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
    <div className={`min-h-screen w-full flex flex-col font-sans transition-colors duration-200 ${isDarkMode ? 'dark' : ''}`}>
      <div className="h-[100dvh] flex flex-col bg-neutral-50 dark:bg-[#0a0a0a] text-neutral-900 dark:text-neutral-50 overflow-hidden bg-dot-pattern">
        {/* Header */}
        <header className="flex-none px-4 md:px-6 py-3 md:py-4 flex justify-between items-center z-50">
          <div className="flex items-center gap-3 md:gap-6 shrink-0">
            <div className="flex items-center gap-2 md:gap-3">
              <div className="w-7 h-7 md:w-8 md:h-8 bg-neutral-900 dark:bg-white rounded-lg flex items-center justify-center shadow-sm shrink-0">
                <FileIcon className="w-3.5 h-3.5 md:w-4 md:h-4 text-white dark:text-neutral-900" />
              </div>
              <div className="hidden sm:block">
                <h1 className="text-sm font-semibold tracking-tight text-neutral-900 dark:text-neutral-100">SlideCode</h1>
                <p className="text-[10px] md:text-[11px] font-medium text-neutral-500 dark:text-neutral-400 uppercase tracking-widest hidden md:block">Generator Studio</p>
              </div>
            </div>
          </div>
          
          <div className="flex items-center gap-2 shrink-0 ml-4">
            <button type="button"
              onClick={() => setIsDarkMode(!isDarkMode)}
              className="flex items-center justify-center p-2 rounded-full bg-white dark:bg-neutral-900 hover:bg-neutral-100 dark:hover:bg-neutral-800 border border-neutral-200 dark:border-neutral-800 text-neutral-600 dark:text-neutral-400 transition-colors shadow-sm"
              title="Toggle Theme"
            >
              {isDarkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
            
            <div className="h-4 w-px bg-neutral-300 dark:bg-neutral-800 mx-2"></div>

            {/* Profile Dropdown */}
            <div className="relative" ref={profileRef}>
              {user ? (
                <button 
                  onClick={() => setIsProfileOpen(!isProfileOpen)}
                  className="w-8 h-8 rounded-full overflow-hidden border border-neutral-200 dark:border-neutral-700 hover:ring-2 hover:ring-emerald-500 transition-all flex-shrink-0 bg-neutral-100 dark:bg-neutral-800 focus:outline-none"
                >
                  {user.picture ? (
                    <img src={user.picture} alt="Profile" className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                  ) : (
                    <div className="w-full h-full bg-emerald-100 dark:bg-emerald-900/50 text-emerald-600 dark:text-emerald-400 flex items-center justify-center font-bold text-xs uppercase">
                      {user.name ? user.name.charAt(0) : user.email.charAt(0)}
                    </div>
                  )}
                </button>
              ) : (
                <button
                  onClick={() => login()}
                  className="flex items-center justify-center gap-1.5 md:gap-2 bg-neutral-100 dark:bg-neutral-800 hover:bg-neutral-200 dark:hover:bg-neutral-700 text-neutral-700 dark:text-neutral-300 py-1.5 px-2 md:px-3 rounded-md font-medium transition-colors shadow-sm text-[11px] md:text-xs border border-transparent dark:border-neutral-700/50"
                  title="Sign In"
                >
                  <User className="w-3.5 h-3.5" />
                  <span className="hidden sm:inline">Sign In</span>
                </button>
              )}

              {isProfileOpen && user && (
                <div className="absolute right-0 mt-3 w-56 bg-white dark:bg-neutral-900 rounded-lg shadow-lg border border-neutral-200 dark:border-neutral-800 overflow-hidden z-50 animate-in fade-in slide-in-from-top-2 duration-200">
                  <div className="px-4 py-3 border-b border-neutral-100 dark:border-neutral-800 bg-neutral-50 dark:bg-neutral-800/50">
                    <p className="text-sm font-semibold text-neutral-900 dark:text-neutral-100 truncate">{user.name}</p>
                    <p className="text-xs font-medium text-neutral-500 dark:text-neutral-400 truncate mt-0.5">{user.email}</p>
                  </div>
                  <div className="p-1">
                    <button
                      onClick={handleLogout}
                      className="w-full text-left px-3 py-2 text-sm text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-md transition-colors flex items-center gap-2 font-medium"
                    >
                      <LogOut className="w-4 h-4" />
                      Sign Out
                    </button>
                  </div>
                </div>
              )}
            </div>

          </div>
        </header>

      {/* Main Content */}
      <main className="flex-1 px-4 md:px-6 pb-4 md:pb-6 flex flex-col md:flex-row gap-4 md:gap-6 overflow-y-auto md:overflow-hidden transition-colors duration-200 z-10">
        
        {/* Left Side: Editor */}
        <section className="w-full md:w-1/2 min-h-[400px] md:min-h-0 shrink-0 md:shrink flex flex-col bg-white/80 dark:bg-[#111111]/80 backdrop-blur-xl border border-neutral-200/50 dark:border-white/10 rounded-2xl shadow-sm overflow-hidden transition-colors duration-200">
          <div className="bg-white/50 dark:bg-white/5 px-4 md:px-5 py-2 md:py-2.5 text-[10px] md:text-[11px] font-semibold text-neutral-500 dark:text-neutral-400 uppercase tracking-widest border-b border-neutral-100 dark:border-white/5 flex flex-col xl:flex-row justify-between items-start xl:items-center gap-2 xl:gap-0">
            <span className="flex items-center gap-2 pt-1 xl:pt-0 shrink-0">
              <span className="w-2 h-2 rounded-full bg-green-500"></span>
              script.js
            </span>
            
            <div className="flex flex-wrap items-center gap-1.5 md:gap-2 normal-case tracking-normal w-full xl:w-auto justify-end">
              {!error && status !== "Ready" && (
                <span className="text-[10px] md:text-xs font-medium text-neutral-500 dark:text-neutral-400 flex items-center gap-1.5 md:gap-2 px-2 md:px-3 py-1 bg-neutral-100 dark:bg-neutral-900 rounded-full border border-neutral-200 dark:border-neutral-800 shrink-0">
                  <Loader2 className="w-3 h-3 animate-spin" />
                  <span className="hidden sm:inline">{status}</span>
                </span>
              )}

              <button type="button"
                onClick={handleRun}
                disabled={isGenerating}
                className="group flex items-center justify-center gap-1.5 bg-neutral-900 dark:bg-white hover:bg-neutral-800 dark:hover:bg-neutral-100 disabled:opacity-50 text-white dark:text-neutral-900 py-1 px-2.5 md:px-3 rounded-md font-medium transition-all shadow-sm text-[10px] md:text-xs shrink-0"
              >
                {isGenerating ? (
                  <Loader2 className="w-3 h-3 md:w-3.5 md:h-3.5 animate-spin" />
                ) : (
                  <Play className="w-3 h-3 md:w-3.5 md:h-3.5 transition-transform group-hover:scale-110" fill="currentColor" />
                )}
                <span className="hidden sm:inline">Run & Preview</span>
                <span className="sm:hidden">Run</span>
              </button>
            </div>
          </div>
          <div className="flex-1 pt-4 relative">
            <Editor
              height="100%"
              defaultLanguage="javascript"
              theme={isDarkMode ? "vs-dark" : "light"}
              value={code}
              onChange={(val) => setCode(val || "")}
              options={{
                minimap: { enabled: false },
                fontSize: 14,
                fontFamily: "'JetBrains Mono', 'Fira Code', 'Menlo', monospace",
                wordWrap: "on",
                scrollBeyondLastLine: false,
                lineNumbersMinChars: 3,
                padding: { top: 12 },
                renderLineHighlight: "none",
                overviewRulerBorder: false,
                hideCursorInOverviewRuler: true
              }}
            />
          </div>
          <div className="bg-white/50 dark:bg-white/5 px-4 md:px-5 py-2 border-t border-neutral-100 dark:border-white/5 flex justify-end items-center">
            <a href="https://gitbrent.github.io/PptxGenJS/" target="_blank" rel="noreferrer" className="text-neutral-500 hover:text-neutral-900 dark:text-neutral-400 dark:hover:text-white normal-case tracking-normal transition-colors flex items-center gap-1 text-[10px] md:text-[11px] font-medium">
              PptxGenJS API Reference <span className="text-[8px] opacity-50">↗</span>
            </a>
          </div>
        </section>

        {/* Right Side: Preview */}
        <section className="w-full md:w-1/2 min-h-[500px] md:min-h-0 shrink-0 md:shrink flex flex-col bg-white/80 dark:bg-[#111111]/80 backdrop-blur-xl border border-neutral-200/50 dark:border-white/10 rounded-2xl shadow-sm overflow-hidden transition-colors duration-200">
          <div className="bg-white/50 dark:bg-white/5 px-4 md:px-5 py-2 md:py-2.5 text-[10px] md:text-[11px] font-semibold text-neutral-500 dark:text-neutral-400 uppercase tracking-widest border-b border-neutral-100 dark:border-white/5 flex flex-col xl:flex-row justify-between items-start xl:items-center gap-2 xl:gap-0">
            <span className="flex items-center gap-2 pt-1 xl:pt-0 shrink-0">
              <span className="w-2 h-2 rounded-full bg-blue-500"></span>
              PDF Preview
            </span>
            <div className="flex flex-wrap items-center gap-1.5 md:gap-2 normal-case tracking-normal w-full xl:w-auto justify-end">
              <button type="button"
                onClick={handleDownloadPptx}
                disabled={!pptxBlobUrl || isGenerating}
                className="flex items-center justify-center gap-1.5 bg-white dark:bg-neutral-800 hover:bg-neutral-50 dark:hover:bg-neutral-700 border border-neutral-200 dark:border-neutral-700 disabled:opacity-40 text-neutral-700 dark:text-neutral-300 py-1 px-2.5 md:px-2.5 rounded-md font-medium transition-all shadow-sm text-[10px] md:text-xs"
                title="Download PPTX"
              >
                <Download className="w-3 h-3 md:w-3.5 md:h-3.5" />
                <span className="hidden sm:inline">.pptx</span>
                <span className="sm:hidden">PPTX</span>
              </button>
              <button type="button"
                onClick={handleDownloadPdf}
                disabled={!previewUrl || isGenerating}
                className="flex items-center justify-center gap-1.5 bg-white dark:bg-neutral-800 hover:bg-neutral-50 dark:hover:bg-neutral-700 border border-neutral-200 dark:border-neutral-700 disabled:opacity-40 text-neutral-700 dark:text-neutral-300 py-1 px-2.5 md:px-2.5 rounded-md font-medium transition-all shadow-sm text-[10px] md:text-xs"
                title="Download PDF"
              >
                <Download className="w-3 h-3 md:w-3.5 md:h-3.5" />
                <span className="hidden sm:inline">.pdf</span>
                <span className="sm:hidden">PDF</span>
              </button>
              
              <div className="h-3 w-px bg-neutral-300 dark:bg-neutral-700 mx-0.5 md:mx-1"></div>
              
              {!driveUrl ? (
                <button type="button"
                  onClick={() => {
                    if (user) {
                      uploadToDrive(user.access_token);
                    } else {
                      pendingUploadRef.current = true;
                      login();
                    }
                  }}
                  disabled={!pptxBlobUrl || isGenerating || isUploading}
                  className="flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-600/40 text-white py-1 px-2 md:px-3 rounded-md font-medium transition-all shadow-sm text-[10px] md:text-xs ring-1 ring-emerald-700/50 dark:ring-emerald-500/50"
                  title="Open in Google Slides"
                >
                  {isUploading ? <Loader2 className="w-3 h-3 md:w-3.5 md:h-3.5 animate-spin" /> : <Cloud className="w-3 h-3 md:w-3.5 md:h-3.5" />}
                  <span className="hidden sm:inline">{isUploading ? 'Opening...' : 'Open in Google Slides'}</span>
                  <span className="sm:hidden">{isUploading ? 'Wait...' : 'Slides'}</span>
                </button>
              ) : (
                <a 
                  href={driveUrl} 
                  target="_blank" 
                  rel="noreferrer"
                  className="flex items-center justify-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-white py-1 px-2 md:px-3 rounded-md font-medium transition-all shadow-sm text-[10px] md:text-xs ring-1 ring-amber-600/50 dark:ring-amber-400/50"
                  title="Open in Google Slides"
                >
                  <Cloud className="w-3 h-3 md:w-3.5 md:h-3.5" />
                  <span className="hidden sm:inline">Open in Google Slides</span>
                  <span className="sm:hidden">Slides</span>
                </a>
              )}
            </div>
          </div>
          
          <div className="flex-1 p-5 flex flex-col relative" ref={containerRef}>
            {error ? (
              <div className="flex-1 overflow-auto rounded-xl flex flex-col items-start justify-center text-neutral-800 dark:text-neutral-200 p-8 bg-red-50/50 dark:bg-red-950/20 border border-red-100 dark:border-red-900/30 transition-colors duration-200">
                <h2 className="text-lg font-bold text-red-600 dark:text-red-400 mb-2">Execution Failed</h2>
                <p className="text-sm font-mono bg-white/50 dark:bg-black/50 p-4 rounded-lg border border-red-200/50 dark:border-red-900/50 mb-4 w-full break-words">
                  {error}
                </p>
                <button
                  onClick={handleExplainError}
                  disabled={isExplaining}
                  className="flex items-center gap-2 px-4 py-2 bg-neutral-900 dark:bg-white text-white dark:text-neutral-900 rounded-lg text-sm font-semibold shadow-sm hover:opacity-90 disabled:opacity-50 transition-opacity"
                >
                  {isExplaining ? <Loader2 className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4 text-amber-500" />}
                  Explain with AI
                </button>
                {explanation && (
                  <div className="mt-6 w-full animate-in fade-in duration-300">
                    <h3 className="text-sm font-bold flex items-center gap-2 mb-2 text-neutral-900 dark:text-neutral-100">
                      <Sparkles className="w-4 h-4 text-amber-500" />
                      AI Analysis
                    </h3>
                    <div className="text-sm leading-relaxed p-5 bg-white dark:bg-neutral-900 rounded-xl border border-amber-200/50 dark:border-amber-900/30 shadow-sm prose prose-sm dark:prose-invert max-w-none">
                      <Markdown>{explanation}</Markdown>
                    </div>
                  </div>
                )}
              </div>
            ) : pdfError ? (
              <div className="flex-1 overflow-auto rounded-xl flex flex-col items-center justify-center text-neutral-800 dark:text-neutral-200 p-8 bg-amber-50/50 dark:bg-amber-950/20 border border-amber-100 dark:border-amber-900/30 transition-colors duration-200">
                <h2 className="text-lg font-bold text-amber-600 dark:text-amber-500 mb-2">PDF Conversion Failed</h2>
                <p className="text-sm font-mono text-center bg-white/50 dark:bg-black/50 p-4 rounded-lg border border-amber-200/50 dark:border-amber-900/50 mb-6 w-full max-w-md break-words">
                  {pdfError}
                </p>
                <button
                  onClick={handleRetryPdf}
                  disabled={isGenerating}
                  className="flex items-center gap-2 px-5 py-2.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-semibold shadow-sm disabled:opacity-50 transition-colors"
                >
                  {isGenerating ? <Loader2 className="w-4 h-4 animate-spin" /> : <Play className="w-4 h-4" />}
                  Retry Preview
                </button>
                <p className="text-xs text-neutral-500 mt-4 text-center max-w-sm">
                  Your PPTX was successfully generated and can be downloaded or opened in Google Slides. Only the browser preview failed.
                </p>
              </div>
            ) : previewUrl && !isGenerating ? (
              <div className="flex-1 bg-neutral-100/50 dark:bg-black/40 rounded-xl shadow-inner border border-neutral-200/50 dark:border-white/5 overflow-hidden relative flex flex-col transition-colors duration-200">
                <div className="flex-1 overflow-auto w-full flex justify-center py-6 custom-scrollbar">
                  <Document
                    file={previewUrl}
                    options={pdfOptions}
                    onLoadSuccess={({ numPages }) => {
                      setNumPages(numPages);
                      setPageNumber(1);
                    }}
                    onLoadError={(error) => {
                      console.error("PDF Viewer Error:", error);
                      setError("Failed to render PDF preview: " + error.message);
                    }}
                    loading={<div className="flex items-center justify-center p-8"><Loader2 className="w-6 h-6 animate-spin text-neutral-400" /></div>}
                    className="max-w-full drop-shadow-xl"
                  >
                    <Page 
                      pageNumber={pageNumber} 
                      width={containerWidth ? containerWidth - 60 : undefined} 
                      renderTextLayer={false}
                      renderAnnotationLayer={false}
                      className="bg-white shrink-0 rounded-sm overflow-hidden"
                    />
                  </Document>
                </div>
                {numPages && numPages > 1 && (
                  <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-white/90 dark:bg-neutral-900/90 backdrop-blur border border-neutral-200 dark:border-neutral-800 shadow-lg rounded-full py-1.5 px-2 flex items-center justify-center gap-1 transition-colors duration-200">
                    <button type="button" 
                      onClick={() => setPageNumber(1)}
                      disabled={pageNumber <= 1}
                      className="p-1 rounded-full hover:bg-neutral-100 dark:hover:bg-neutral-800 disabled:opacity-30 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronFirst className="w-4 h-4" />
                    </button>
                    <button type="button" 
                      onClick={() => setPageNumber(Math.max(1, pageNumber - 1))}
                      disabled={pageNumber <= 1}
                      className="p-1 rounded-full hover:bg-neutral-100 dark:hover:bg-neutral-800 disabled:opacity-30 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <span className="text-xs font-semibold tracking-wider text-neutral-600 dark:text-neutral-400 tabular-nums px-2">{pageNumber} / {numPages}</span>
                    <button type="button" 
                      onClick={() => setPageNumber(Math.max(1, Math.min(numPages, pageNumber + 1)))}
                      disabled={pageNumber >= numPages}
                      className="p-1 rounded-full hover:bg-neutral-100 dark:hover:bg-neutral-800 disabled:opacity-30 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                    <button type="button" 
                      onClick={() => setPageNumber(numPages)}
                      disabled={pageNumber >= numPages}
                      className="p-1 rounded-full hover:bg-neutral-100 dark:hover:bg-neutral-800 disabled:opacity-30 text-neutral-700 dark:text-neutral-300 transition-colors"
                    >
                      <ChevronLast className="w-4 h-4" />
                    </button>
                  </div>
                )}
              </div>
            ) : (
              <div className="flex-1 border border-dashed border-neutral-300 dark:border-neutral-800 rounded-xl flex flex-col items-center justify-center text-neutral-400 dark:text-neutral-500 p-8 text-center bg-transparent transition-colors duration-200">
                <div className="w-12 h-12 bg-neutral-100 dark:bg-neutral-900 rounded-full flex items-center justify-center mb-4 shadow-sm">
                  <Play className="w-5 h-5 text-neutral-400 dark:text-neutral-600 ml-1" />
                </div>
                <p className="text-sm font-semibold text-neutral-600 dark:text-neutral-400 mb-1">Canvas is empty</p>
                <p className="max-w-[220px] text-xs text-neutral-500 dark:text-neutral-500 leading-relaxed">
                  Write JavaScript on the left and hit <strong>Run & Preview</strong>.
                </p>
              </div>
            )}
          </div>
        </section>
      </main>
      </div>
    </div>
  );
}
