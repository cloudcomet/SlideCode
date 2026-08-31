# 🚀 SlideCode

> **A powerful, web-based Generator Studio that seamlessly converts code and structured text into polished PowerPoint presentations (.pptx) with live PDF previews and direct Google Drive integration.**

![Build Status](https://img.shields.io/badge/build-passing-brightgreen)
![TypeScript](https://img.shields.io/badge/TypeScript-Ready-blue)
![React](https://img.shields.io/badge/React-18-blueviolet)

SlideCode provides a highly responsive, split-pane environment where you can write presentation logic or content on the left, and instantly preview the rendered slides on the right. 

## ✨ Features

- **📝 Live Editor:** Write and manage your slide generation logic in a comfortable, dark-mode-ready editor.
- **💾 Auto-Save:** Your editor code is securely auto-saved to local storage, preventing data loss on accidental refreshes.
- **👁️ Instant PDF Previews:** Compile your slide code into a live, scrollable PDF preview right in the browser.
- **🔄 Robust Error Handling:** Includes an explicit PDF conversion retry mechanism to gracefully handle temporary network or server timeouts without losing your generated presentation.
- **📑 Advanced Navigation:** Jump instantly to the first or last page of your generated preview alongside standard pagination.
- **💾 Direct Downloads:** Export your generated presentations as `.pptx` or `.pdf` files with a single click.
- **☁️ Google Drive Integration:** Sign in with your Google account to directly push generated presentations to your Google Drive.
- **🌓 Responsive Design & Dark Mode:** A sleek, mobile-responsive layout that supports seamless light and dark modes for optimal viewing in any environment.

## 🛠️ Tech Stack

- **Frontend:** React 18, Vite, Tailwind CSS, Lucide Icons
- **Backend:** Node.js, Express
- **Presentation Engine:** [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
- **Authentication:** Google OAuth2
- **Language:** TypeScript across the entire stack

## 🚀 Getting Started

### Prerequisites

- Node.js (v18 or higher)
- npm or yarn

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/slidecode.git
   cd slidecode
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Environment Setup:**
   Copy the example environment file and configure your API keys (if applicable):
   ```bash
   cp .env.example .env
   ```
   *Note: Firebase configuration is intentionally git-ignored for security. If you are deploying this yourself, ensure your Firebase details are provided securely via your deployment platform.*

4. **Start the development server:**
   ```bash
   npm run dev
   ```
   The application will boot up at `http://localhost:3000`.

## 📦 Build for Production

To create a production-ready build:

```bash
npm run build
```
This compiles the Vite React frontend into the `dist/` directory and bundles the Express server using ESBuild. 

To start the production server:
```bash
npm start
```

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](#) if you want to contribute.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.
