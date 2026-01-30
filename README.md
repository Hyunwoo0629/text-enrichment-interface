# Text Enrichment Interface Studio

Web-based interface for applying typography and highlighting styles to Word documents (.docx).

## Features

- **Upload Word Documents**: Drag-and-drop or click to upload .docx files
- **Direct Text Styling**: Select text directly and apply styles
- **Typography Options**:
  - **Bold**: Make text bold
  - **Italic**: Make text italic
  - **Underline**: Add underline to text
  - **Strikethrough**: Add strikethrough effect

- **Highlighting Options**:
  - **Background Highlight**: Apply colored background
  - **Text Color**: Change the color of selected text
  - **Box Border**: Add rectangular border around text
  - **Circle Border**: Add elliptical border around text

- **Color Customization**: 
  - Individual color pickers for highlight, text, and border colors
  - Quick color preset buttons

## Directory Structure

```
text-enrichment-interface/
├── back/
│   ├── app.py              # Flask backend server
│   ├── requirements.txt    # Python dependencies
│   ├── uploads/            # Uploaded documents 
│   └── data/               # Style data 
├── front/
│   ├── index.html          # Main HTML file
│   ├── styles.css          # CSS styles 
│   └── script.js           # JavaScript application
└── README.md
```

## Setup Instructions

### Prerequisites

- Python 3.8 or higher
- pip
- Modern web browser

### Step 1: Install Dependencies

Open terminal in the `backend` folder:

```bash
cd text-enrichment-interface/back
pip install -r requirements.txt
```

### Step 2: Run the Server

```bash
python app.py
/ python3 app.py
```

You should see:
```
==================================================
Server Starting...
Server running at: http://localhost:5001
==================================================
```

### Step 3: Open the Application

Go to: **http://localhost:5001**

## Usage

### Uploading a Document

1. Click "Upload Document" button, or
2. Drag and drop a .docx file onto the viewer

### Applying Styles

1. **Select a Tool**: Click a tool button in the sidebar (Bold, Italic, Highlight, etc.)
2. **Select Text**: Click and drag to select text in the document
3. **Style Applied**: The style is automatically applied when you release the mouse

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+B` | Bold tool |
| `Ctrl+I` | Italic tool |
| `Ctrl+U` | Underline tool |
| `H` | Highlight tool |
| `T` | Text color tool |
| `R` | Rectangle border tool |
| `C` | Circle border tool |
| `Ctrl+Z` | Undo |
| `Ctrl+S` | Save styles |
| `Escape` | Deselect tool |

### Managing Styles

- View all applied styles in the right panel
- Click the X button on any style to delete it
- Use "Clear All" to remove all styles
- Use "Save" to persist styles to the server

## Technology Stack

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Backend**: Python Flask
- **Document Parsing**: python-docx

## Troubleshooting

### "Module not found: docx"
Run: `pip install python-docx`

### Document not uploading
- Ensure it's a .docx file (not .doc)
- Check file isn't corrupted
- Check terminal for error messages
