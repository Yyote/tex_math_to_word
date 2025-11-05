# LaTeX/Markdown to Word Converter

Convert LaTeX `.tex` and Markdown `.md` files to Microsoft Word documents (`.docx`) with **native Word equations**! This tool automatically transforms embedded LaTeX equations into OMML (Office Math Markup Language) format for proper rendering in Microsoft Word.

## Features

- 📄 **Multiple input formats**: Convert both LaTeX (`.tex`) and Markdown (`.md`) files
- ✨ **Automatic equation conversion**: LaTeX equations → native Word equations (OMML)
- 📄 **Full document support**: Handles sections, subsections, figures, captions, and citations
- 🧮 **Multiple equation environments**: Support for `equation`, `align`, `gather`, `multline`, and inline equations
- 📝 **Text formatting**: Preserves subscripts, superscripts, bold, italic, and other text formatting
- 🏷️ **Figure captions**: Extracts figure captions and labels while converting inline equations within them
- 💬 **Citation preservation**: Keeps citation labels in square brackets (e.g., `[...]`)

## How It Works

This tool builds on top of the excellent [**texmath**](https://github.com/jgm/texmath) Haskell package—a versatile and mature converter supporting multiple math markup formats including LaTeX, MathML, and OMML. 

Our implementation adds a Python wrapper that:
1. Extracts LaTeX equations from `.tex` files
2. Uses texmath to convert them to OMML
3. Embeds the OMML equations into a Word document
4. Processes LaTeX structure (sections, figures, citations) for clean Word formatting

## Installation

### Prerequisites
- Python 3.7+
- [texmath](https://github.com/jgm/texmath) installed and available at `~/.local/bin/texmath`

### Install texmath (Haskell)
Follow the installation instructions at https://github.com/jgm/texmath. On macOS with Homebrew:
```bash
brew install texmath
```

Or install from Hackage:
```bash
cabal install texmath --install-method=copy --installdir ~/.local/bin
```

### Install Python dependencies
```bash
pip install -r requirements.txt
```

## Usage

### LaTeX files
```bash
python latex_to_word.py path/to/your/file.tex
```

### Markdown files
```bash
python markdown_to_word.py path/to/your/file.md
```

### Verbose output
```bash
python latex_to_word.py -v path/to/your/file.tex
python markdown_to_word.py -v path/to/your/file.md
```

**Notes:**
- Both **absolute and relative paths** are supported for input files
- The converted Word document will be saved in the **current working directory** with the same filename (e.g., `file.docx`)
- Example: `python latex_to_word.py ~/Documents/paper.tex` creates `paper.docx` in the current folder

## Supported LaTeX Features

### Equation Environments
- `\begin{equation}...\end{equation}` and `equation*`
- `\begin{align}...\end{align}` and `align*`
- `\begin{gather}...\end{gather}` and `gather*`
- `\begin{multline}...\end{multline}` and `multline*`
- Display equations: `$$...$$`
- Inline equations: `$...$`

### Text Formatting
- `\textbf{...}`, `\textit{...}`, `\texttt{...}`, `\emph{...}`
- `\textsubscript{...}`, `\textsuperscript{...}`
- `\mathrm{...}`, `\mathbf{...}`, `\mathcal{...}`, `\mathbb{...}`

### Document Structure
- `\section{...}`, `\subsection{...}`, `\subsubsection{...}`
- `\begin{figure}...\end{figure}` - extracts captions and labels
- `\cite{...}` - preserves citation labels
- `\reffig{...}`, `\refeqn{...}` - converts to "Fig." and "Eq."

### Special Handling
- Removes LaTeX comments (`% ...`)
- Skips preamble (starts from `\begin{document}` or first section)
- Converts `~` (non-breaking spaces) to regular spaces
- Handles delimiter commands like `\Bigl(`, `\Bigr)`, etc.

## Example

**Input LaTeX:**
```latex
\section{Theory}

The energy is given by Einstein's equation:
$$E = mc^2$$

We define H\textsubscript{2}O concentration as $c$.
```

**Output Word document:**
- Section heading: "Theory"
- Paragraph with the text and native Word equation for $E = mc^2$
- Subscript properly formatted in "H₂O"
- Inline equation rendered as Word equation

## Acknowledgments

This project leverages [**texmath**](https://github.com/jgm/texmath) by John MacFarlane—a comprehensive and well-maintained Haskell package for converting between various mathematical markup formats. We highly recommend checking out the original project for more information on its capabilities.

Development assistance provided by AI.

## License

MIT License - See LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Roadmap

- [ ] Web interface for online conversion (GitHub Pages + backend API)
- [ ] Word to LaTeX conversion
- [ ] Word to Markdown conversion
- [ ] Support for more LaTeX packages and environments
- [ ] Batch file conversion

## Requirements

- Python 3.7+
- python-docx
- texmath (Haskell executable)
