# LaTeX/Markdown to Word Converter

Convert LaTeX `.tex` and Markdown `.md` files to Microsoft Word documents (`.docx`) with **native Word equations**! This tool automatically transforms embedded LaTeX equations into OMML (Office Math Markup Language) format for proper rendering in Microsoft Word.

Also includes a BibTeX (`.bib`) to Word converter for generating formatted bibliography documents.

## Features

- 📄 **Multiple input formats**: Convert both LaTeX (`.tex`) and Markdown (`.md`) files
- 📚 **Bibliography conversion**: Convert BibTeX (`.bib`) files to formatted Word documents
- ✨ **Automatic equation conversion**: LaTeX equations → native Word equations (OMML)
- 📄 **Full document support**: Handles chapters, sections, subsections, figures, captions, and citations
- 🧮 **Multiple equation environments**: Support for `equation`, `align`, `gather`, `multline`, and inline equations
- 📝 **Text formatting**: Preserves subscripts, superscripts, bold, italic, and other text formatting
- 🏷️ **Figure captions**: Extracts figure captions and labels (wrapped in brackets) while converting inline equations within them
- 🔢 **Equation labels**: Preserves equation labels below display equations
- 🔗 **Reference preservation**: Keeps all reference labels (figures, equations, sections) in square brackets
- 📋 **List support**: Converts `itemize` (bullets) and `enumerate` (numbered) environments to proper Word lists
- � **Advanced LaTeX handling**: Supports `\texorpdfstring`, `\resizebox`, and escaped characters (`\%`)
- �💬 **Citation preservation**: Keeps citation labels in square brackets (e.g., `[author2023]`)

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

### 1. Clone the repository
```bash
git clone https://github.com/yeewantung/tex_math_to_word.git
cd tex_math_to_word
```

### 2. Install texmath (Haskell)
Follow the installation instructions at https://github.com/jgm/texmath. On macOS with Homebrew:
```bash
brew install texmath
```

Or install from Hackage:
```bash
cabal install texmath --install-method=copy --installdir ~/.local/bin
```

### 3. Install Python dependencies
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

### BibTeX files
```bash
python latex_bib_to_word.py path/to/your/references.bib
```

Or simply run without arguments to enter the file path interactively:
```bash
python latex_bib_to_word.py
# Then enter: path/to/your/references.bib
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
- For BibTeX files: `python latex_bib_to_word.py sample.bib` creates `sample.docx` with formatted references

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
- `\chapter{...}` - converts to title heading (level 0)
- `\section{...}`, `\subsection{...}`, `\subsubsection{...}`
- `\begin{figure}...\end{figure}` - extracts captions (wrapped as `[Caption: ...]`) and labels
- `\begin{itemize}...\item...\end{itemize}` - creates bulleted lists
- `\begin{enumerate}...\item...\end{enumerate}` - creates numbered lists
- `\cite{...}` - preserves citation labels as `[label]`
- `\ref{...}` - preserves references as `[label]`
- `\reffig{...}`, `\refeqn{...}` - converts to "Fig." and "Eq."

### Equation Features
- Equation labels: `\label{eqn_name}` appears as `[eqn_name]` below equations
- `\resizebox{...}{...}{$...$}` - automatically strips wrapper and processes equation
- Inline equations within figure captions are properly converted

### Special Handling
- `\texorpdfstring{pdf_version}{simple_version}` - extracts simple version
- Removes LaTeX comments (`% ...`)
- Skips preamble (starts from `\begin{document}` or first section)
- Converts `~` (non-breaking spaces) to regular spaces
- Converts `\%` to `%`
- Handles delimiter commands like `\Bigl(`, `\Bigr)`, etc.

## Example

**Input LaTeX:**
```latex
\chapter{Introduction}

\section{Theory}

The energy is given by Einstein's equation:
\begin{equation}
E = mc^2
\label{eqn_energy}
\end{equation}

Key properties:
\begin{itemize}
    \item Mass-energy equivalence as shown in \ref{eqn_energy}
    \item Speed of light $c = 3 \times 10^8$ m/s
\end{itemize}

We define H\textsubscript{2}O concentration as $c$.
```

**Output Word document:**
- Chapter heading: "Introduction" (title style)
- Section heading: "Theory"
- Display equation for $E = mc^2$ with label `[eqn_energy]` below it
- Bulleted list with two items, including inline equation and reference
- Subscript properly formatted in "H₂O"
- All equations rendered as native Word equations

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
- bibtexparser (for BibTeX conversion)
- texmath (Haskell executable)

## BibTeX to Word Converter

The `latex_bib_to_word.py` script converts BibTeX bibliography files to formatted Word documents. Each reference appears with its citation key as a label followed by the formatted reference content.

### Supported BibTeX Entry Types
- `@article` - Journal articles
- `@inproceedings` / `@conference` - Conference papers
- `@book` - Books

### Reference Format
Each reference is formatted as:
```
[citation_key] Author(s), "Title", Journal/Conference, vol. X, no. Y, pp. Z, Month Year, doi: XX.XXXX/XXXXX
```

### Example Input (BibTeX):
```bibtex
@article{Einstein1905,
  author = {A. Einstein},
  title = {On the electrodynamics of moving bodies},
  journal = {Annalen der Physik},
  volume = {17},
  pages = {891-921},
  year = {1905}
}
```

### Example Output (Word):
**[Einstein1905]** A. Einstein, "On the electrodynamics of moving bodies", Annalen der Physik, vol. 17, pp. 891-921, 1905.
