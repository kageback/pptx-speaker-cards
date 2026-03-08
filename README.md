# pptx-speaker-cards

Generate printable speaker cards from PowerPoint presentation notes. Creates A4 landscape PDF with 4 cards per page in a 2x2 grid, perfect for printing and using during presentations.

## Features

- **Rich Text Formatting**: Preserves bold, italic, and bullet formatting from PowerPoint notes
- **Smart Title Extraction**: Uses first line of notes as card title
- **Auto Font Sizing**: Automatically reduces font size to fit content on cards
- **Continuation Cards**: Automatically splits long content across multiple cards
- **OneDrive/SharePoint Support**: Download and process presentations directly from sharing links
- **Customizable Layout**: Adjust margins, font sizes, and other formatting options
- **Hidden Slides**: Optionally include or exclude hidden slides
- **Progress Feedback**: Shows processing progress in console

## Installation

### Using pip

```bash
pip install pptx-speaker-cards
```

### Using pipx (recommended for CLI tools)

```bash
pipx install pptx-speaker-cards
```

### Using uv

```bash
uv tool install pptx-speaker-cards
```

## Quick Start

```bash
# Process a local PowerPoint file
pptx-speaker-cards presentation.pptx

# Process from OneDrive/SharePoint public link
pptx-speaker-cards "https://your-sharepoint.com/.../file.pptx?e=..."

# Custom output filename
pptx-speaker-cards slides.pptx --output my-cards.pdf

# Adjust margins
pptx-speaker-cards slides.pptx --margin-top 10 --margin-left 8
```

## Usage

```
pptx-speaker-cards [-h] [--version] [--slide-number {yes,no}]
                   [--title-font-size SIZE] [--body-font-size SIZE]
                   [--include-hidden] [--output PATH]
                   [--margin-top MM] [--margin-bottom MM]
                   [--margin-left MM] [--margin-right MM]
                   input
```

### Positional Arguments

- `input` - PowerPoint file path or OneDrive/SharePoint URL

### Optional Arguments

- `-h, --help` - Show help message and exit
- `--version` - Show program version number and exit
- `--slide-number {yes,no}` - Show slide numbers on cards (default: yes)
- `--title-font-size SIZE` - Title font size in points (default: 12.0)
- `--body-font-size SIZE` - Override body font size (default: auto-fit from 10pt down to 6pt)
- `--include-hidden` - Include hidden slides in output
- `--output PATH` - Output PDF path (default: [input]_speaker_notes.pdf)
- `--margin-top MM` - Top margin in millimeters (default: 5)
- `--margin-bottom MM` - Bottom margin in millimeters (default: 5)
- `--margin-left MM` - Left margin in millimeters (default: 5)
- `--margin-right MM` - Right margin in millimeters (default: 5)

## Examples

### Basic Usage

```bash
# Process local file with default settings
pptx-speaker-cards my-presentation.pptx
```

### OneDrive/SharePoint Links

```bash
# Process presentation from OneDrive sharing link
pptx-speaker-cards "https://onedrive.live.com/...?e=xyz123"

# Process from SharePoint
pptx-speaker-cards "https://company.sharepoint.com/:p:/g/...?e=abc456"
```

### Customization

```bash
# Custom output filename
pptx-speaker-cards slides.pptx --output speaker-cards-2026.pdf

# Larger title font
pptx-speaker-cards slides.pptx --title-font-size 14

# Fixed body font size (disable auto-sizing)
pptx-speaker-cards slides.pptx --body-font-size 9

# Wider margins for easier cutting
pptx-speaker-cards slides.pptx --margin-top 10 --margin-bottom 10

# Include hidden slides
pptx-speaker-cards slides.pptx --include-hidden

# No slide numbers
pptx-speaker-cards slides.pptx --slide-number no
```

### Module Execution

```bash
# Can also be run as a module
python -m pptx_speaker_cards presentation.pptx
```

## How It Works

### Card Layout

Each card includes:
- **Title** (top-left): First line from speaker notes
- **Slide Number** (top-right): Original slide number (e.g., "5")
- **Content** (body): Remaining speaker notes with formatting preserved

### Continuation Cards

When notes are too long to fit on a single card, the tool automatically:
1. Reduces font size from 10pt down to 6pt minimum
2. If still too long, splits content across multiple cards
3. Labels continuation cards (e.g., "5+1", "5+2")
4. Marks continuation titles with "..continued[Original Title]"

### Text Formatting

Preserves from PowerPoint:
- **Bold text**
- *Italic text*
- Bullet points (•)
- Indentation levels
- Tab characters (converted to 4 spaces)
- Line breaks and paragraph spacing

### Page Layout

- **Page Size**: A4 landscape (297mm × 210mm)
- **Cards per Page**: 4 (2×2 grid)
- **Cut Lines**: Dotted lines between cards for easy cutting
- **Default Margins**: 5mm on all sides of each card

## PowerPoint Notes Format

For best results, format your PowerPoint notes as:

```
Title of the Slide
Main talking point here.
• First bullet point
• Second bullet point

Additional paragraphs...
```

The first line becomes the card title, and everything else becomes the card content.

## Development

### Setting Up Development Environment

```bash
# Clone the repository
git clone https://github.com/kageback/pptx-speaker-cards.git
cd pptx-speaker-cards

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run directly
python -m pptx_speaker_cards your-file.pptx
```

### Building the Package

```bash
# Install build tools
pip install build twine

# Build distribution files
python -m build

# Check the built package
twine check dist/*
```

### Running Tests

```bash
# Install test dependencies
pip install pytest

# Run tests
pytest tests/
```

## Requirements

- Python 3.8 or higher
- python-pptx >= 0.6.23
- reportlab >= 4.1.0
- requests >= 2.31.0

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- **Issues**: Report bugs or request features on [GitHub Issues](https://github.com/kageback/pptx-speaker-cards/issues)
- **Documentation**: See this README and `pptx-speaker-cards --help`

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and changes.

## Author

Magnus Kågebäck

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint processing
- Built with [ReportLab](https://www.reportlab.com/) for PDF generation
