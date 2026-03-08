# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-03-06

### Added
- Initial release of pptx-speaker-cards
- Generate speaker cards from PowerPoint notes in A4 landscape PDF format
- 4 cards per page in 2×2 grid with dotted cut lines
- Rich text formatting support (bold, italic, bullets)
- Smart title extraction from first line of notes
- Automatic font size reduction (10pt down to 6pt minimum)
- Automatic continuation cards for long content
- OneDrive/SharePoint public link support
- Customizable margins and font sizes
- Optional slide number display
- Hidden slide support (include/exclude)
- Progress feedback during processing
- Command-line interface with comprehensive options
- Module execution support (`python -m pptx_speaker_cards`)
- Tab character handling (converted to 4 spaces)
- Tighter line spacing for more content per card

### Features
- **Local files**: Process PowerPoint files from local filesystem
- **URL support**: Download and process from OneDrive/SharePoint links
- **Auto-fitting**: Automatically adjusts font size to fit content
- **Smart splitting**: Splits long notes across multiple cards at paragraph boundaries
- **Clean output**: Preserves formatting while creating printable cards

[1.0.0]: https://github.com/kageback/pptx-speaker-cards/releases/tag/v1.0.0
