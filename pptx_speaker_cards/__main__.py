"""
Allow package to be executed as a module:
    python -m pptx_speaker_cards
"""

from .cli import main

if __name__ == '__main__':
    main()
