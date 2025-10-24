"""
Main entry point for GibbsCAM Processor.
Works both as a module and as a standalone executable.
"""

# Import processor main function
# Use absolute import that works in PyInstaller
import processor

if __name__ == "__main__":
    processor.main()