# Contributing to PST Splitter - Ensue Edition

Thank you for your interest in contributing to PST Splitter! This document provides guidelines for contributing to the project.

## 🚀 **Development Setup**

### Prerequisites
- Python 3.12+ (recommended)
- Windows 10/11 with Microsoft Outlook installed
- PyInstaller for building executables

### Environment Setup
```powershell
# Clone the repository
git clone <repository-url>
cd PST_Splitter

# Create virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

## 🧪 **Testing**

### Running Tests
```powershell
# Run all tests
pytest -q

# Run specific test file
pytest tests/test_grouping.py -v

# Run with coverage
pytest --cov=pstsplitter tests/
```

### Test Coverage
The project maintains comprehensive test coverage including:
- ✅ PST file grouping logic
- ✅ Date range filtering
- ✅ Folder filtering
- ✅ Error handling scenarios
- ✅ Performance optimization features

## 🔨 **Building**

### Development Build
```powershell
# Run from source
python -m pstsplitter
```

### Production Build
```powershell
# Build single-file executable
pyinstaller pstsplitter_onefile.spec --clean

# Build with dependencies folder
pyinstaller pstsplitter.spec --clean
```

## 📝 **Code Style**

### Python Standards
- Follow PEP 8 guidelines
- Use type hints where appropriate
- Include comprehensive docstrings
- Maintain consistent naming conventions

### Code Quality Tools
```powershell
# Type checking
python -m mypy src/pstsplitter

# Linting
ruff check src/pstsplitter

# Formatting
black src/pstsplitter
```

## 🐛 **Reporting Issues**

### Before Reporting
1. Check existing issues to avoid duplicates
2. Test with the latest version
3. Use the Export Analysis feature to generate detailed logs

### Issue Template
```markdown
**Description**
Brief description of the issue

**Steps to Reproduce**
1. Step one
2. Step two
3. Step three

**Expected Behavior**
What should happen

**Actual Behavior**
What actually happens

**Environment**
- PST Splitter Version: 
- Windows Version: 
- Outlook Version: 
- PST File Size: 

**Logs**
Please attach exported analysis files (JSON/CSV/TXT)
```

## 🚀 **Feature Requests**

### Guidelines
- Check existing feature requests first
- Provide clear use cases and benefits
- Consider backward compatibility
- Include mockups or examples if applicable

## 📋 **Pull Request Process**

### Before Submitting
1. ✅ Run all tests (`pytest -q`)
2. ✅ Check code quality (`mypy`, `ruff`)
3. ✅ Update documentation if needed
4. ✅ Test the executable build
5. ✅ Update CHANGELOG.md

### PR Guidelines
- Use clear, descriptive commit messages
- Reference related issues
- Include screenshots for UI changes
- Keep PRs focused and atomic
- Write comprehensive PR descriptions

### Commit Message Format
```
type(scope): brief description

Detailed explanation of changes if needed

Fixes #issue-number
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

## 🏗️ **Development Guidelines**

### Architecture
- **GUI**: Tkinter-based responsive design
- **PST Processing**: Win32 COM API integration
- **Performance**: Batch processing with progress tracking
- **Error Handling**: Comprehensive exception management

### Key Files
- `src/pstsplitter/gui.py` - Main UI implementation
- `src/pstsplitter/splitter.py` - Core PST processing logic
- `src/pstsplitter/outlook.py` - Outlook COM integration
- `src/pstsplitter/util.py` - Utility functions

### UI Development
- Design for 15-inch screen compatibility
- Implement responsive layouts
- Include proper progress indicators
- Provide comprehensive user feedback

### Performance Considerations
- Use batch processing for large datasets
- Implement proper cancellation support
- Monitor memory usage
- Provide real-time progress updates

## 🔒 **Security Guidelines**

- Never commit sensitive data
- Handle user files with appropriate permissions
- Validate all user inputs
- Follow secure coding practices
- Test with various PST file sizes and conditions

## 📚 **Documentation**

### Required Documentation
- Update README.md for new features
- Include inline code comments
- Update docstrings for public APIs
- Create user guides for complex features

### Documentation Standards
- Use clear, concise language
- Include code examples
- Provide screenshots for UI features
- Keep documentation up-to-date

## 🤝 **Community Guidelines**

- Be respectful and inclusive
- Help others learn and contribute
- Share knowledge and best practices
- Provide constructive feedback
- Follow the project's code of conduct

## 📞 **Getting Help**

- **Issues**: GitHub Issues for bugs and feature requests
- **Discussions**: GitHub Discussions for questions
- **Email**: info@ensue.com for private inquiries
- **Documentation**: Check README.md and inline docs

## 🏷️ **Release Process**

1. Update version in `version_info.py`
2. Update CHANGELOG.md
3. Run comprehensive tests
4. Build and test executables
5. Create release notes
6. Tag release in Git
7. Publish executables

Thank you for contributing to PST Splitter - Ensue Edition! 🚀
