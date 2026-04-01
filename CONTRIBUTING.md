# Contributing to PowerSTIG Converter UI

Thank you for your interest in contributing to PowerSTIG Converter UI! 🎉

## Project Maintainer

This project is maintained by [@MrasmussenGit](https://github.com/MrasmussenGit).

## How to Contribute

### Reporting Issues

If you find a bug or have a feature request:

1. **Check existing issues** to avoid duplicates
2. **Create a new issue** with a clear title and description
3. **Include details**:
   - Steps to reproduce (for bugs)
   - Expected vs actual behavior
   - Screenshots if applicable
   - Your environment (.NET version, Windows version, PowerSTIG version)

### Submitting Changes

1. **Fork the repository**
2. **Create a feature branch** from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
   Or for bug fixes:
   ```bash
   git checkout -b fix/bug-description
   ```

3. **Make your changes**:
   - Follow the existing code style
   - Add comments for complex logic
   - Update documentation if needed

4. **Test your changes**:
   - Build the solution successfully
   - Test all affected features
   - Verify no existing features are broken

5. **Commit your changes**:
   ```bash
   git commit -m "Add: Brief description of your changes"
   ```
   Use prefixes like:
   - `Add:` for new features
   - `Fix:` for bug fixes
   - `Update:` for improvements
   - `Refactor:` for code restructuring
   - `Docs:` for documentation changes

6. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```

7. **Create a Pull Request**:
   - Provide a clear description of the changes
   - Reference any related issues
   - Explain why the change is needed

## Development Guidelines

### Code Style

- **C#**: Follow standard C# conventions
  - Use PascalCase for public members
  - Use camelCase for private fields with `_` prefix
  - Use meaningful variable names
  - Add XML documentation for public methods

- **XAML**: Follow WPF best practices
  - Keep UI and logic separated
  - Use existing styles from `Styles/` folder
  - Maintain consistent spacing and indentation

### Project Structure

```
PowerStigConverterUI/
├── *.xaml           # Window definitions
├── *.xaml.cs        # Code-behind files
├── Styles/          # XAML style resources
├── *.cs             # Helper classes
└── README.md
```

### Testing

Before submitting a PR:

1. Build the solution in Release mode
2. Test the following scenarios:
   - Convert a STIG with ZIP file
   - Convert a STIG with direct XCCDF
   - Compare two STIGs
   - Split a Windows OS STIG
3. Verify HTML report generation
4. Check error handling with invalid inputs

## Questions or Need Help?

- **Open an issue** for general questions
- **Start a discussion** for ideas or feedback
- **Contact**: [@MrasmussenGit](https://github.com/MrasmussenGit)

## Code of Conduct

- Be respectful and professional
- Provide constructive feedback
- Help others learn and grow
- Focus on the code, not the person

## Recognition

Contributors will be acknowledged in the project documentation.

---

Thank you for contributing! 🚀
