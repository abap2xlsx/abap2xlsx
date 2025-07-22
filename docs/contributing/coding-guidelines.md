# 🔧 Coding Guidelines

Development standards and best practices for contributing to abap2xlsx. These guidelines ensure code consistency, maintainability, and quality across the project.

## 📋 Table of Contents

- [Naming Conventions](#naming-conventions)
- [Code Quality with abaplint](#code-quality-with-abaplint)
- [Development Standards](#development-standards)
- [Best Practices](#best-practices)

---

## 🏷️ Naming Conventions

### Project Context

In abap2xlsx, over time, the ABAP code came to mix different [naming standards](https://github.com/abap2xlsx/abap2xlsx/issues/773). Naming standards may vary from one class to another, but one naming standard is usually correctly applied in each class.

### Compatibility Considerations

- **Existing Code**: It's not possible to impose one naming standard by fixing existing names, because clients may have developed custom programs which may already refer to them
- **Consistency Rule**: When adding new variables, parameters, or types in existing objects or methods, maintain consistency with existing naming in that context
- **Method Updates**: If you fix a method, reuse its naming standards so that its code looks consistent. If multiple standards exist, choose the most used one

### Standard Naming Rules

For creating new ABAP objects or method implementations (mix of Hungarian notation and [Clean Code](https://github.com/SAP/styleguides/blob/main/clean-abap/CleanABAP.md#avoid-encodings-esp-hungarian-notation-and-prefixes)):

#### Method Names

- **Style**: Clean code approach (descriptive, no prefixes)

#### Local Classes

- **Classes**: `LCL_` prefix
- **Exception Classes**: `LCX_` prefix (for the only exception class)

#### Types

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `TV_` | `TV_COLUMN_WIDTH` |
| **Structure** | `TS_` | `TS_CELL_DATA` |
| **Table** | `TT_` | `TT_WORKSHEET_LIST` |

#### Attributes

- **Instance/Class Attributes**: Clean code (no prefixes)
- **Constants/Global Attributes**: `C_` prefix

#### Local Variables

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `LV_` | `LV_ROW_COUNT` |
| **Object** | `LO_` | `LO_WORKSHEET` |
| **Structure** | `LS_` | `LS_CELL_STYLE` |
| **Table** | `LT_` | `LT_DATA_ROWS` |
| **Data Reference** | `LR_` | `LR_FIELD_SYMBOL` |

#### Constants

- **Local Constants**: `LC_` prefix

#### Field Symbols

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `<LV_` | `<LV_FIELD_VALUE>` |
| **Structure** | `<LS_` | `<LS_ROW_DATA>` |
| **Table** | `<LT_` | `<LT_TABLE_DATA>` |

#### Method Parameters

**IMPORTING Parameters:**

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `IV_` | `IV_COLUMN_INDEX` |
| **Object** | `IO_` | `IO_EXCEL_WRITER` |
| **Structure** | `IS_` | `IS_CELL_STYLE` |
| **Table** | `IT_` | `IT_DATA_ROWS` |

**EXPORTING Parameters:**

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `EV_` | `EV_SUCCESS_FLAG` |
| **Object** | `EO_` | `EO_NEW_WORKSHEET` |
| **Structure** | `ES_` | `ES_RESULT_DATA` |
| **Table** | `ET_` | `ET_ERROR_LOG` |

**RETURNING Parameters:**

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `RV_` | `RV_CELL_VALUE` |
| **Object** | `RO_` | `RO_WORKSHEET` |
| **Structure** | `RS_` | `RS_CELL_INFO` |
| **Table** | `RT_` | `RT_WORKSHEET_LIST` |
| **Data Reference** | `RR_` | `RR_DATA_REF` |

> **Note**: Don't use generic suffix "RESULT". Use specific names like `RO_WORKSHEET` for method `GET_ACTIVE_WORKSHEET`

**CHANGING Parameters:**

| Type | Prefix | Example |
|------|--------|---------|
| **Elementary** | `CV_` | `CV_COUNTER` |
| **Structure** | `CS_` | `CS_LAYOUT_INFO` |
| **Table** | `CT_` | `CT_MODIFIED_ROWS` |
| **Object** | `CO_` | `CO_EXCEL_OBJECT` |

### Alternative Naming Standards

Use these if they occur in existing method implementations and are the majority:

- **Types**: `T_`, `TY_`, `LTY`, `MTY_` were often used
- **Field Symbols**: `<FS_` and `<F_` were often used  
- **Elementary Parameters**: `P` instead of `V` (e.g., `IP_` or `EP_`)
- **RETURNING Parameters**: `E` instead of `R`
- **CHANGING Parameters**: `X` instead of `C`

---

## 🔍 Code Quality with abaplint

abap2xlsx uses [abaplint](https://abaplint.org/) for automated code quality checking and enforcement of coding standards.

### Configuration Files

The project maintains multiple abaplint configurations:

- **`abaplint.json`** - Main configuration for standard ABAP systems
- **`abaplint-steampunk.json`** - Configuration for SAP Cloud Platform/Steampunk

### Key Configuration Settings

#### Syntax Version

- **Standard Systems**: `v702` minimum
- **Cloud Systems**: `Cloud` version

#### Object Naming Rules

The project enforces specific naming patterns:

| Object Type | Pattern | Example |
|-------------|---------|---------|
| **Classes** | `^ZC(L\|X)\_EXCEL` | `ZCL_EXCEL`, `ZCX_EXCEL_ERROR` |
| **Interfaces** | `^ZIF\_EXCEL` | `ZIF_EXCEL_WRITER` |
| **Tables** | `^ZEXCEL` | `ZEXCEL_CELL_DATA` |
| **Table Types** | `^ZEXCEL` | `ZEXCEL_T_WORKSHEETS` |
| **Data Elements** | `^ZEXCEL` | `ZEXCEL_COLUMN_INDEX` |

### Manual Code Checking

#### Install abaplint CLI

```bash
npm install -g @abaplint/cli
```

#### Check Code Quality

```bash
# Check all files with main configuration
abaplint

# Check specific files
abaplint src/zcl_excel.clas.abap

# Check with specific configuration
abaplint -c abaplint-steampunk.json

# Output format options
abaplint -f json          # JSON output
abaplint -f checkstyle    # Checkstyle XML format
abaplint -f junit         # JUnit XML format
```

#### Check Specific Rules

```bash
# Check only naming conventions
abaplint --only object_naming,keyword_case

# Exclude specific rules
abaplint --exclude unused_variables,line_length
```

### Automated Code Fixing

#### Fix Automatically Correctable Issues

```bash
# Fix keyword casing
abaplint --fix keyword_case

# Fix whitespace issues
abaplint --fix whitespace_end,sequential_blank

# Fix multiple issues at once
abaplint --fix keyword_case,whitespace_end,colon_missing_space
```

#### Common Auto-Fixable Rules

| Rule | Description | Fix Command |
|------|-------------|-------------|
| `keyword_case` | ABAP keyword casing | `--fix keyword_case` |
| `whitespace_end` | Trailing whitespace | `--fix whitespace_end` |
| `sequential_blank` | Multiple blank lines | `--fix sequential_blank` |
| `colon_missing_space` | Space after colon | `--fix colon_missing_space` |
| `space_before_colon` | Space before colon | `--fix space_before_colon` |

### Key Enabled Rules

#### Critical Rules (Always Enabled)

- **`dangerous_statement`** - Detects potentially harmful statements
- **`method_implemented_twice`** - Prevents duplicate method implementations  
- **`uncaught_exception`** - Ensures proper exception handling
- **`parser_error`** - Catches syntax errors
- **`unknown_types`** - Validates type references

#### Code Quality Rules

- **`superfluous_value`** - Removes unnecessary VALUE statements
- **`prefer_corresponding`** - Encourages CORRESPONDING usage
- **`omit_preceding_zeros`** - Removes leading zeros in numbers
- **`prefer_raise_exception_new`** - Modern exception raising

#### Naming and Style Rules

- **`keyword_case`** - Enforces UPPER case keywords
- **`names_no_dash`** - Prevents dashes in names
- **`object_naming`** - Enforces project naming conventions

### Integration with Development Workflow

#### Pre-commit Checks

```bash
# Add to your git pre-commit hook
#!/bin/sh
abaplint --exit-code
if [ $? -ne 0 ]; then
    echo "abaplint checks failed. Please fix issues before committing."
    exit 1
fi
```

#### CI/CD Integration

```yaml
# GitHub Actions example
- name: Run abaplint
  run: |
    npm install -g @abaplint/cli
    abaplint -f checkstyle > abaplint-results.xml
```

---

## 📐 Development Standards

### Code Structure

- **Clean Code Principles**: Follow SAP's Clean ABAP guidelines
- **Single Responsibility**: Each class/method should have one clear purpose
- **Meaningful Names**: Use descriptive names that explain intent

### Error Handling

- **Use Class-Based Exceptions**: Prefer `ZCX_*` exception classes
- **Proper Exception Propagation**: Don't catch and ignore exceptions
- **Meaningful Error Messages**: Provide context for debugging

### Performance Considerations

- **Database Access**: Minimize database calls, use efficient SELECT statements
- **Memory Management**: Clean up object references when no longer needed
- **Large Data Sets**: Use `ZCL_EXCEL_WRITER_HUGE_FILE` for large Excel files

### Documentation

- **Method Documentation**: Document complex business logic
- **Class Documentation**: Explain class purpose and usage patterns
- **Public API**: Maintain comprehensive documentation for public interfaces

---

## ✅ Best Practices

### Before Submitting Code

1. **Run abaplint checks**: `abaplint` - ensure no critical issues
2. **Fix auto-correctable issues**: `abaplint --fix keyword_case,whitespace_end`
3. **Test with demo programs**: Verify functionality works correctly
4. **Follow naming conventions**: Use appropriate prefixes for new code
5. **Check existing patterns**: Maintain consistency with surrounding code

### Code Review Checklist

- [ ] abaplint passes without critical errors
- [ ] Naming conventions followed consistently
- [ ] No hardcoded values (use constants)
- [ ] Proper error handling implemented
- [ ] Performance considerations addressed
- [ ] Documentation updated if needed

### Continuous Improvement

- **Monitor Rule Updates**: Keep abaplint configuration current
- **Community Feedback**: Incorporate suggestions from code reviews
- **Performance Metrics**: Track and improve code quality over time

---

## 🔗 Additional Resources

- **[abaplint Documentation](https://docs.abaplint.org/)** - Complete abaplint reference
- **[Clean ABAP Guidelines](https://github.com/SAP/styleguides/blob/main/clean-abap/CleanABAP.md)** - SAP's official style guide
- **[Contributing Guidelines](../CONTRIBUTING.md)** - Project contribution process
- **[GitHub Issue #773](https://github.com/abap2xlsx/abap2xlsx/issues/773)** - Naming standards discussion
