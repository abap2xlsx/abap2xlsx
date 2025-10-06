# 📊 abap2xlsx - Read and Generate Excel Spreadsheets with ABAP

[![GitHub release](https://img.shields.io/github/release/abap2xlsx/abap2xlsx.svg)](https://github.com/abap2xlsx/abap2xlsx/releases/latest)
[![License](https://img.shields.io/github/license/abap2xlsx/abap2xlsx.svg)](LICENSE)
[![SAP Community](https://img.shields.io/badge/SAP%20Community-abap2xlsx-blue)](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)
[![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/abap2xlsx/abap2xlsx)

> **A comprehensive ABAP library for reading and generating Excel spreadsheets (.xlsx) directly from SAP systems**

For general information please refer to the blog series [abap2xlsx - Generate your professional Excel spreadsheet from ABAP](http://scn.sap.com/community/abap/blog/2010/07/12/abap2xlsx--generate-your-professional-excel-spreadsheet-from-abap) and the [documentation](https://abap2xlsx.github.io/abap2xlsx/).

## 🚀 Quick Start

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer    TYPE REF TO zcl_excel_writer_2007.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello World!' ).

CREATE OBJECT lo_writer.
DATA(lv_xstring) = lo_writer->write_file( lo_excel ).
```

## 📋 Installation

**Version support:** Minimum tested version is SAP_ABA 731, it might work on older versions still but we need volunteers to test it.

### Installation Methods

- **🎯 [Beginner's Guide](GETTING_STARTED_FOR_BEGINNERS.md)** - Complete step-by-step installation
- **📘 [Official abapGit Installation Guide](https://abap2xlsx.github.io/abap2xlsx/abapGit-installation)** - Detailed abapGit setup
- **⚡ Quick Install:** Clone `https://github.com/abap2xlsx/abap2xlsx.git` via abapGit

### Demo Programs

Note that the **Demo programs** are provided in a [separate repository](https://github.com/abap2xlsx/demos), and can be installed after abap2xlsx.

## ✨ Key Features

| Category | Capabilities |
|----------|--------------|
| **📊 Excel Generation** | .xlsx, .xlsm, .csv formats with advanced formatting |
| **📈 Data Integration** | Internal table binding, ALV conversion, template filling |
| **🎨 Advanced Features** | Conditional formatting, autofilters, charts, named ranges |

## 📖 Documentation & Resources

- **📘 [Complete Documentation](https://abap2xlsx.github.io/abap2xlsx/)** - Official documentation site
- **❓ [FAQ](docs/FAQ.md)** - Common solutions and troubleshooting
- **📝 [Blog Series](http://scn.sap.com/community/abap/blog/2010/07/12/abap2xlsx--generate-your-professional-excel-spreadsheet-from-abap)** - Original tutorials

## 🆘 Support & Community

- **💬 [SAP Community](https://community.sap.com/t5/forums/searchpage/tab/message?q=abap2xlsx)** - General questions
- **💬 [Slack #abap2xlsx](https://sapmentors.slack.com/archives/CGG0UHDMG)** - Real-time chat
- **🐛 [GitHub Issues](https://github.com/abap2xlsx/abap2xlsx/issues)** - Bug reports & features

## 🤝 Contributing

For questions, bug reports and more information on contributing to the project, please refer to the [contributing guidelines](./CONTRIBUTING.md).

## 📄 License

Licensed under the Apache License, Version 2.0. See the [LICENSE](LICENSE) file for details.

---

**Ready to start?** 🚀 Check the [documentation](https://abap2xlsx.github.io/abap2xlsx/) or explore [demo programs](https://github.com/abap2xlsx/demos)!
