# VBSdb — Enterprise Database GUI & CRUD Engine for Classic ASP (2001–2003)

[![Version](https://img.shields.io/badge/Version-Pro--Revision.2003.12-blue.svg)](#)
[![Tech Stack](https://img.shields.io/badge/Tech_Stack-Classic_ASP_%7C_VBScript_%7C_ADO-orange.svg)](#)
[![License](https://img.shields.io/badge/License-Open_Source-green.svg)](#)

> [!NOTE]
> **VBSdb** was a commercial software product developed, packaged, and sold between 2001 and 2003 via **`vbsdb.com`**. This repository contains the complete codebase of **VBSdb Pro (Revision 2003.12)**, now published as open source to showcase early web framework development, security-first engineering, and software commercialization.

---

## 📖 The Story of VBSdb
In the early 2000s, during the peak of Classic ASP and IIS web development, building CRUD (Create, Read, Update, Delete) interfaces was a highly repetitive, error-prone, and time-consuming process. Developers spent days manually writing HTML tables, paginating ADO recordsets, implementing client-side JavaScript validation, and writing safe SQL `INSERT`/`UPDATE` routines.

**VBSdb** was engineered to solve this. It was a component-oriented Classic ASP database engine that allowed developers to build fully featured database administration screens with a few lines of configuration code. Sold as a commercial package, it powered numerous back-offices and content management systems worldwide.

---

## 🌟 Key Features (Ahead of its Time in 2002)

### 1. Zero-Code UI Generation
*   **Automatic Grid & Form Generation**: Renders read-only grids (`GRID`), individual record forms (`FORM`), or split-screen layouts (`GRID-FORM`).
*   **Intelligent Pagination**: Hand-rolled, highly optimized record navigation (First, Prev, Next, Last) with page position labels (e.g., *"Records 11 to 20 of 500"*).
*   **Sortable Headers**: Instant column sorting (ascending/descending) achieved dynamically by intercepting and adjusting SQL order clauses.
*   **Horizontal Striping & Custom CSS**: A clean, customizable aesthetic with alternating row stripes and structural styling using CSS (see [VbsDb.css](file:///home/mionome/Documents/yamma/VBSdb/VbsDb.css)).

### 2. Security-First Architecture (Anti-Form Tampering)
Long before modern web frameworks automated request security, VBSdb implemented a dedicated security layer to prevent malicious user input and parameter tampering (see [inc_EditFormTamperingControl.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_EditFormTamperingControl.asp)):
*   **Operation Grant Checks**: Ensures that `ADD`, `UPDATE`, or `DELETE` requests can only execute if the developer explicitly configured that action as allowed in the navigation buttons.
*   **Dropdown/Option Validation**: Verifies that submitted values for dynamic drop-down lists (`InputSelectFields`) or radio/checkboxes (`InputEnumeratedFields`) match allowed choices, preventing clients from injecting unlisted options via form manipulation.
*   **Query Scope Validation**: Ensures the record being updated or deleted is within the scope of the developer's original query (e.g., preventing users from modifying another user's records by tampering with the primary key in hidden fields).

### 3. Native Multi-Database Support
VBSdb utilized **ADO (ActiveX Data Objects)** and **ADOX (ActiveX Data Objects Extensions)** to inspect database schemas at runtime. It automatically adapted its query syntax for:
*   **Microsoft Access (`.mdb`)** via JET OLEDB (4.0 & 3.51) and ODBC
*   **Microsoft SQL Server**
*   **Oracle**
*   **FoxPro**
*   Custom drivers via configurable connection strings

### 4. Advanced Validation & Input Control
*   **Regular Expression Validation**: Developers could define complex regex checks per field with customizable user alerts (see [inc_EditValidateRegExp.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_EditValidateRegExp.asp)).
*   **Required Fields**: Built-in enforcement of mandatory fields with client-side JavaScript alert hooks.
*   **Autoincrement Parsing**: Automated discovery of primary key identity fields across databases to prevent editing auto-assigned primary keys.
*   **Date Processing**: Specialized date parsing and locale-aware date-formatting rules.

### 5. Multi-Language Localization
Out of the box, VBSdb supported localization across 9 languages. It mapped all labels, navigation text, validation warnings, and titles based on the system configuration:
*   🇩🇰 Danish &nbsp;|&nbsp; 🇳🇱 Dutch &nbsp;|&nbsp; 🇬🇧 English &nbsp;|&nbsp; 🇫🇷 French &nbsp;|&nbsp; 🇩🇪 German &nbsp;|&nbsp; 🇬🇷 Greek &nbsp;|&nbsp; 🇮🇹 Italian &nbsp;|&nbsp; 🇵🇹 Portuguese &nbsp;|&nbsp; 🇪🇸 Spanish
*   Check the language packs folder for implementation details: [languages/](file:///home/mionome/Documents/yamma/VBSdb/languages)

---

## 🛠️ How It Works

A complete CRUD, paginated, and validated view could be generated in a single Classic ASP script using VBSdb:

```asp
<!--#include file="inc.asp"-->
<%
  Dim objDb
  
  ' 1. Initialize the VBSdb dictionary object
  VbsDbNew objDb
  
  ' 2. Configure Database Connection & Target Query
  objDb("MdbPath") = "databases/inventory.mdb" ' MapPath is resolved automatically
  objDb("Sql") = "SELECT * FROM Products WHERE Discontinued = False"
  objDb("GlobalId") = "inventory_manager"       ' Unique ID for session tracking
  
  ' 3. Define GUI Layout Mode & Controls
  objDb("ViewMode") = "GRID-FORM"               ' Split screen: Grid list on top, Form on bottom
  objDb("GridPageSize") = 10                    ' Paginates at 10 items
  objDb("ViewNavigationButtons") = "FIRST;PREV;NEXT;LAST;SEARCH;ADD;UPDATE;DELETE"
  
  ' 4. Configure Database Write-backs
  objDb("EditTableName") = "Products"
  objDb("EditKeyFields") = "ProductId"          ' Primary key
  
  ' 5. Configure Dynamic Input Options (Select drop-down from query)
  objDb("InputSelectFields") = "SupplierId|SELECT SupplierId, SupplierName FROM Suppliers"
  
  ' 6. Define Custom Field Labels
  objDb("GlobalFieldHeaders") = "ProductId|ID;ProductName|Product Name;UnitPrice|Price"
  
  ' 7. Secure the Form & Execute Renders
  objDb("EditFormTamperingControl") = True      ' Enforces strict validation
  
  ' Render UI & Process incoming forms
  VbsDb objDb
  
  ' 8. Close connection & free memory
  VbsDbClose objDb
%>
```

---

## 🔍 Architecture & File Breakdown

The project follows a modular structure where specific behaviors are split into include files. Here are the key subsystems:

| Core System | Path & Link | Description |
| :--- | :--- | :--- |
| **Bootstrapper** | [inc.asp](file:///home/mionome/Documents/yamma/VBSdb/inc.asp) | The master include file which groups and imports all core subsystems. |
| **Main Engine** | [inc_vbsDb.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_vbsDb.asp) | Manages initialization (`VbsDbNew`), connection acquisition, dictionary mapping, and session parameters. |
| **Security Layer** | [inc_EditFormTamperingControl.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_EditFormTamperingControl.asp) | Prevents unauthorized queries, values, and database operations via request validation. |
| **Database Writer** | [inc_vbsDbEditDb.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_vbsDbEditDb.asp) | Handles SQL transaction generation, field conversions, quote escaping, and execution. |
| **Grid Renderer** | [inc_vbsDbDrawGrid.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_vbsDbDrawGrid.asp) | Renders the HTML table grids, sorting hooks, and inline action buttons. |
| **Form Renderer** | [inc_vbsDbDrawInputScreen.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_vbsDbDrawInputScreen.asp) | Dynamically draws add/edit forms, binding drop-downs, memo fields, and check boxes. |
| **Robust Debugger** | [inc_vbsDbErrors.asp](file:///home/mionome/Documents/yamma/VBSdb/inc_vbsDbErrors.asp) | A massive module providing detailed schema debugging logs, ADO versions, and server environmental variables. |

---

## 💼 Portfolio Highlights & Relevance Today

Although Classic ASP has been succeeded by modern frameworks, the code in **VBSdb** demonstrates several foundational engineering skills that are highly relevant to modern enterprise software development:

*   **Framework & API Design**: Creating a developer-friendly framework that simplifies complex database integrations down to declarative configuration options (a precursor to modern low-code/CRUD frameworks).
*   **Active Directory & ADO Internals**: Direct interactions with database system tables (`ADODB.Recordset` and `ADOX.Catalog` schema queries) to extract table layouts dynamically.
*   **Web Security Best Practices**: Early foresight into web security vulnerabilities, actively mitigating form tampering and option validation attacks prior to the widespread adoption of CSRF, SQLi, and parameters validation libraries.
*   **Localization (L10n/i18n)**: Architecting a clean, pluggable dictionary-based translation system supporting 9 languages.
*   **Product Lifecycle & Launch**: The experience of developing, marketing, licensing, and commercially supporting a software-as-a-service/library product from scratch.
