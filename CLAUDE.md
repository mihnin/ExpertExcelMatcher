# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based fuzzy string matching application called "Expert Excel Matcher v2.2" designed to match software product names between two data sources (Excel or CSV). The application uses a tkinter GUI and implements multiple fuzzy matching algorithms to find the best correspondences between product names.

**Version**: 2.2.0
**Last Updated**: 2025-10-22

**Key Features**:
- **Universal file support**: Excel (.xlsx, .xls) and CSV files with automatic encoding detection (UTF-8-BOM, UTF-8, CP1251, Windows-1251, Latin1)
- Tests ALL available methods (not just top 5)
- Dynamic time estimation based on method types
- Adaptive UI with scrolling support and intuitive description of capabilities
- Multiple column comparison (1-2 columns)
- Custom column selection and inheritance
- Built-in Exact Match (ВПР) method
- Lexicographic ranking algorithm for method selection
- Simplified file validation (any data types accepted)
- Refactored codebase with helper methods and constants (v2.1)

## Commands

### Installation
```bash
pip install pandas openpyxl xlsxwriter rapidfuzz textdistance jellyfish
```

### Running the Application
```bash
python expert_matcher.py
```

### Building Executable
```bash
# Install PyInstaller
pip install pyinstaller

# Build single executable with GUI (recommended)
pyinstaller --onefile --windowed --name "ExpertExcelMatcher" expert_matcher.py

# Output will be in dist/ExpertExcelMatcher.exe
```

### Testing & Validation
```bash
# Test length penalty improvements
python test_improvements.py

# Validate report accuracy
python check_report.py
```

## Architecture

### Core Components

1. **MatchingMethod Class** (lines 53-104)
   - Encapsulates a single fuzzy matching algorithm
   - Key attributes: `name`, `func`, `library`, `use_process`, `scorer`
   - **Note**: `is_fast` parameter was REMOVED - now all methods are treated equally
   - Main method: `find_best_match()` - finds best match using the specific algorithm
   - Supports two modes:
     - RapidFuzz optimized mode using `process.extractOne()` (100x faster)
     - Manual iteration through all choices for other libraries

2. **ExpertMatcher Class** (lines 141-1621)
   - Main application controller managing the GUI and processing logic
   - Four operational modes:
     - **Auto mode**: Tests ALL available methods on a sample and selects the best using lexicographic ranking (100% > 90-99% > avg)
     - **Compare mode**: Tests ALL methods on sample data, displays comparison metrics, uses SAME ranking logic as Auto
     - **Full Compare mode**: Applies ALL methods to ALL data and exports comprehensive Excel report (30-60 min)
     - **Manual mode**: Uses a user-selected specific method
     - **Multi-manual mode**: Tests multiple user-selected methods and exports comparison
   - **Note**: Auto and Compare modes use IDENTICAL selection logic, ensuring Auto always picks the #1 method from Compare

3. **Universal File Reading** (`read_data_file()`, NEW in v2.1)
   - Universal method for reading both Excel and CSV files
   - **CSV encoding detection**: Tries multiple encodings automatically (UTF-8-sig, UTF-8, CP1251, Windows-1251, Latin1)
   - **Format detection**: Based on file extension (.csv vs .xlsx/.xls)
   - Used by all file reading operations: validation, column loading, data processing
   - Allows mixing formats: File 1 can be Excel, File 2 can be CSV
   - Example: `df = self.read_data_file(filename, nrows=100)` - works for any format

4. **Matching Libraries Integration**
   - **RapidFuzz** (RAPIDFUZZ_AVAILABLE): Primary library, fastest performance
     - WRatio (recommended), Token Set, Token Sort, Partial Ratio, Ratio, QRatio
   - **TextDistance** (TEXTDISTANCE_AVAILABLE): Scientific distance metrics
     - Jaro-Winkler, Jaro, Jaccard, Sorensen-Dice, Cosine
   - **Jellyfish** (JELLYFISH_AVAILABLE): Phonetic similarity algorithms
     - Jaro-Winkler, Jaro

### Key Workflows

1. **Method Registration** (`register_all_methods()`, lines 124-172)
   - Registers all available matching methods based on installed libraries
   - **No longer distinguishes between "fast" and "slow"** - all methods are equal
   - All methods are tested in auto/compare modes

2. **Statistics Calculation** (`calculate_statistics()`, lines 183-212)
   - **CRITICAL**: Uses non-cumulative categories (each record counted once)
   - Categories: 100%, 90-99%, 70-89%, 50-69%, 1-49%, 0%
   - Includes validation check: sum of categories must equal total records
   - This function was specifically corrected to fix statistical accuracy

3. **Optimized Matching Pipeline**
   - Preprocessing: Normalize all EA Tool names once (`normalize_string()`)
   - Create mapping dictionary: normalized → original names
   - For each АСКУПО record:
     - Normalize query string
     - Use RapidFuzz `process.extractOne()` for fast methods (single optimized call)
     - Apply 50% score cutoff (below 50% treated as no match)
   - Progress tracking with time estimates

### GUI Structure (Notebook with 4 tabs)

1. **Setup Tab**: File selection, library status, mode selection
2. **Comparison Tab**: Side-by-side method performance metrics (TreeView)
3. **Results Tab**: Final matching results with statistics and export options
4. **Help Tab**: Comprehensive user guide with file requirements and mode explanations

### String Normalization (`normalize_string()`, lines 221-228)
- Converts to lowercase
- Strips whitespace
- Collapses multiple spaces to single space
- Handles None/NaN values

### Length Penalty Mechanism (lines 84-103, 116-126)
- **Critical feature for accuracy**: Prevents short strings from incorrectly matching long strings
- Applied in `MatchingMethod.find_best_match()` after fuzzy score calculation
- Two penalty modes:
  - **Short strings (≤3 chars)**: Quadratic penalty `length_ratio²` - requires near-exact length match
  - **Long strings (>3 chars)**: Square root penalty `length_ratio^0.5` - more lenient
- Formula: `adjusted_score = raw_score × length_penalty`
- Matches with adjusted_score < 50% are rejected
- Example: "R" vs "NGINX" - prevents false 100% match from Partial Ratio

### Export Functionality
- Full report with statistics sheet
- Filtered exports: 100% matches, <90% matches, 0% matches
- Color-coded Excel output (green=perfect, red=no match)
- Comparison table export for method benchmarking

## Important Implementation Notes

1. **Performance Optimization & Time Estimation**
   - RapidFuzz's `process.extractOne()` is 100x faster than manual iteration
   - **Dynamic time estimation**: Calculated based on method types
     - RapidFuzz methods: ~2-3 seconds per method on sample data
     - Other methods (TextDistance, Jellyfish): ~15-30 seconds per method on sample data
   - Formula: `estimated_time = (rapidfuzz_count * base_time + other_count * slow_time) / 60`
   - Sample-based testing (150-200 records) for auto/compare modes
   - Full dataset processing: ~2-3 minutes regardless of method

2. **Statistical Accuracy**
   - The `calculate_statistics()` function at line 229 is the **corrected version**
   - Previous versions had cumulative counting bugs
   - Always verify `check_sum == total` when modifying statistics
   - Use `check_report.py` to validate report accuracy after changes

3. **Column Handling**
   - Application assumes first column of each Excel file contains product names
   - Column names are dynamically retrieved: `df.columns[0]`

4. **Error Handling**
   - Library availability checked at startup with clear status indicators
   - Score normalization: converts [0,1] range to [0,100] percentage
   - Try-catch blocks protect against malformed data in matching loops

5. **Method Selection Strategy (UNIFIED LOGIC)**
   - **Both Auto and Compare modes use IDENTICAL lexicographic sorting:**
     - Priority 1: Maximum 100% matches (perfect)
     - Priority 2: Maximum 90-99% matches (high)
     - Priority 3: Maximum average score
   - This ensures consistency: Auto mode always selects the same "best" method shown as #1 in Compare mode
   - **ALL methods are tested** (no filtering by speed anymore)
   - Users can see full comparison and choose based on accuracy, not just speed
   - Scoring tuple: `(perfect_count, high_count, avg_score)` - compared lexicographically
   - Example: Method with (50, 30, 85%) beats (48, 40, 90%) because 50 > 48 in first position

## File Expectations

- **Input File 1**: АСКУПО database (e.g., "Уникальные_ПО_продукты.xlsx" or "data.csv")
- **Input File 2**: EA Tool database (e.g., "EA Tool short name v1.xlsx" or "products.csv")
- **Supported formats**: Excel (.xlsx, .xls) or CSV (.csv) - can mix formats!
- **CSV encoding**: Automatic detection (UTF-8, UTF-8-BOM, CP1251, Windows-1251, Latin1)
- **CSV delimiter**: Comma (standard)
- Files can have different formats (e.g., File 1 = Excel, File 2 = CSV)
- Selected columns must contain software product names (strings)
- Files are validated via `validate_excel_file()` on selection
  - Uses universal `read_data_file()` method for both Excel and CSV
  - Checks for empty files, missing columns, insufficient text data
  - Requires minimum 3 text entries in first column

## Utility Scripts

### check_report.py
- **Purpose**: Validates report accuracy and completeness
- **Usage**: Run after generating reports to verify all records are processed
- **Checks**:
  - All АСКУПО records present in report
  - Statistical category counts are correct and non-overlapping
  - Identifies missing or duplicate records

### test_improvements.py
- **Purpose**: Tests the length penalty mechanism
- **Usage**: Run to verify short string matching behavior
- **Test case**: Ensures "R" doesn't falsely match longer product names
- **Expected**: All matches below 50% threshold should be rejected

## Development Notes

- Application text is in Russian (UI labels, messages, comments)
- Main entry point: `main()` function at line 1614
- GUI built with tkinter, uses ttk for advanced widgets
- Excel export uses xlsxwriter for formatting and color coding
- All matching operations use normalized strings internally, original strings preserved for output

## Project Recovery / Setup from GitHub

If cloning this project fresh from GitHub:

```bash
# 1. Clone repository
git clone <repo-url>
cd ExpertExcelMatcher

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the application
python expert_matcher.py

# 4. (Optional) Build executable
pip install pyinstaller
pyinstaller --onefile --windowed --name "ExpertExcelMatcher" expert_matcher.py
# Or use the existing .spec file:
pyinstaller ExpertExcelMatcher.spec
```

### What's in the repository:
- ✅ All Python source code
- ✅ `requirements.txt` with exact dependency versions
- ✅ `ExpertExcelMatcher.spec` for reproducible builds
- ✅ Documentation (README.md, CLAUDE.md, BUILD.md)
- ✅ Test/validation scripts
- ❌ Excel data files (excluded for privacy/size)
- ❌ Build artifacts (build/, dist/ folders)
- ❌ Local configuration (.claude/ folder)

### Files you'll need to add manually:
- Your Excel input files (АСКУПО and EA Tool databases)
- Place them in the project root or anywhere accessible
