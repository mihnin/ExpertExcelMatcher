# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based fuzzy string matching application called "Expert Excel Matcher v3.0.0" designed to match software product names between two data sources (Excel or CSV). The application uses a tkinter GUI and implements multiple fuzzy matching algorithms to find the best correspondences between product names.

**Version**: 3.0.0 (Post-Critical-Fixes)
**Last Updated**: 2025-10-23
**Architecture**: Modular (8 modules in src/, ~1,250 lines in main file, down from 2,739)

**Key Features**:
- **Universal file support**: Excel (.xlsx, .xls) and CSV files with automatic encoding detection (UTF-8-BOM, UTF-8, CP1251, Windows-1251, Latin1)
- Tests ALL available methods (not just top 5)
- Dynamic time estimation based on method types
- Adaptive UI with scrolling support and intuitive description of capabilities
- **Multiple column comparison (1-2 columns) with auto-mode checkbox**
- Custom column selection and inheritance
- Built-in Exact Match (ВПР) method
- Lexicographic ranking algorithm for method selection
- **Advanced normalization with real-time engine updates**
- **Debug columns showing normalized values in Excel export**
- Simplified file validation (any data types accepted)
- Refactored codebase with helper methods and constants

## Commands

### Installation
```bash
# Install from requirements.txt (recommended)
pip install -r requirements.txt

# Or install manually
pip install pandas openpyxl xlsxwriter rapidfuzz textdistance jellyfish transliterate
```

### Running the Application
```bash
python expert_matcher.py
```

### Testing
```bash
# Run all tests (33 tests, 32 should pass)
python -m pytest tests/ -v

# Run specific test module
python -m pytest tests/test_matching.py -v

# Check syntax
python -m py_compile expert_matcher.py
```

### Building Executable
```bash
# Clean previous builds
rm -rf build dist

# Optimized build (recommended, ~78 MB)
python -m PyInstaller --clean --noconfirm ExpertExcelMatcher_optimized.spec

# Basic build (larger, ~100 MB)
pyinstaller --onefile --windowed --name "ExpertExcelMatcher" expert_matcher.py

# Output: dist/ExpertExcelMatcher.exe
```

### Validation Scripts
```bash
# Test length penalty mechanism
python test_improvements.py

# Validate report accuracy after changes
python check_report.py
```

## Architecture (Post-Refactoring v2.2)

The codebase has been refactored into a modular architecture with 8 specialized modules in `src/`:

### Module Structure

1. **src/constants.py** (76 lines)
   - `AppConstants`: Version, UI text, colors, fonts
   - `NormalizationConstants`: Regex patterns for legal forms, versions, stop words

2. **src/models.py** (294 lines)
   - `MatchingMethod`: Encapsulates a fuzzy matching algorithm
     - Key attributes: `name`, `func`, `library`, `use_process`, `scorer`
     - Main method: `find_best_match()` - finds best match with length penalty
     - Supports RapidFuzz optimized mode (100x faster) or manual iteration
   - `MatchResult`: Individual match result with score
   - `MethodStatistics`: Statistics for a method (100%, 90-99%, etc.)

3. **src/matching_engine.py** (191 lines)
   - `NormalizationOptions`: Configuration for text normalization
   - `MatchingEngine`: Core fuzzy matching logic
     - `normalize_string()`: Lowercase, whitespace, legal forms, versions, transliteration
     - `calculate_statistics()`: Non-cumulative category counting (CRITICAL: each record counted once)
     - Implements 50% score cutoff threshold

4. **src/data_manager.py** (273 lines)
   - `DataManager`: Universal file I/O and data management
     - `read_data_file()`: Reads Excel (.xlsx, .xls) or CSV with auto-encoding detection
     - `validate_file()`: Checks for empty files, missing columns, insufficient data
     - `set_source1_file()` / `set_source2_file()`: File selection with validation
     - Column selection and inheritance logic

5. **src/excel_exporter.py** (414 lines)
   - `ExcelExporter`: Consolidated export functionality (was 7 separate functions)
     - Full report with statistics sheet
     - Filtered exports: 100% matches, <90% matches, 0% matches
     - Color-coded output (green=perfect, red=no match)
     - Comparison table export

6. **src/ui_components.py** (399 lines)
   - Reusable UI widgets (eliminates duplication):
     - `TreeviewWithScrollbar`: Treeview with Y/X scrollbars
     - `ScrollableFrame`: Frame with vertical scrolling
     - `MethodSelectorListbox`: Listbox for method selection
     - Helper functions: `create_label_frame`, `create_title_header`, etc.

7. **src/ui_manager.py** (590 lines)
   - `UIManager`: Creates all GUI tabs (delegated from ExpertMatcher)
     - `create_setup_tab()`: File selection, mode selection, normalization options
     - `create_comparison_tab()`: Method comparison TreeView
     - `create_results_tab()`: Results display with export buttons
     - `create_help_tab()`: Scrollable help documentation
     - Event handlers: column selection, method selection, normalization toggles

8. **src/help_content.py** (443 lines)
   - `HelpContent`: Static class with all help text methods
     - `get_file_requirements()`, `get_modes_description()`, etc.
     - Keeps UI code clean by separating documentation

### ExpertMatcher Class (expert_matcher.py, 1,263 lines)

Main application controller, now **53.9% smaller** after refactoring:
- Delegates to specialized managers: `self.engine`, `self.data_manager`, `self.exporter`, `self.ui_manager`
- **Four operational modes**:
  - **Auto**: Tests selected methods on sample, picks best via lexicographic ranking
  - **Compare**: Tests on sample (≤200 records), displays metrics
  - **Full Compare**: Tests on ALL data, exports comparison Excel
  - **Manual**: User selects specific method(s)
- **Note**: Auto and Compare use IDENTICAL ranking logic (100% matches > 90-99% > avg score)
- **Important**: No method is hardcoded as "recommended" - winner determined by actual testing

### Key Workflows

1. **Method Registration** (`register_all_methods()`)
   - Dynamically registers methods based on installed libraries (RapidFuzz, TextDistance, Jellyfish)
   - **No hardcoded "recommended" methods** - all treated equally until testing
   - Returns list of `MatchingMethod` objects

2. **Statistics Calculation** (`MatchingEngine.calculate_statistics()`)
   - **CRITICAL**: Non-cumulative counting (each record counted exactly once)
   - Categories: 100%, 90-99%, 70-89%, 50-69%, 1-49%, 0%
   - Validation: sum of categories must equal total records
   - Located in `src/matching_engine.py` (moved from main file during refactoring)

3. **Matching Pipeline** (`ExpertMatcher.start_processing()`)
   - Preprocessing: Normalize all source2 names once via `engine.normalize_string()`
   - Create normalized → original name mapping
   - For each source1 record:
     - Normalize query string with selected normalization options
     - Call `method.find_best_match()` which uses:
       - RapidFuzz `process.extractOne()` (optimized, 100x faster)
       - Or manual iteration for TextDistance/Jellyfish
     - Apply length penalty to prevent false matches
     - Apply 50% score cutoff
   - Update progress bar with dynamic time estimates

4. **Delegation Pattern** (ExpertMatcher uses managers)
   - `self.engine = MatchingEngine(options)` - normalization & statistics
   - `self.data_manager = DataManager()` - file I/O
   - `self.exporter = ExcelExporter(engine, results)` - Excel generation
   - `self.ui_manager = UIManager(self)` - GUI creation
   - This pattern isolates responsibilities and simplifies testing

### GUI Structure (Notebook with 4 tabs)

Created by `UIManager.create_widgets()`:
1. **Setup Tab** (`create_setup_tab`): File selection, column selection, mode radio buttons, normalization checkboxes
   - **Multi-column mode checkbox**: Auto-enables when 2 columns selected in both sources
2. **Comparison Tab** (`create_comparison_tab`): TreeView showing all methods ranked by performance
3. **Results Tab** (`create_results_tab`): Match results TreeView, statistics, 4 export buttons
   - **Displays only comparison columns** (not inherited columns) for clarity
4. **Help Tab** (`create_help_tab`): Scrollable canvas with comprehensive documentation

**UI delegates event handling back to ExpertMatcher**: file selection, processing start, export actions

**CRITICAL**: Event handlers in UIManager (`on_askupo_column_select`, `on_eatool_column_select`) must synchronize selections with `data_manager` to prevent column mismatch bugs

### String Normalization (`MatchingEngine.normalize_string()`)
- **Location**: `src/matching_engine.py`
- **Configurable via NormalizationOptions**:
  - Remove legal forms (ООО, Ltd, Inc, GmbH, etc.)
  - Remove versions (2021, v4.x, x64, SP1, etc.)
  - Remove stop words (и, в, на, the, a, and, etc.)
  - Transliterate Cyrillic → Latin (Фотошоп → Fotoshop)
  - Remove punctuation
- **Always applied**: lowercase, strip whitespace, collapse multiple spaces
- Handles None/NaN gracefully

### Length Penalty Mechanism (`MatchingMethod.find_best_match()`)
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

## Performance Optimizations (v3.0.0 - 2025-10-23)

### 1. Exact Match (ВПР) - O(1) Dictionary Lookup
**Problem**: Method used O(N²) iteration through all choices for each query
- For 1000×1000 records: 1,000,000 iterations + 2,000,000 normalizations
- Time: ~30 seconds for 1000 records

**Solution**: Implemented `is_exact_match=True` flag in `MatchingMethod` (models.py:50)
```python
# In find_best_match() - lines 69-75
if self.is_exact_match:
    # O(1) dictionary lookup instead of O(N) iteration
    if query in choice_dict:
        return choice_dict[query], 100.0
    else:
        return "", 0.0
```

**Registration**: expert_matcher.py:219-224
```python
MatchingMethod("Exact Match (ВПР)",
              self.exact_match_func, "builtin",
              use_process=False, scorer=None, is_exact_match=True)
```

**Impact**: 30× speedup - from ~30 seconds to <1 second for 1000 records

### 2. Eliminated Duplicate Source2 Iterations
**Problem**: Code iterated through source2 DataFrame TWICE in preprocessing (expert_matcher.py:956-970)
- First loop: build combined names list
- Second loop: build row dictionary (DUPLICATE work)

**Solution**: Merged into single loop (expert_matcher.py:952-967, 1039-1050)
```python
# BEFORE: Two separate loops
for _, row in eatool_df.iterrows():
    combined = self.engine.combine_columns(row, eatool_cols)
    eatool_combined_names.append(combined)

for idx, row in eatool_df.iterrows():  # DUPLICATE!
    combined = self.engine.combine_columns(row, eatool_cols)
    eatool_row_dict[combined] = row

# AFTER: Single combined loop
for _, row in eatool_df.iterrows():
    combined = self.engine.combine_columns(row, eatool_cols)
    eatool_combined_names.append(combined)
    eatool_row_dict[combined] = row  # Fill dictionary in same pass
```

**Impact**: 2× speedup for data preprocessing phase

### 3. Performance Characteristics by Method Type

| Method Type | Performance | Optimization |
|------------|-------------|--------------|
| **Exact Match (ВПР)** | ⚡ INSTANT | ✅ O(1) dictionary (v3.0.0) |
| **RapidFuzz (10 methods)** | ⚡ VERY FAST | ✅ C++ `process.extractOne()` |
| **TextDistance (5 methods)** | ⚠️ SLOW | ❌ No batch API available |
| **Jellyfish (2 methods)** | ⚠️ SLOW | ❌ No batch API available |

**Time Estimates** (per 1000 records):
- Exact Match: <1 second (after optimization)
- RapidFuzz: 2-3 seconds per method
- TextDistance/Jellyfish: 15-30 seconds per method

**Key Insight**: The `is_exact_match` flag completely bypasses the scoring function - lookup happens before any string comparison, making it the fastest possible implementation.

## CRITICAL Bug Fixes (v3.0.0 - 2025-10-23)

### 1. Normalization Engine Not Updating
**Problem**: The `MatchingEngine` was created once in `__init__` and never updated when normalization checkboxes changed.
**Solution**: Added `self._update_matching_engine()` call at the start of `start_processing()` (expert_matcher.py:509)
**Impact**: Normalization checkboxes now ACTUALLY affect matching results

### 2. Column Selection Synchronization
**Problem**: When columns were selected in GUI, only `self.selected_askupo_cols` was updated, but `data_manager.selected_source1_cols` remained with the default first column. This caused all methods except the first to use WRONG columns.
**Solution**: Added synchronization in `on_askupo_column_select()` and `on_eatool_column_select()` (ui_manager.py:540, 559):
```python
self.parent.data_manager.selected_source1_cols = self.parent.selected_askupo_cols
self.parent.data_manager.selected_source2_cols = self.parent.selected_eatool_cols
```
**Impact**: ALL methods now use the same selected columns consistently

### 3. Multi-Column Mode Auto-Detection
**Problem**: The multi-column checkbox didn't automatically reflect the actual state
**Solution**: Added auto-mode logic that enables/disables checkbox based on actual column selection (ui_manager.py:553-557, 578-582)
**Impact**: Checkbox now automatically turns on when 2 columns selected in BOTH sources

### 4. Display Only Comparison Columns
**Problem**: Results TreeView showed ALL columns (including inherited), not just the 2 selected for comparison
**Solution**: Modified `display_results()` to filter only comparison columns by building column names from `selected_askupo_cols` (expert_matcher.py:1197-1205)
**Impact**: Results now clearly show only the columns that were actually compared

### 5. Debug Columns in Excel Export
**Added Feature**: Two new columns in Excel exports show normalized values:
- `[DEBUG] Нормализованный Источник 1` - shows combined+normalized source 1
- `[DEBUG] Нормализованный Источник 2` - shows combined+normalized source 2
- **Yellow highlighting** with italic formatting for easy identification
**Impact**: Users can verify how normalization transforms their data

### 6. Removed Duplicate Handlers
**Problem**: Duplicate `on_askupo_column_select` and `on_eatool_column_select` methods in expert_matcher.py (unused)
**Solution**: Removed duplicates, kept only working handlers in UIManager
**Impact**: Cleaner codebase, less confusion

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
   - **Lexicographic sorting** (same for Auto and Compare modes):
     - Priority 1: Maximum 100% matches (perfect)
     - Priority 2: Maximum 90-99% matches (high)
     - Priority 3: Maximum average score
   - Sorting tuple: `(-perfect_count, -high_count, -avg_score)`
   - Example: Method(10, 32, 58.1%) beats Method(4, 1, 61.7%) because 10 > 4 in first position
   - **IMPORTANT**: No method is pre-labeled as "recommended" - winner determined by actual test results
   - This was corrected in v2.2 (previously WRatio was hardcoded as "recommended")

6. **Matching Libraries**
   - **RapidFuzz** (fastest, C++): WRatio, Token Set, Token Sort, Partial Ratio, Ratio, QRatio
   - **TextDistance** (scientific): Jaro-Winkler, Jaro, Jaccard, Sorensen-Dice, Cosine
   - **Jellyfish** (phonetic): Jaro-Winkler, Jaro
   - **Built-in**: Exact Match (ВПР) - returns 100% or 0% only

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

### Important Conventions
- **Application text**: Russian (UI labels, messages, comments in docstrings)
- **Main entry point**: `main()` function at end of expert_matcher.py
- **Imports**: All src modules imported at top of expert_matcher.py
- **Delegation pattern**: ExpertMatcher delegates to 4 managers (engine, data_manager, exporter, ui_manager)
- **No hardcoded recommendations**: Methods ranked only by test results, not by constants

### Code Organization After Refactoring
- **expert_matcher.py** (1,263 lines): Main controller, event handlers, processing logic
- **src/** (8 modules, 2,685 lines): Specialized functionality
- **tests/** (4 files, ~450 lines): 33 unit tests covering normalization, statistics, matching, data management
- **Build**: Use `ExpertExcelMatcher_optimized.spec` for minimal size (~78 MB vs ~100 MB basic build)

### When Modifying Code
- **Statistics changes**: Always run `check_report.py` to validate accuracy
- **Normalization changes**: Update `NormalizationConstants` in `src/constants.py`
  - **CRITICAL**: If modifying `NormalizationOptions`, ensure `_update_matching_engine()` is called before processing
- **UI changes**: Modify `src/ui_manager.py`, not expert_matcher.py
  - **Column selection handlers**: Always synchronize with `data_manager` (see bug fix #2)
- **Export changes**: Modify `src/excel_exporter.py`
  - **Excel formatting**: Check `_apply_color_coding()` for special column highlighting
- **New matching logic**: Update `src/matching_engine.py`
- **Run tests**: `python -m pytest tests/ -v` after changes
- **NEVER use emojis in Excel column names** - causes file corruption

### Performance Critical Patterns (DO NOT BREAK!)

**1. Exact Match Optimization** (models.py:69-75)
- The `is_exact_match` flag enables O(1) dictionary lookup
- NEVER remove this optimization - it provides 30× speedup
- Pattern: Check flag BEFORE calling scoring function
- If adding new built-in methods, consider if they can use this pattern

**2. Avoid Duplicate DataFrame Iterations**
- ALWAYS combine multiple operations on same DataFrame into single loop
- Pattern seen in expert_matcher.py:952-967 and 1039-1050
- Bad: `for row in df.iterrows(): do_x()` then `for row in df.iterrows(): do_y()`
- Good: `for row in df.iterrows(): do_x(); do_y()`
- Each DataFrame iteration is expensive - minimize passes

**3. Preprocessing Source2 Once**
- Source2 normalization happens ONCE before main loop (expert_matcher.py:966)
- `eatool_normalized` list created upfront
- `choice_dict` maps normalized→original
- NEVER normalize inside the source1 loop - kills performance

**4. RapidFuzz vs Manual Iteration**
- RapidFuzz methods use `process.extractOne()` (100× faster)
- Check `use_process=True` flag in MatchingMethod
- TextDistance/Jellyfish must use manual iteration (no batch API)
- When adding new methods, prefer RapidFuzz if available

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
