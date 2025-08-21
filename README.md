<<<<<<< HEAD
# 🚀 Robust Excel Inner Join Application

A powerful, dynamic Java application that performs robust inner joins between Excel files with intelligent data handling and flexible file detection.

## ✨ Key Features

### 🎯 **Robust & Dynamic**
- **Automatic File Detection**: Automatically finds main data and vendor files from common naming patterns
- **Intelligent Join Key Detection**: Smart algorithm to find the best join columns automatically
- **Flexible Data Handling**: Works with different column orders, case variations, and data types
- **Main Data as Authority**: Always preserves main data column structure as the authoritative source

### 🔧 **Smart Capabilities**
- **Dynamic Column Ordering**: Result always matches main data column structure exactly
- **No Duplicate Columns**: Eliminates redundant data and prefixes
- **Data Type Flexibility**: Handles numeric, text, and mixed data types intelligently
- **Error Recovery**: Comprehensive error handling with helpful troubleshooting tips

### 📊 **Advanced Analytics**
- **Join Statistics**: Detailed reporting on match rates and data quality
- **Data Integrity Validation**: Checks data consistency before processing
- **Match Rate Analysis**: Provides insights into join success rates

## 🗂️ **Supported File Patterns**

### Main Data Files (Authoritative Source)
- `MainData.xlsx`
- `Call-Center-Sentiment-Sample-Data.xlsx`
- `main_data.xlsx`
- `maindata.xlsx`

### Vendor Data Files
- `Data_Vendor.xlsx`
- `vendor_data.xlsx`
- `vendordata.xlsx`
- `vendor.xlsx`

## 🔑 **Intelligent Join Key Detection**

The application automatically detects the best join key using this priority order:

1. **Preferred Columns**: `id`, `ID`, `Id`, `customer_id`, `customerid`, `CustomerId`
2. **Case-Insensitive Matching**: Handles variations in case
3. **ID-Containing Columns**: Any column containing "id"
4. **Name-Based Columns**: Columns containing "name"
5. **First Common Column**: Fallback to any common column

## 🎯 **How It Works**

### 1. **Dynamic File Detection**
```
🔍 Found main data file: MainData.xlsx
🔍 Found vendor data file: Data_Vendor.xlsx
```

### 2. **Intelligent Analysis**
```
📋 Main Data columns (9): [id, First Name, Last Name, PhoneNumber, Address, Skills, Trained Status, Company, Experience]
📋 Vendor Data columns (2): [id, First Name]
🔗 Common columns (2): [First Name, id]
➕ Extra vendor columns (0): []
```

### 3. **Smart Join Key Selection**
```
🔍 Intelligently detecting best join key...
✅ Found perfect match: id
🔑 Using join keys: Main[id] ↔ Vendor[id]
```

### 4. **Detailed Results**
```
📊 Join Statistics:
✅ Successful matches: 2
📝 Total main records: 5
⚠️ Records with null join keys: 0
📈 Match rate: 40.0%
```

## 🚀 **Running the Application**

### Prerequisites
- Java 11 or higher
- Maven 3.6 or higher
- Excel files in .xlsx format

### Quick Start
```bash
# Compile the application
mvn compile

# Run the application
mvn exec:java
```

### Output
- Creates timestamped result files: `RobustInnerJoinResult_YYYYMMDD_HHMMSS.xlsx`
- Maintains exact main data column structure
- No duplicate or prefixed columns

## 📁 **File Structure Requirements**

### Main Data File (Database Source)
- Acts as the authoritative data source
- Defines the final column structure
- All columns will be preserved in the result

### Vendor Data File
- Contains matching records to join with main data
- Must have at least one common column with main data
- Extra columns will be appended to the result

## 🎛️ **Advanced Features**

### 🔄 **Flexible Data Matching**
- **Exact Matching**: Direct value comparison
- **Case-Insensitive**: Handles text case variations
- **Numeric Comparison**: Handles floating-point precision
- **String Trimming**: Removes leading/trailing spaces

### 🛡️ **Error Handling**
- **File Not Found**: Clear error messages with suggestions
- **Invalid Data**: Graceful handling of corrupt or missing data
- **No Matches**: Detailed diagnostics for troubleshooting
- **Column Mismatch**: Intelligent column mapping

### 📈 **Data Quality Insights**
- **Integrity Validation**: Checks for consistent column structures
- **Null Key Detection**: Identifies records with missing join keys
- **Match Rate Analysis**: Percentage of successful joins
- **Performance Metrics**: Processing statistics

## 🔧 **Troubleshooting**

### Common Issues & Solutions

**❌ No matches found**
- Check if join key values exist in both files
- Verify data formatting (numbers vs text)
- Ensure case sensitivity matches

**❌ Files not detected**
- Verify file names match supported patterns
- Check file permissions and formats
- Ensure files are .xlsx format

**❌ Column structure issues**
- Main data file defines the authoritative structure
- Extra vendor columns will be automatically appended
- No manual column mapping required

## 📊 **Example Output Structure**

### Before (Raw Files)
**MainData.xlsx**: `[id, First Name, Last Name, PhoneNumber, Address, Skills, Trained Status, Company, Experience]`
**Data_Vendor.xlsx**: `[id, First Name]`

### After (Result)
**RobustInnerJoinResult.xlsx**: `[id, First Name, Last Name, PhoneNumber, Address, Skills, Trained Status, Company, Experience]`
- ✅ Exact main data column order preserved
- ✅ No duplicate columns
- ✅ No prefixes or suffixes
- ✅ Clean, database-ready structure

## 🎯 **Key Benefits**

1. **🔄 Adaptability**: Works with changing file structures automatically
2. **🎯 Accuracy**: Main data structure always preserved exactly
3. **🛡️ Reliability**: Comprehensive error handling and validation
4. **📊 Insights**: Detailed analytics and reporting
5. **⚡ Efficiency**: Optimized performance with smart algorithms
6. **🔧 Maintenance**: Self-configuring with minimal manual intervention

## 📚 **Technical Details**

- **Language**: Java 11+
- **Build Tool**: Maven
- **Excel Library**: Apache POI 5.4.0
- **Data Structure**: LinkedHashMap for order preservation
- **Join Algorithm**: Optimized inner join with intelligent key detection
- **Memory Management**: Efficient streaming for large files 
=======
# Data_Compare
>>>>>>> 9a911603777ab2a787abd5009b745b27027d795a
