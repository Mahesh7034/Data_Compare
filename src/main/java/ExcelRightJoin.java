import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelRightJoin {
    
    // Configuration constants
    private static final String DEFAULT_MAIN_DATA_FILE = "MainFile/MainData.xlsx";
    private static final String DEFAULT_VENDOR_DATA_FILE = "InputFolder/Data_Vendor.xlsx";
    private static final String[] POSSIBLE_MAIN_FILES = {
        "MainFile/MainData.xlsx"

    };
    private static final String[] POSSIBLE_VENDOR_FILES = {
        "InputFolder/Data_Vendor.xlsx"
    };
    private static final String[] PREFERRED_JOIN_COLUMNS = {"id", "ID", "Id", "customer_id", "customerid", "CustomerId"};
    
    public static void main(String[] args) {
        try {
            System.out.println("=== Excel Inner Join Application ===");
            System.out.println("Initializing dynamic file detection...");
            
            // Dynamically detect available files
            String mainDataFile = detectMainDataFile();
            String vendorDataFile = detectVendorDataFile();
            
            if (mainDataFile == null || vendorDataFile == null) {
                System.err.println("ERROR: Required Excel files not found!");
                System.err.println("Looking for main data files: " + Arrays.toString(POSSIBLE_MAIN_FILES));
                System.err.println("Looking for vendor data files: " + Arrays.toString(POSSIBLE_VENDOR_FILES));
                return;
            }
            
            System.out.println("✅ Main Data File: " + mainDataFile);
            System.out.println("✅ Vendor Data File: " + vendorDataFile);
            
            // Generate output filename with timestamp
            String timestamp = new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new java.util.Date());
            String outputFile = "OutputFolder/InnerJoinResult_" + timestamp + ".xlsx";
            
            System.out.println("\n=== Reading and Analyzing Files ===");
            
            // Read both Excel files with enhanced error handling
            List<Map<String, Object>> mainData = readExcelFile(mainDataFile);
            List<Map<String, Object>> vendorData = readExcelFile(vendorDataFile);
            
            if (mainData.isEmpty()) {
                System.err.println("ERROR: No data found in main file: " + mainDataFile);
                return;
            }
            
            if (vendorData.isEmpty()) {
                System.err.println("ERROR: No data found in vendor file: " + vendorDataFile);
                return;
            }
            
            // Analyze file structures
            analyzeFileStructure(mainDataFile, mainData, "MAIN DATA (Authoritative Source)");
            analyzeFileStructure(vendorDataFile, vendorData, "VENDOR DATA");
            
            // Perform  inner join
            List<Map<String, Object>> innerJoinResult = performInnerJoin(mainData, vendorData, mainDataFile);
            
            if (innerJoinResult.isEmpty()) {
                System.err.println("WARNING: No matching records found between main data and vendor data!");
                System.err.println("Please check if the files have compatible join keys.");
                return;
            }
            
            // Write result to new Excel file
            writeExcelFile(innerJoinResult, outputFile, mainDataFile);
            
            System.out.println("\n=== RESULTS ===");
            System.out.println("✅  inner join completed successfully!");
            System.out.println("📁 Result saved to: " + outputFile);
            System.out.println("📊 Total records in result: " + innerJoinResult.size());
            System.out.println("🔍 Result maintains main data column structure exactly");
            
            // Verify the output file
            System.out.println("\n=== Verification ===");
            List<Map<String, Object>> verifyData = readExcelFile(outputFile);
            if (!verifyData.isEmpty()) {
                System.out.println("✅ Output file verification successful");
                System.out.println("📋 Final column structure: " + verifyData.get(0).keySet());
            }
            
        } catch (Exception e) {
            System.err.println("❌ CRITICAL ERROR: " + e.getMessage());
            e.printStackTrace();
            System.err.println("\n🔧 Troubleshooting Tips:");
            System.err.println("1. Ensure Excel files are not open in another application");
            System.err.println("2. Check file permissions");
            System.err.println("3. Verify file formats are .xlsx");
            System.err.println("4. Ensure files contain proper header rows");
        }
    }

    /**
     * Dynamically detect main data file from possible options
     */
    private static String detectMainDataFile() {
        for (String fileName : POSSIBLE_MAIN_FILES) {
            File file = new File(fileName);
            if (file.exists() && file.canRead()) {
                System.out.println("🔍 Found main data file: " + fileName);
                return fileName;
            }
        }
        return null;
    }
    
    /**
     * Dynamically detect vendor data file from possible options
     */
    private static String detectVendorDataFile() {
        for (String fileName : POSSIBLE_VENDOR_FILES) {
            File file = new File(fileName);
            if (file.exists() && file.canRead()) {
                System.out.println("🔍 Found vendor data file: " + fileName);
                return fileName;
            }
        }
        return null;
    }
    
    /**
     * Analyze and display file structure with enhanced details
     */
    private static void analyzeFileStructure(String filePath, List<Map<String, Object>> data, String fileType) {
        System.out.println("\n=== " + fileType + " ===");
        System.out.println("📄 File: " + filePath);
        System.out.println("📊 Total records: " + data.size());
        
        if (!data.isEmpty()) {
            Set<String> columns = data.get(0).keySet();
            System.out.println("📋 Columns (" + columns.size() + "): " + columns);
            
            // Show sample data
            System.out.println("📝 Sample record:");
            Map<String, Object> sample = data.get(0);
            for (Map.Entry<String, Object> entry : sample.entrySet()) {
                String value = entry.getValue() != null ? entry.getValue().toString() : "null";
                if (value.length() > 30) {
                    value = value.substring(0, 30) + "...";
                }
                System.out.println("   " + entry.getKey() + ": " + value);
            }
        }
    }
    
    /**
     * Read Excel file and return data as List of Maps - Enhanced version
     */
    public static List<Map<String, Object>> readExcelFile(String filePath) throws IOException {
        List<Map<String, Object>> data = new ArrayList<>();
        
        FileInputStream fis = null;
        Workbook workbook = null;
        
        try {
            fis = new FileInputStream(filePath);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Read first sheet
            
            if (sheet.getLastRowNum() < 0) {
                System.out.println("WARNING: No data found in " + filePath);
                return data;
            }
            
            // Find header row (check more rows for real data table)
            List<String> headers = new ArrayList<>();
            int headerRowIndex = -1;
            
            for (int rowIdx = 0; rowIdx <= Math.min(10, sheet.getLastRowNum()); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row != null) {
                    List<String> possibleHeaders = new ArrayList<>();
                    boolean hasValidHeaders = false;
                    int nonEmptyCount = 0;
                    
                    for (int cellIdx = 0; cellIdx < row.getLastCellNum(); cellIdx++) {
                        Cell cell = row.getCell(cellIdx);
                        String cellValue = getCellValueAsString(cell);
                        possibleHeaders.add(cellValue);
                        
                        if (!cellValue.trim().isEmpty()) {
                            nonEmptyCount++;
                            // Check if this looks like a header (non-empty, not just numbers)
                            if (!isNumeric(cellValue)) {
                                hasValidHeaders = true;
                            }
                        }
                    }
                    
                    // Look for rows with multiple columns and valid headers
                    // For small files, be less strict about column count
                    if (hasValidHeaders && nonEmptyCount >= 2 && !possibleHeaders.isEmpty()) {
                        headers = possibleHeaders;
                        headerRowIndex = rowIdx;
                        System.out.println("Found headers in row " + rowIdx + ": " + headers);
                        System.out.println("Non-empty columns: " + nonEmptyCount);
                        break;
                    }
                }
            }
            
            // If no headers found, check if first row has simple headers
            if (headers.isEmpty() && sheet.getLastRowNum() >= 0) {
                Row firstRow = sheet.getRow(0);
                if (firstRow != null) {
                    List<String> simpleHeaders = new ArrayList<>();
                    boolean hasSimpleHeaders = true;
                    
                    for (int i = 0; i < firstRow.getLastCellNum(); i++) {
                        Cell cell = firstRow.getCell(i);
                        String cellValue = getCellValueAsString(cell);
                        simpleHeaders.add(cellValue);
                        
                        // If any cell is empty or just numbers, it's probably not headers
                        if (cellValue.trim().isEmpty() || isNumeric(cellValue)) {
                            hasSimpleHeaders = false;
                        }
                    }
                    
                    if (hasSimpleHeaders && !simpleHeaders.isEmpty()) {
                        headers = simpleHeaders;
                        headerRowIndex = 0;
                        System.out.println("Found simple headers in row 0: " + headers);
                    } else {
                        // Create generic column names
                        for (int i = 0; i < firstRow.getLastCellNum(); i++) {
                            headers.add("Column_" + (i + 1));
                        }
                        headerRowIndex = -1; // Start reading from row 0
                        System.out.println("No headers found, using generic names: " + headers);
                    }
                }
            }
            
            // Read data rows
            int startRow = headerRowIndex + 1;
            if (headerRowIndex == -1) startRow = 0; // If no header found, start from first row
            
            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    // Use LinkedHashMap to preserve column order
                    Map<String, Object> rowData = new LinkedHashMap<>();
                    boolean hasData = false;
                    
                    for (int j = 0; j < headers.size() && j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        String header = headers.get(j);
                        Object value = getCellValue(cell);
                        
                        if (value != null && !value.toString().trim().isEmpty()) {
                            hasData = true;
                        }
                        
                        rowData.put(header, value);
                    }
                    
                    // Only add row if it has some data
                    if (hasData) {
                        data.add(rowData);
                    }
                }
            }
            
        } catch (Exception e) {
            System.err.println("ERROR reading " + filePath + ": " + e.getMessage());
            throw e; // Re-throw to be caught by main
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    System.err.println("Error closing file " + filePath + ": " + e.getMessage());
                }
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    System.err.println("Error closing workbook for " + filePath + ": " + e.getMessage());
                }
            }
        }
        
        return data;
    }
    
    /**
     * Check if a string represents a numeric value
     */
    private static boolean isNumeric(String str) {
        try {
            Double.parseDouble(str.trim());
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    
    /**
     * Get the original column order from Excel file with enhanced error handling
     */
    public static List<String> getOriginalColumnOrder(String filePath) throws IOException {
        List<String> columnOrder = new ArrayList<>();
        
        FileInputStream fis = null;
        Workbook workbook = null;
        
        try {
            fis = new FileInputStream(filePath);
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            
            if (sheet.getLastRowNum() >= 0) {
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                        Cell cell = headerRow.getCell(i);
                        String cellValue = getCellValueAsString(cell);
                        if (!cellValue.trim().isEmpty()) {
                            columnOrder.add(cellValue);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("ERROR reading column order from " + filePath + ": " + e.getMessage());
            throw e;
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    System.err.println("Error closing workbook: " + e.getMessage());
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    System.err.println("Error closing file: " + e.getMessage());
                }
            }
        }
        
        return columnOrder;
    }

    /**
     * Intelligently detect the best join key between datasets
     */
    private static String[] detectBestJoinKey(Set<String> mainColumns, Set<String> vendorColumns) {
        System.out.println("🔍 Intelligently detecting best join key...");
        
        // First, try preferred join column names in order
        for (String preferredCol : PREFERRED_JOIN_COLUMNS) {
            if (mainColumns.contains(preferredCol) && vendorColumns.contains(preferredCol)) {
                System.out.println("✅ Found perfect match: " + preferredCol);
                return new String[]{preferredCol, preferredCol};
            }
        }
        
        // Second, try case-insensitive matching for preferred columns
        for (String preferredCol : PREFERRED_JOIN_COLUMNS) {
            String mainMatch = findCaseInsensitiveMatch(preferredCol, mainColumns);
            String vendorMatch = findCaseInsensitiveMatch(preferredCol, vendorColumns);
            if (mainMatch != null && vendorMatch != null) {
                System.out.println("✅ Found case-insensitive match: " + mainMatch + " <-> " + vendorMatch);
                return new String[]{mainMatch, vendorMatch};
            }
        }
        
        // Third, try any column containing 'id'
        for (String mainCol : mainColumns) {
            for (String vendorCol : vendorColumns) {
                if (mainCol.toLowerCase().contains("id") && vendorCol.toLowerCase().contains("id")) {
                    System.out.println("✅ Found ID-containing columns: " + mainCol + " <-> " + vendorCol);
                    return new String[]{mainCol, vendorCol};
                }
            }
        }
        
        // Fourth, try name-based matching
        for (String mainCol : mainColumns) {
            for (String vendorCol : vendorColumns) {
                if (mainCol.toLowerCase().contains("name") && vendorCol.toLowerCase().contains("name")) {
                    System.out.println("✅ Found name-containing columns: " + mainCol + " <-> " + vendorCol);
                    return new String[]{mainCol, vendorCol};
                }
            }
        }
        
        // Last resort: use first common column
        Set<String> commonColumns = new HashSet<>(mainColumns);
        commonColumns.retainAll(vendorColumns);
        if (!commonColumns.isEmpty()) {
            String commonCol = commonColumns.iterator().next();
            System.out.println("⚠️ Using first common column: " + commonCol);
            return new String[]{commonCol, commonCol};
        }
        
        System.err.println("❌ No suitable join key found!");
        return new String[]{null, null};
    }
    
    /**
     * Find case-insensitive match for a column name
     */
    private static String findCaseInsensitiveMatch(String target, Set<String> columns) {
        for (String col : columns) {
            if (col.equalsIgnoreCase(target)) {
                return col;
            }
        }
        return null;
    }
    
    /**
     * Validate data integrity before performing join
     */
    private static boolean validateDataIntegrity(List<Map<String, Object>> mainData, List<Map<String, Object>> vendorData) {
        System.out.println("🔍 Validating data integrity...");
        
        if (mainData.isEmpty()) {
            System.err.println("❌ Main data is empty");
            return false;
        }
        
        if (vendorData.isEmpty()) {
            System.err.println("❌ Vendor data is empty");
            return false;
        }
        
        // Check if all main data records have the same column structure
        Set<String> expectedColumns = mainData.get(0).keySet();
        for (int i = 1; i < mainData.size(); i++) {
            if (!mainData.get(i).keySet().equals(expectedColumns)) {
                System.err.println("⚠️ Inconsistent column structure in main data at record " + (i + 1));
            }
        }
        
        // Check if all vendor data records have the same column structure
        Set<String> vendorColumns = vendorData.get(0).keySet();
        for (int i = 1; i < vendorData.size(); i++) {
            if (!vendorData.get(i).keySet().equals(vendorColumns)) {
                System.err.println("⚠️ Inconsistent column structure in vendor data at record " + (i + 1));
            }
        }
        
        System.out.println("✅ Data integrity validation completed");
        return true;
    }

    /**
     * Enhanced  inner join with intelligent key detection
     */
    public static List<Map<String, Object>> performInnerJoin(
            List<Map<String, Object>> mainData, 
            List<Map<String, Object>> vendorData,
            String mainDataFilePath) {
        
        List<Map<String, Object>> result = new ArrayList<>();
        
        System.out.println("\n🔄 Starting  Inner Join Process...");
        
        // Validate data integrity first
        if (!validateDataIntegrity(mainData, vendorData)) {
            System.err.println("❌ Data integrity validation failed");
            return result;
        }
        
        // Get column information
        Set<String> mainDataColumns = new HashSet<>(mainData.get(0).keySet());
        Set<String> vendorDataColumns = new HashSet<>(vendorData.get(0).keySet());
        
        System.out.println("📋 Main Data columns (" + mainDataColumns.size() + "): " + mainDataColumns);
        System.out.println("📋 Vendor Data columns (" + vendorDataColumns.size() + "): " + vendorDataColumns);
        
        // Find intersection and extra columns
        Set<String> commonColumns = new HashSet<>(mainDataColumns);
        commonColumns.retainAll(vendorDataColumns);
        Set<String> extraVendorColumns = new HashSet<>(vendorDataColumns);
        extraVendorColumns.removeAll(mainDataColumns);
        
        System.out.println("🔗 Common columns (" + commonColumns.size() + "): " + commonColumns);
        System.out.println("➕ Extra vendor columns (" + extraVendorColumns.size() + "): " + extraVendorColumns);
        
        // Get the exact column order from main data Excel file
        List<String> mainDataColumnOrder = new ArrayList<>();
        try {
            mainDataColumnOrder = getOriginalColumnOrder(mainDataFilePath);
            System.out.println("📑 Original main data column order: " + mainDataColumnOrder);
        } catch (IOException e) {
            System.err.println("⚠️ Warning: Could not read original column order, using runtime order");
            mainDataColumnOrder.addAll(mainData.get(0).keySet());
        }
        
        // Intelligently detect best join key
        String[] joinKeys = detectBestJoinKey(mainDataColumns, vendorDataColumns);
        String mainJoinKey = joinKeys[0];
        String vendorJoinKey = joinKeys[1];
        
        if (mainJoinKey == null || vendorJoinKey == null) {
            System.err.println("❌ No suitable join key found between datasets!");
            System.err.println("💡 Suggestion: Ensure both files have a common identifier column (like 'id', 'ID', etc.)");
            return result;
        }
        
        System.out.println("🔑 Using join keys: Main[" + mainJoinKey + "] ↔ Vendor[" + vendorJoinKey + "]");
        
        // Perform the inner join
        int matchCount = 0;
        int nullKeyCount = 0;
        Set<Object> processedKeys = new HashSet<>();
        
        for (Map<String, Object> mainRecord : mainData) {
            Object joinValue = mainRecord.get(mainJoinKey);
            
            if (joinValue == null || joinValue.toString().trim().isEmpty()) {
                nullKeyCount++;
                continue;
            }
            
            // Look for matching records in vendor data
            boolean matchFound = false;
            for (Map<String, Object> vendorRecord : vendorData) {
                Object vendorValue = vendorRecord.get(vendorJoinKey);
                
                if (vendorValue != null && isMatchingValue(joinValue, vendorValue)) {
                    // Create joined record
                    Map<String, Object> joinedRecord = new LinkedHashMap<>();
                    
                    // Add all main data columns in their exact original order
                    for (String column : mainDataColumnOrder) {
                        Object value = mainRecord.get(column);
                        joinedRecord.put(column, value);
                    }
                    
                    // Add any extra columns from vendor data that don't already exist
                    for (String extraColumn : extraVendorColumns) {
                        Object value = vendorRecord.get(extraColumn);
                        joinedRecord.put(extraColumn, value);
                    }
                    
                    result.add(joinedRecord);
                    matchFound = true;
                    matchCount++;
                    processedKeys.add(joinValue);
                    break; // Found match, move to next main record
                }
            }
        }
        
        // Report join statistics
        System.out.println("\n📊 Join Statistics:");
        System.out.println("✅ Successful matches: " + matchCount);
        System.out.println("📝 Total main records: " + mainData.size());
        System.out.println("⚠️ Records with null join keys: " + nullKeyCount);
        System.out.println("📈 Match rate: " + String.format("%.1f%%", (double) matchCount / mainData.size() * 100));
        
        if (matchCount == 0) {
            System.err.println("❌ No matches found! Please check:");
            System.err.println("   • Join key values in both files");
            System.err.println("   • Data formatting (numbers vs text)");
            System.err.println("   • Case sensitivity");
        }
        
        return result;
    }
    
    /**
     * Check if two values match (handles different data types)
     */
    private static boolean isMatchingValue(Object value1, Object value2) {
        if (value1 == null || value2 == null) {
            return false;
        }
        
        // Direct equality
        if (Objects.equals(value1, value2)) {
            return true;
        }
        
        // String comparison (case-insensitive, trimmed)
        String str1 = value1.toString().trim();
        String str2 = value2.toString().trim();
        
        if (str1.equalsIgnoreCase(str2)) {
            return true;
        }
        
        // Numeric comparison
        try {
            double num1 = Double.parseDouble(str1);
            double num2 = Double.parseDouble(str2);
            return Math.abs(num1 - num2) < 0.0001; // Handle floating point precision
        } catch (NumberFormatException e) {
            // Not numeric, fall through
        }
        
        return false;
    }
    
    /**
     * Write data to Excel file
     */
    public static void writeExcelFile(List<Map<String, Object>> data, String filePath, String mainDataFilePath) throws IOException {
        if (data.isEmpty()) {
            System.out.println("No data to write!");
            return;
        }
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(" Inner Join Result");
        
        // Get all column names from the main data file
        List<String> allColumns = new ArrayList<>();
        try {
            allColumns = getOriginalColumnOrder(mainDataFilePath);
        } catch (IOException e) {
            System.err.println("Warning: Could not read original column order for output, using data order");
            if (!data.isEmpty()) {
                allColumns.addAll(data.get(0).keySet());
            }
        }
        
        // Create header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < allColumns.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(allColumns.get(i));
        }
        
        // Create data rows
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            Map<String, Object> record = data.get(i);
            
            for (int j = 0; j < allColumns.size(); j++) {
                Cell cell = row.createCell(j);
                Object value = record.get(allColumns.get(j));
                
                if (value != null) {
                    if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else {
                        cell.setCellValue(value.toString());
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        }
        
        // Auto-size columns
        for (int i = 0; i < allColumns.size(); i++) {
            sheet.autoSizeColumn(i);
        }
        
        // Write to file
        FileOutputStream fos = new FileOutputStream(filePath);
        workbook.write(fos);
        workbook.close();
        fos.close();
    }
    
    /**
     * Get cell value as Object
     */
    public static Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }
    
    /**
     * Get cell value as String
     */
    public static String getCellValueAsString(Cell cell) {
        Object value = getCellValue(cell);
        return value != null ? value.toString() : "";
    }
    
    /**
     * Print data structure for debugging
     */
    public static void printDataStructure(List<Map<String, Object>> data) {
        if (data.isEmpty()) {
            System.out.println("No data found!");
            return;
        }
        
        Map<String, Object> firstRecord = data.get(0);
        System.out.println("Columns: " + firstRecord.keySet());
        System.out.println("Total records: " + data.size());
        
        // Show first record as sample
        if (!data.isEmpty()) {
            System.out.println("Sample record:");
            Map<String, Object> sample = data.get(0);
            for (Map.Entry<String, Object> entry : sample.entrySet()) {
                System.out.println("  " + entry.getKey() + ": " + entry.getValue());
            }
        }
    }
} 
