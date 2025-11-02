package org.trainman.service.util;

import org.trainman.model.Employee;
import org.trainman.service.util.ValidationUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.Locale;

public class ExcelParser {

    public static class ParseResult {
        private final List<Employee> employees = new ArrayList<>();
        private final List<String> errors = new ArrayList<>();

        public List<Employee> getEmployees() { return employees; }
        public List<String> getErrors() { return errors; }
    }

    public static ParseResult parse(InputStream in) {
        ParseResult result = new ParseResult();
        try (Workbook wb = WorkbookFactory.create(in)) {
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                try {
                    Employee e = new Employee();
                    // Use template's canonical column order to avoid mismatches
                    e.setEmpId(getString(row,0));
                    e.setName(getString(row,1));
                    e.setGender(getString(row,2));
                    e.setDoj(getDate(row,3));
                    e.setNsbtBatchNo(getString(row,4));
                    e.setStatus(getString(row,5));
                    e.setGrade(getString(row,6));
                    e.setBu(getString(row,7));
                    e.setMprNo(getString(row,8));
                    e.setIoName(getString(row,9));
                    e.setReleasedDate(getDate(row,10));
                    e.setResignationDate(getDate(row,11));

                    // Apply ranking based on status and grade
                    if ("Under Training".equalsIgnoreCase(e.getStatus())) {
                        e.setGrade(null); 
                        e.setRanking(null);
                    } else if ("Terminated".equalsIgnoreCase(e.getStatus())) {
                        e.setGrade("D"); 
                        e.setRanking("NI");
                    } else {
                        e.setRanking(RankingUtil.mapGradeToRanking(e.getGrade(), e.getStatus()));
                    }

                    result.getEmployees().add(e);
                } catch (Exception ex) {
                    result.getErrors().add("Row " + (i+1) + ": " + ex.getMessage());
                }
            }
        } catch (Exception e) {
            result.getErrors().add("Fatal parse error: " + e.getMessage());
        }
        return result;
    }

    private static String getString(Row r, int idx) {
        Cell c = r.getCell(idx);
        return formatCellAsString(c);
    }

    private static LocalDate getDate(Row r, int idx) {
        Cell c = r.getCell(idx);
        if (c == null) return null;
        // if numeric and date formatted, return directly
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
            return c.getLocalDateTimeCellValue().toLocalDate();
        }
        // else try formatted string
        String s = formatCellAsString(c);
        if (s == null || s.isEmpty()) return null;
        return parseLenientDate(s);
    }

    // Utility: robustly format any cell to a string (handles FORMULA, NUMERIC, STRING, BLANK)
    private static String formatCellAsString(Cell c) {
        if (c == null) return null;
        try {
            Workbook wb = c.getSheet().getWorkbook();
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            String text = formatter.formatCellValue(c, evaluator);
            if (text == null) return null;
            // normalize non-breaking/zero-width/unprintable spaces to regular space
            text = text.replace('\u00A0', ' ').replace('\u200B', ' ').replace('\uFEFF', ' ');
            // collapse whitespace and trim
            text = text.replaceAll("\\s+", " ").trim();
            return text.isEmpty() ? null : text;
        } catch (Exception ex) {
            // fallback
            try {
                String t = c.toString();
                if (t == null) return null;
                t = t.replace('\u00A0', ' ').replace('\u200B', ' ').replace('\uFEFF', ' ');
                t = t.replaceAll("\\s+", " ").trim();
                return t.isEmpty() ? null : t;
            } catch (Exception e) {
                return null;
            }
        }
    }

    private static final DateTimeFormatter DTF = DateTimeFormatter.ofPattern("dd-MM-yyyy");
    // Additional common formatters (month short names require Locale.ENGLISH)
    private static final DateTimeFormatter DTF_DD_MMM_YYYY = DateTimeFormatter.ofPattern("dd-MMM-yyyy", Locale.ENGLISH);
    private static final DateTimeFormatter DTF_DD_MMM_YY = DateTimeFormatter.ofPattern("dd-MMM-yy", Locale.ENGLISH);
    private static final DateTimeFormatter DTF_SLASH = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    private static final DateTimeFormatter ISO = DateTimeFormatter.ISO_LOCAL_DATE;

    private static final List<DateTimeFormatter> DATE_FORMATTERS = Arrays.asList(
            ISO,
            DTF,
            DTF_DD_MMM_YYYY,
            DTF_DD_MMM_YY,
            DTF_SLASH
    );

    private static LocalDate parseLenientDate(String s) {
        // First try parsing as ISO (yyyy-MM-dd) without modification
        for (DateTimeFormatter fmt : DATE_FORMATTERS) {
            try {
                return LocalDate.parse(s, fmt);
            } catch (DateTimeParseException ignored) {
            }
        }
        // Some Excel exports use dots or spaces (e.g., "02 Sep 2025" or "02.Sep.2025")
        String cleaned = s.replace('.', '-').replaceAll("\\s+", "-").replaceAll("[^0-9A-Za-z-]", "-");
        for (DateTimeFormatter fmt : DATE_FORMATTERS) {
            try {
                return LocalDate.parse(cleaned, fmt);
            } catch (DateTimeParseException ignored) {
            }
        }
        // If still not parsed, throw with a helpful message used by caller
        throw new DateTimeParseException("Text '" + s + "' could not be parsed as a supported date format", s, 0);
    }

    private static final Set<String> ALLOWED_STATUS = new HashSet<>(Arrays.asList(
            "Allocated", "Under Training", "Resigned", "Terminated", 
            "Temp Allocation", "Waiting for Allocation"
    ));

    public static ParseResult parseXlsx(InputStream in) throws Exception {
        ParseResult result = new ParseResult();
        try (Workbook wb = new XSSFWorkbook(in)) {
            Sheet sheet = wb.getSheetAt(0);
            if (sheet == null) {
                result.errors.add("Workbook has no sheets");
                return result;
            }

            // Try to find a valid header row within the first 3 rows
            Map<String, Integer> colIndex = Collections.emptyMap();
            int headerRowNum = sheet.getFirstRowNum();
            for (int hr = sheet.getFirstRowNum(); hr <= sheet.getFirstRowNum() + 2 && hr <= sheet.getLastRowNum(); hr++) {
                Row headerRow = sheet.getRow(hr);
                if (headerRow == null) continue;
                Map<String, Integer> candidate = safeMapHeader(headerRow);
                if (!candidate.isEmpty()) {
                    boolean ok = true;
                    for (String req : new String[]{"empid","name","status","doj"}) {
                        if (!candidate.containsKey(req)) { ok = false; break; }
                    }
                    if (ok) { colIndex = candidate; headerRowNum = hr; break; }
                }
            }

            // If no valid header found, fall back to positional mapping using the template order
            boolean usedPositionalFallback = false;
            if (colIndex.isEmpty()) {
                usedPositionalFallback = true;
                Map<String, Integer> positional = new HashMap<>();
                // canonical template order
                positional.put("empid", 0);
                positional.put("name", 1);
                positional.put("gender", 2);
                positional.put("doj", 3);
                positional.put("nsbt batchno", 4);
                positional.put("status", 5);
                positional.put("grade", 6);
                positional.put("bu", 7);
                positional.put("mpr no", 8);
                positional.put("io name", 9);
                positional.put("released date", 10);
                positional.put("resiznation date", 11);
                colIndex = positional;
            }

            int rowNum = 0;
            for (int r = (headerRowNum + 1); r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                rowNum = row.getRowNum() + 1; // human readable

                String empid = getCellString(row, colIndex.get("empid"));
                if (empid == null || empid.trim().isEmpty()) {
                    result.errors.add("Row " + rowNum + ": empid is empty");
                    continue;
                }
                Employee e = new Employee();
                e.setEmpId(empid.trim());
                e.setName(getCellString(row, colIndex.get("name")));
                e.setGender(getCellString(row, colIndex.get("gender")));
                e.setNsbtBatchNo(getCellString(row, colIndex.get("nsbt batchno")));
                String status = getCellString(row, colIndex.get("status"));
                if (status == null) status = "";
                status = status.trim();
                if (!ALLOWED_STATUS.contains(status)) {
                    result.errors.add("Row " + rowNum + ": invalid status '" + status + "'");
                    continue;
                }
                e.setStatus(status);

                // read fields required for validation BEFORE date checks
                e.setGrade(getCellString(row, colIndex.get("grade")));
                e.setBu(getCellString(row, colIndex.get("bu")));
                e.setMprNo(getCellString(row, colIndex.get("mpr no")));
                e.setIoName(getCellString(row, colIndex.get("io name")));

                try {
                    LocalDate doj = parseDateCell(row, colIndex.get("doj"));
                    LocalDate resignDate = parseDateCell(row, colIndex.get("resiznation date"));
                    LocalDate releasedDate = parseDateCell(row, colIndex.get("released date"));

                    // Validate dates
                    if (doj == null) {
                        result.errors.add("Row " + rowNum + ": DOJ is required");
                        continue;
                    }
                    if (doj.isAfter(LocalDate.now())) {
                        result.errors.add("Row " + rowNum + ": DOJ cannot be in the future");
                        continue;
                    }
                    if (resignDate != null && resignDate.isBefore(doj)) {
                        result.errors.add("Row " + rowNum + ": Resignation date cannot be before DOJ");
                        continue;
                    }
                    if (releasedDate != null && releasedDate.isBefore(doj)) {
                        result.errors.add("Row " + rowNum + ": Release date cannot be before DOJ");
                        continue;
                    }

                    e.setDoj(doj);
                    e.setResignationDate(resignDate);
                    e.setReleasedDate(releasedDate);

                    // Validate the employee object using ValidationUtil
                    List<String> validationErrors = ValidationUtil.validate(e);
                    if (!validationErrors.isEmpty()) {
                        String diag = String.format(" [parsed: EMPID=%s, IO_NAME=%s, BU=%s, MPR_NO=%s]",
                                e.getEmpId(), e.getIoName(), e.getBu(), e.getMprNo());
                        result.errors.add("Row " + rowNum + ": " + String.join(", ", validationErrors) + diag);
                        continue;
                    }

                } catch (DateTimeParseException dtpe) {
                    result.errors.add("Row " + rowNum + ": invalid date format - " + dtpe.getMessage());
                    continue;
                } catch (Exception ex) {
                    result.errors.add("Row " + rowNum + ": " + ex.getMessage());
                    continue;
                }

                result.employees.add(e);
            }

            if (usedPositionalFallback) {
                result.errors.add(0, "Warning: header row not detected; positional fallback used (assuming template column order). If columns are shifted, please use the provided template and do not merge header cells.");
            }

        }
        return result;
    }

    // Safe header mapping that returns normalized map without throwing
    private static Map<String, Integer> safeMapHeader(Row header) {
        Map<String, Integer> idx = new HashMap<>();
        if (header == null) return idx;
        for (Cell c : header) {
            String v = formatCellAsString(c);
            if
