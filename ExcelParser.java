package org.trainman.service.util;

import org.trainman.model.Employee;
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

    // Simple positional parser (keeps legacy behaviour) - expects template column order
    public static ParseResult parse(InputStream in) {
        ParseResult result = new ParseResult();
        try (Workbook wb = WorkbookFactory.create(in)) {
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                try {
                    Employee e = new Employee();
                    // Template canonical order (positional)
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

    // --- Helpers for cell reading ---
    private static String getString(Row r, int idx) {
        Cell c = r.getCell(idx);
        return formatCellAsString(c);
    }

    private static LocalDate getDate(Row r, int idx) {
        Cell c = r.getCell(idx);
        if (c == null) return null;
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
            return c.getLocalDateTimeCellValue().toLocalDate();
        }
        String s = formatCellAsString(c);
        if (s == null || s.isEmpty()) return null;
        return parseLenientDate(s);
    }

    // Robust formatter (handles formulas, numeric, date formatting and trims/normalizes whitespace)
    private static String formatCellAsString(Cell c) {
        if (c == null) return null;
        try {
            Workbook wb = c.getSheet().getWorkbook();
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            String text = formatter.formatCellValue(c, evaluator);
            if (text == null) return null;
            // Normalize non-breaking/zero-width/unprintable spaces to regular space
            text = text.replace('\u00A0', ' ').replace('\u200B', ' ').replace('\uFEFF', ' ');
            text = text.replaceAll("\\s+", " ").trim();
            return text.isEmpty() ? null : text;
        } catch (Exception ex) {
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

    // Date parsers
    private static final DateTimeFormatter DTF = DateTimeFormatter.ofPattern("dd-MM-yyyy");
    private static final DateTimeFormatter DTF_DD_MMM_YYYY = DateTimeFormatter.ofPattern("dd-MMM-yyyy", Locale.ENGLISH);
    private static final DateTimeFormatter DTF_DD_MMM_YY = DateTimeFormatter.ofPattern("dd-MMM-yy", Locale.ENGLISH);
    private static final DateTimeFormatter DTF_SLASH = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    private static final DateTimeFormatter ISO = DateTimeFormatter.ISO_LOCAL_DATE;

    private static final List<DateTimeFormatter> DATE_FORMATTERS = Arrays.asList(
            ISO, DTF, DTF_DD_MMM_YYYY, DTF_DD_MMM_YY, DTF_SLASH
    );

    private static LocalDate parseLenientDate(String s) {
        if (s == null) throw new DateTimeParseException("null", "", 0);
        for (DateTimeFormatter fmt : DATE_FORMATTERS) {
            try {
                return LocalDate.parse(s, fmt);
            } catch (DateTimeParseException ignored) {}
        }
        String cleaned = s.replace('.', '-').replaceAll("\\s+", "-").replaceAll("[^0-9A-Za-z-]", "-");
        for (DateTimeFormatter fmt : DATE_FORMATTERS) {
            try {
                return LocalDate.parse(cleaned, fmt);
            } catch (DateTimeParseException ignored) {}
        }
        throw new DateTimeParseException("Text '" + s + "' could not be parsed as a supported date format", s, 0);
    }

    private static final Set<String> ALLOWED_STATUS = new HashSet<>(Arrays.asList(
            "Allocated", "Under Training", "Resigned", "Terminated", "Temp Allocation", "Waiting for Allocation"
    ));

    // Header-aware parser that maps columns by header name (robust to minor header variations)
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
                    for (String req : new String[]{"empid", "name", "status", "doj"}) {
                        if (!candidate.containsKey(req)) { ok = false; break; }
                    }
                    if (ok) { colIndex = candidate; headerRowNum = hr; break; }
                }
            }

            // Fallback to positional template order if header not detected
            boolean usedPositionalFallback = false;
            if (colIndex.isEmpty()) {
                usedPositionalFallback = true;
                Map<String, Integer> positional = new HashMap<>();
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

            // iterate rows after header
            for (int r = headerRowNum + 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                int rowNum = row.getRowNum() + 1;

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

                    // Validate
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

    // Return string value for cell index (handles numeric/date via formatCellAsString)
    private static String getCellString(Row row, Integer idx) {
        if (idx == null) return null;
        Cell c = row.getCell(idx);
        if (c == null) return null;
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
            LocalDate d = c.getLocalDateTimeCellValue().toLocalDate();
            return d.format(DTF);
        }
        String s = formatCellAsString(c);
        if (s == null) return null;
        // if numeric-looking like 12345.0 -> 12345
        try {
            double dval = Double.parseDouble(s);
            long l = (long) dval;
            if (Math.abs(dval - l) < 0.0001) return String.valueOf(l);
            return s;
        } catch (NumberFormatException nfe) {
            return s.isBlank() ? null : s;
        }
    }

    private static LocalDate parseDateCell(Row row, Integer idx) {
        if (idx == null) return null;
        Cell c = row.getCell(idx);
        if (c == null) return null;
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
            return c.getLocalDateTimeCellValue().toLocalDate();
        } else {
            String s = formatCellAsString(c);
            if (s == null || s.isEmpty()) return null;
            return parseLenientDate(s);
        }
    }

    // Safe header mapping - returns canonical keys mapped to column index
    private static Map<String, Integer> safeMapHeader(Row header) {
        Map<String, Integer> idx = new HashMap<>();
        if (header == null) return idx;
        for (Cell c : header) {
            String v = formatCellAsString(c);
            if (v == null) continue;
            idx.put(v.trim().toLowerCase(), c.getColumnIndex());
        }
        Map<String, Integer> normalized = new HashMap<>();
        for (Map.Entry<String, Integer> e : idx.entrySet()) {
            String k = e.getKey().toLowerCase().trim();
            k = k.replaceAll("[\\u00A0\\u200B\\uFEFF]", " ").trim();
            String kn = k.replaceAll("[^a-z0-9]", "");

            if (kn.contains("empid") || kn.contains("employeeid")) normalized.put("empid", e.getValue());
            if (kn.equals("name") || kn.contains("employeename")) normalized.put("name", e.getValue());
            if (kn.contains("gender") || kn.equals("sex")) normalized.put("gender", e.getValue());
            if (kn.contains("doj") || kn.contains("dateofjoin") || kn.contains("dateofjoining")) normalized.put("doj", e.getValue());
            if (kn.contains("resign") || kn.contains("resiz") || kn.contains("leavingdate") || kn.contains("resignationdate")) normalized.put("resiznation date", e.getValue());
            if (kn.contains("released") || kn.contains("releasedate") || kn.contains("releasedon")) normalized.put("released date", e.getValue());
            if ((kn.contains("nsbt") || kn.contains("batch")) && kn.contains("no")) normalized.put("nsbt batchno", e.getValue());
            if (kn.contains("status") || kn.contains("employeestatus")) normalized.put("status", e.getValue());
            if (kn.contains("grade") || kn.contains("employeegrade")) normalized.put("grade", e.getValue());
            if (kn.equals("bu") || kn.contains("businessunit")) normalized.put("bu", e.getValue());
            if (kn.contains("mpr") || kn.contains("projectno")) normalized.put("mpr no", e.getValue());
            if (kn.contains("io") || kn.contains("immediateofficer") || kn.contains("supervisor")) normalized.put("io name", e.getValue());
        }
        return normalized;
    }
}
