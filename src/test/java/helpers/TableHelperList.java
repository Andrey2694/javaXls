package helpers;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import enums.TableReferral;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class TableHelperList {
    public List<List<String>> convertTableToList(String path, int skipCount) {
        return convertTableToList(path, TableReferral.IS_TABLE_HORIZONTAL, skipCount);
    }

    public List<List<String>> convertTableToList(String path) {
        return convertTableToList(path, TableReferral.IS_TABLE_HORIZONTAL, 0);
    }

    public List<List<String>> convertTableToList(String path, TableReferral referral, int skipCount) {
        List<List<String>> data = new ArrayList<>();

        if (path.endsWith("csv")) {
            try (CSVReader csvReader = new CSVReader(new FileReader(path))) {
                String[] values;
                while ((values = csvReader.readNext()) != null) {
                    data.add(Arrays.asList(values));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            FileInputStream file = null;
            Workbook workbook = null;
            try {
                file = new FileInputStream(path);
                if (path.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(file);
                } else {
                    workbook = new XSSFWorkbook(file);
                }
                Sheet sheet = workbook.getSheetAt(0);

                int i = 0;
                for (Row row : sheet) {
                    data.add(i, new ArrayList<>());
                    for (Cell cell : row) {
                        data.get(i).add(getCellValue(cell));
                    }
                    i++;
                }
                if (TableReferral.IS_TABLE_HORIZONTAL.name().equals(referral.name())) {
//               delete rows from 0 to skipCount
                    data.subList(0, skipCount).clear();
                } else {
//               delete elements from each row
                    for (List<String> innerList : data) {
                        if (!innerList.isEmpty()) {
                            innerList.subList(0, skipCount).clear();
                        }
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                if (workbook != null) {
                    try {
                        workbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (file != null) {
                    try {
                        file.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }

        return data;
    }

    public void writeXls(List<List<String>> data) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        int count = 0;

        for (List<String> value : data) {
            XSSFRow row = (XSSFRow) sheet.createRow(count++);
            for (int a = 0; a < value.size(); a++) {
                row.createCell(a).setCellValue(value.get(a));
            }
        }

        File currDir = new File("src/test/resources/temp.xlsx");
        try {
            FileOutputStream outputStream = new FileOutputStream(currDir);
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void writeCsv(List<List<String>> data) {
        File file = new File("src/test/resources/temp.csv");

        try {
            FileWriter outfile = new FileWriter(file);
            CSVWriter writer = new CSVWriter(outfile);

            List<String[]> convertedList = new ArrayList<>();
            for (List<String> innerList : data) {
                String[] array = innerList.toArray(new String[0]);
                convertedList.add(array);
            }

            writer.writeAll(convertedList);
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //    maybe will be useful converting to Json
    public ObjectNode convertTableToJson(String path) {
        ObjectMapper mapper = new ObjectMapper();
        ObjectNode excelData = mapper.createObjectNode();
        FileInputStream file = null;
        Workbook workbook = null;
        try {
            file = new FileInputStream(path);

            if (path.endsWith(".xls") || path.endsWith(".xlsx")) {
                if (path.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(file);
                } else {
                    workbook = new XSSFWorkbook(file);
                }

                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();

                    List<String> headers = new ArrayList<>();
                    ArrayNode sheetData = mapper.createArrayNode();
                    for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                        Row row = sheet.getRow(j);
                        if (j == 0) {
                            for (int k = 0; k < row.getLastCellNum(); k++) {
                                headers.add(row.getCell(k).getStringCellValue());
                            }
                        } else {
                            ObjectNode rowData = mapper.createObjectNode();
                            for (int k = 0; k < headers.size(); k++) {
                                Cell cell = row.getCell(k);
                                String headerName = headers.get(k);
                                if (cell != null) {
                                    rowData.put(headerName, getCellValue(cell));
                                } else {
                                    rowData.put(headerName, "");
                                }
                            }
                            sheetData.add(rowData);
                        }
                    }
                    excelData.set(sheetName, sheetData);
                }

                return excelData;
            } else {
                throw new IllegalArgumentException("File format not supported.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (file != null) {
                try {
                    file.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }
        return null;
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue() + "";
                } else {
                    return cell.getNumericCellValue() + "";
                }
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case FORMULA:
                return cell.getCellFormula() + "";
            default:
                return "";
        }
    }
}

