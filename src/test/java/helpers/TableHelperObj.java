package helpers;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import enums.TableReferral;
import org.apache.poi.xssf.usermodel.XSSFRow;
import tables.BaseTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class TableHelperObj {
    public <T extends BaseTable> List<T> convertTableToObjList(String path, Class<T> tClass, int skipCount) {
        return convertTableToObjList(path, tClass, TableReferral.IS_TABLE_HORIZONTAL, skipCount);
    }

    public <T extends BaseTable> List<T> convertTableToObjList(String path, Class<T> tClass) {
        return convertTableToObjList(path, tClass, TableReferral.IS_TABLE_HORIZONTAL, 0);
    }

    public <T extends BaseTable> List<T> convertTableToObjList(String path, Class<T> tClass, TableReferral referral, int skipCount) {
        List<T> data = new ArrayList<>();

        if (path.endsWith("csv")) {
            try (CSVReader csvReader = new CSVReader(new FileReader(path))) {
                // Read the first row as field names
                List<String> fieldNames = Arrays.asList(csvReader.readNext());

                String[] values;
                while ((values = csvReader.readNext()) != null) {
                    T instance = tClass.getDeclaredConstructor().newInstance();
                    for (int i = 0; i < values.length; i++) {
                        initializeFields(instance, fieldNames, i, values[i]);
                    }
                    data.add(instance);
                }
            } catch (IOException | ReflectiveOperationException e) {
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

                Row firstRow = sheet.getRow(0);
                List<String> fieldNames = extractFieldNames(firstRow);

                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    T instance = tClass.getDeclaredConstructor().newInstance();
                    for (Cell cell : row) {
                        initializeFields(instance, fieldNames, cell.getColumnIndex(), getCellValue(cell));
                    }
                    data.add(instance);
                }
            } catch (IOException | ReflectiveOperationException e) {
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

        if (TableReferral.IS_TABLE_HORIZONTAL.name().equals(referral.name())) {
            // delete rows from 0 to skipCount
            data.subList(0, skipCount).clear();
        } else {
            // set null for skipCount fields
            removeFieldsFromList(data, skipCount);
        }

        return data;
    }

    public <T extends BaseTable> void writeXls(List<T> data) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        int rowIndex = 0;

        // Get the field names from the first object in the list
        T firstObject = data.get(0);
        Field[] fields = firstObject.getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for (Field field : fields) {
            fieldNames.add(field.getName());
        }

        // Write the field names as the first row in the sheet
        Row headerRow = sheet.createRow(rowIndex++);
        for (int i = 0; i < fieldNames.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(fieldNames.get(i));
        }

        // Write data rows
        for (T obj : data) {
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (Field field : fields) {
                field.setAccessible(true);
                try {
                    Object value = field.get(obj);
                    Cell cell = row.createCell(cellIndex++);
                    if (value != null) {
                        cell.setCellValue(value.toString());
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }

        File outputFile = new File("src/test/resources/temp.xlsx");
        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
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


    public <T extends BaseTable> void writeCsv(List<T> data) {
        File file = new File("src/test/resources/temp.csv");

        try (FileWriter fileWriter = new FileWriter(file);
             CSVWriter writer = new CSVWriter(fileWriter)) {

            // Get the field names from the first object in the list
            T firstObject = data.get(0);
            Field[] fields = firstObject.getClass().getDeclaredFields();
            List<String> fieldNames = new ArrayList<>();
            for (Field field : fields) {
                fieldNames.add(field.getName());
            }

            // Write the field names as the first row
            writer.writeNext(fieldNames.toArray(new String[0]));

            // Write data rows
            for (T obj : data) {
                List<String> rowData = new ArrayList<>();
                for (Field field : fields) {
                    field.setAccessible(true);
                    try {
                        Object value = field.get(obj);
                        rowData.add(value != null ? value.toString() : "");
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
                writer.writeNext(rowData.toArray(new String[0]));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<String> extractFieldNames(Row row) {
        List<String> fieldNames = new ArrayList<>();
        for (Cell cell : row) {
            fieldNames.add(getCellValue(cell));
        }
        return fieldNames;
    }

    private void initializeFields(BaseTable instance, List<String> fieldNames, int cellIndex, String cellValue) {
        String fieldName = fieldNames.get(cellIndex);
        instance.initializeField(fieldName, cellValue);
    }

    private String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
            }
        }
        return cellValue;
    }

    private <T extends BaseTable> void removeFieldsFromList(List<T> dataList, int skipCount) {
        for (T obj : dataList) {
            Class<?> objClass = obj.getClass();
            Field[] fields = objClass.getDeclaredFields();
            try {
                for (int i = 0; i < skipCount; i++) {
                    try {
                        fields[i].setAccessible(true);
                        fields[i].set(obj, null);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            } catch (ArrayIndexOutOfBoundsException e) {
                e.printStackTrace();
            }
        }
    }
}
