package tables;

public class Table2 extends BaseTable {
    public String name;
    public String offices;
    public String employees;

    @Override
    public void initializeField(String fieldName, String cellValue) {
        switch (fieldName) {
            case "name":
                name = cellValue;
                break;
            case "offices":
                offices = cellValue;
                break;
            case "employees":
                employees = cellValue;
                break;
        }
    }
}
