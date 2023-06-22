package tables;

public class Table1 extends BaseTable {
    public String companyUrl;
    public String reviewsUrl;
    public String vacanciesUrl;

    @Override
    public void initializeField(String fieldName, String cellValue) {
        switch (fieldName) {
            case "companyUrl":
                companyUrl = cellValue;
                break;
            case "reviewsUrl":
                reviewsUrl = cellValue;
                break;
            case "vacanciesUrl":
                vacanciesUrl = cellValue;
                break;
        }
    }

}
