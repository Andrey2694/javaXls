import helpers.TableHelperObj;
import tables.Table2;

import java.util.List;

public class Main {
    public static void main(String[] args) {
        TableHelperObj tablesHelper = new TableHelperObj();
//
//        List<Table2> data = tablesHelper.convertTableToObjList("src/test/resources/book.csv", Table2.class, TableReferral.IS_TABLE_VERTICAL, 1);
//        List<Table2> data = tablesHelper.convertTableToObjList("src/test/resources/book.csv", Table2.class, 1);
        List<Table2> data = tablesHelper.convertTableToObjList("src/test/resources/result.xlsx", Table2.class);
        tablesHelper.writeXls(data);
        tablesHelper.writeCsv(data);

    }
}
