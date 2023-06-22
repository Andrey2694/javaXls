package enums;

public enum TableReferral {
    IS_TABLE_HORIZONTAL(true),
    IS_TABLE_VERTICAL(true);

    TableReferral(boolean tableReferral) {
        this.tableReferral = tableReferral;
    }

    private final boolean tableReferral;

    public boolean getTableReferral() {
        return tableReferral;
    }
}
