package systems.rcd.fwk.poi.xls.data;

public interface RcdXlsCell {
    RcdXlsCellType getType();

    String getStringValue();

    double getNumericValue();

    boolean getBooleanValue();
}
