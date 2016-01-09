package systems.rcd.fwk.poi.xls.impl.data;

import org.apache.poi.ss.usermodel.Cell;

import systems.rcd.fwk.poi.xls.data.RcdXlsCell;
import systems.rcd.fwk.poi.xls.data.RcdXlsCellType;

public class RcdPoiXlsCell implements RcdXlsCell {

    private RcdXlsCellType type;

    private String stringValue;

    private double numericValue;

    private boolean booleanValue;

    public RcdPoiXlsCell(final Cell cell) {
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
            type = RcdXlsCellType.STRING;
            stringValue = cell.getRichStringCellValue().getString();
            break;
        case Cell.CELL_TYPE_NUMERIC:
            type = RcdXlsCellType.NUMBER;
            numericValue = cell.getNumericCellValue();
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            type = RcdXlsCellType.BOOLEAN;
            booleanValue = cell.getBooleanCellValue();
            break;
        case Cell.CELL_TYPE_FORMULA:

            try {
                booleanValue = cell.getBooleanCellValue();
                type = RcdXlsCellType.BOOLEAN;
                break;
            } catch (final Exception e) {
            }
            try {
                numericValue = cell.getNumericCellValue();
                type = RcdXlsCellType.NUMBER;
                break;
            } catch (final Exception e) {
            }
            stringValue = cell.getRichStringCellValue().toString();
            type = RcdXlsCellType.STRING;
            break;
        }
    }

    @Override
    public RcdXlsCellType getType() {
        return type;
    }

    @Override
    public String getStringValue() {
        return stringValue;
    }

    @Override
    public double getNumericValue() {
        return numericValue;
    }

    @Override
    public boolean getBooleanValue() {
        return booleanValue;
    }
}
