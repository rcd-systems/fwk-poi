package systems.rcd.fwk.poi.xls.impl.data;

import java.time.Instant;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import systems.rcd.fwk.core.format.xls.data.RcdXlsCell;
import systems.rcd.fwk.core.format.xls.data.RcdXlsCellType;

public class RcdPoiXlsCell implements RcdXlsCell {

    private String stringValue;

    private Instant instantValue;

    private Double numericValue;

    private Boolean booleanValue;

    public RcdPoiXlsCell(final Cell cell) {
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
            stringValue = cell.getRichStringCellValue().getString();
            break;
        case Cell.CELL_TYPE_NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                instantValue = cell.getDateCellValue().toInstant();
            } else {
                numericValue = cell.getNumericCellValue();
            }
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            booleanValue = cell.getBooleanCellValue();
            break;
        case Cell.CELL_TYPE_FORMULA:

            try {
                booleanValue = cell.getBooleanCellValue();
                break;
            } catch (final Exception e) {
            }
            try {
                numericValue = cell.getNumericCellValue();
                break;
            } catch (final Exception e) {
            }
            stringValue = cell.getRichStringCellValue().toString();
            break;
        }
    }

    @Override
    public RcdXlsCellType getType() {
        if (stringValue != null) {
            return RcdXlsCellType.STRING;
        }
        if (instantValue != null) {
            return RcdXlsCellType.INSTANT;
        }
        if (numericValue != null) {
            return RcdXlsCellType.NUMBER;
        }
        return RcdXlsCellType.BOOLEAN;
    }

    @Override
    public String getStringValue() {
        return stringValue;
    }

    @Override
    public Instant getInstantValue() {
        return instantValue;
    }

    @Override
    public double getNumericValue() {
        return numericValue;
    }

    @Override
    public boolean getBooleanValue() {
        return booleanValue;
    }

    @Override
    public String toString() {
        if (stringValue != null) {
            return stringValue;
        }
        if (instantValue != null) {
            return instantValue.toString();
        }
        if (numericValue != null) {
            return numericValue.toString();
        }
        return booleanValue.toString();
    }
}
