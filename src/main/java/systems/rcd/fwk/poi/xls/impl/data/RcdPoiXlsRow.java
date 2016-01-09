package systems.rcd.fwk.poi.xls.impl.data;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import systems.rcd.fwk.poi.xls.data.RcdXlsCell;
import systems.rcd.fwk.poi.xls.data.RcdXlsRow;

public class RcdPoiXlsRow extends ArrayList<RcdXlsCell> implements RcdXlsRow {

    public RcdPoiXlsRow(final Row row) {
        super(row.getLastCellNum());
        for (int i = 0; i < row.getLastCellNum(); i++) {
            final Cell cell = row.getCell(i, Row.RETURN_BLANK_AS_NULL);
            if (cell == null) {
                add(null);
            } else {
                add(new RcdPoiXlsCell(cell));
            }
        }
    }

}
