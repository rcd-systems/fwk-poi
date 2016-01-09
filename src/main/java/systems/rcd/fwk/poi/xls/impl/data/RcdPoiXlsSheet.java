package systems.rcd.fwk.poi.xls.impl.data;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import systems.rcd.fwk.poi.xls.data.RcdXlsRow;
import systems.rcd.fwk.poi.xls.data.RcdXlsSheet;

public class RcdPoiXlsSheet extends ArrayList<RcdXlsRow> implements RcdXlsSheet {
    public RcdPoiXlsSheet(final Sheet sheet) {
        super(sheet.getLastRowNum());
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            final Row row = sheet.getRow(i);
            if (row == null) {
                add((RcdXlsRow) null);
            } else {
                add(new RcdPoiXlsRow(row));
            }
        }
    }

}
