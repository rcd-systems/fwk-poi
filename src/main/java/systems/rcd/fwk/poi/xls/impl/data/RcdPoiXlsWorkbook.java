package systems.rcd.fwk.poi.xls.impl.data;

import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import systems.rcd.fwk.poi.xls.data.RcdXlsSheet;
import systems.rcd.fwk.poi.xls.data.RcdXlsWorkbook;

public class RcdPoiXlsWorkbook extends LinkedHashMap<String, RcdXlsSheet> implements RcdXlsWorkbook {
    public RcdPoiXlsWorkbook(final Workbook workbook) {
        super((int) (workbook.getNumberOfSheets() / 0.75 + 1));
        for (final Sheet sheet : workbook) {
            put(sheet.getSheetName(), new RcdPoiXlsSheet(sheet));
        }
    }

    @Override
    public String toString() {
        final StringBuffer sb = new StringBuffer();
        for (final Map.Entry<String, RcdXlsSheet> entry : entrySet()) {
            sb.append("Sheet '" + entry.getKey() + "'").append(System.lineSeparator())
                    .append(entry.getValue().toString()).append(System.lineSeparator());
        }
        return sb.toString();
    }
}
