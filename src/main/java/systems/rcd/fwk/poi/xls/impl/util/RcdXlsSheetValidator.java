package systems.rcd.fwk.poi.xls.impl.util;

import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import systems.rcd.fwk.poi.xls.data.RcdXlsCell;
import systems.rcd.fwk.poi.xls.data.RcdXlsCellType;
import systems.rcd.fwk.poi.xls.data.RcdXlsRow;
import systems.rcd.fwk.poi.xls.data.RcdXlsSheet;

public class RcdXlsSheetValidator {

    private final Map<Integer, RcdXlsCellType> columnTypeMap = new LinkedHashMap<>();

    private final Set<Integer> mandatorinessSet = new LinkedHashSet<>();

    public RcdXlsSheetValidator setColumnType(final int columnIndex, final RcdXlsCellType columnType) {
        columnTypeMap.put(columnIndex, columnType);
        return this;
    }

    public RcdXlsSheetValidator setColumnMandatoriness(final int columnIndex) {
        mandatorinessSet.add(columnIndex);
        return this;
    }

    public List<String> validate(final String sheetName, final RcdXlsSheet sheet) {
        final List<String> errors = new LinkedList<String>();

        for (int i = 0; i < sheet.size(); i++) {
            final RcdXlsRow row = sheet.get(i);

            if (row == null) {
                errors.add("Sheet '" + sheetName + "': Empty row (#" + (i + 1) + "')");
                continue;
            }

            for (final Integer mandatoryColumnIndex : mandatorinessSet) {
                final RcdXlsCell rcdXlsCell = row.get(mandatoryColumnIndex);
                if (rcdXlsCell == null) {
                    errors.add("Sheet '" + sheetName + "': Empty cell (" + (i + 1) + ";" + (mandatoryColumnIndex + 1)
                            + "')");
                }
            }

            for (final Map.Entry<Integer, RcdXlsCellType> columnTypeEntry : columnTypeMap.entrySet()) {
                final int columnIndex = columnTypeEntry.getKey();
                final RcdXlsCell rcdXlsCell = row.get(columnIndex);
                if (rcdXlsCell != null && columnTypeEntry.getValue().equals(rcdXlsCell.getType())) {
                    errors.add("Sheet '" + sheetName + "': Incorrect type in the cell (" + (i + 1) + ";"
                            + (columnIndex + 1) + "').");
                }
            }
        }

        return errors;
    }
}
