package systems.rcd.fwk.poi.xls.data;

import java.time.Instant;

public interface RcdXlsCell {
    RcdXlsCellType getType();

    String getStringValue();

    Instant getInstantValue();

    double getNumericValue();

    boolean getBooleanValue();
}
