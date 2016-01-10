package systems.rcd.fwk.poi.xls;

import java.nio.file.Path;

import systems.rcd.fwk.core.ctx.RcdContext;
import systems.rcd.fwk.core.ctx.RcdService;
import systems.rcd.fwk.poi.xls.data.RcdXlsWorkbook;
import systems.rcd.fwk.poi.xls.impl.RcdPoiXlsService;

public interface RcdXlsService extends RcdService {

    static RcdXlsWorkbook read(final Path path) throws Exception {
        return RcdContext.getService(RcdXlsService.class).instRead(path);
    }

    static void setDefaultService() {
        RcdContext.setGlobalServiceSupplier(RcdXlsService.class, () -> new RcdPoiXlsService());
    }

    RcdXlsWorkbook instRead(Path path) throws Exception;

}
