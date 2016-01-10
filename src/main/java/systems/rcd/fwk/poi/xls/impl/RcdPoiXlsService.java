package systems.rcd.fwk.poi.xls.impl;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import systems.rcd.fwk.core.ctx.RcdContext;
import systems.rcd.fwk.core.format.xls.RcdXlsService;
import systems.rcd.fwk.poi.xls.impl.data.RcdPoiXlsWorkbook;

public class RcdPoiXlsService implements RcdXlsService {

    public static void init() {
        RcdContext.setGlobalServiceSupplier(RcdXlsService.class, () -> new RcdPoiXlsService());
    }

    @Override
    public RcdPoiXlsWorkbook instRead(final Path path) throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        final InputStream inputStream = Files.newInputStream(path, StandardOpenOption.READ);
        final Workbook workbook = WorkbookFactory.create(inputStream);
        return new RcdPoiXlsWorkbook(workbook);
    }
}
