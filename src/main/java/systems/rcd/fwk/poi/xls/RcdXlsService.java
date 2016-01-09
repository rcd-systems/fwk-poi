package systems.rcd.fwk.poi.xls;

import java.io.IOException;
import java.nio.file.Path;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import systems.rcd.fwk.poi.xls.data.RcdXlsWorkbook;

public interface RcdXlsService {

    RcdXlsWorkbook instRead(Path path) throws EncryptedDocumentException, InvalidFormatException, IOException;

}
