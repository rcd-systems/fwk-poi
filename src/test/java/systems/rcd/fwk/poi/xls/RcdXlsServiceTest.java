package systems.rcd.fwk.poi.xls;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.junit.Before;
import org.junit.Test;

import systems.rcd.fwk.poi.xls.data.RcdXlsWorkbook;

public class RcdXlsServiceTest {

    private static final String FILE_DIRECTORY = "src/test/resources/systems/rcd/fwk/poi/Xls";
    private static final Path INPUT_PATH = Paths.get(FILE_DIRECTORY, "input.xls");

    @Before
    public void before() {
        RcdXlsService.setDefaultService();
    }

    @Test
    public void test() throws Exception {
        final RcdXlsWorkbook xlsWorkbook = RcdXlsService.read(INPUT_PATH);
    }
}
