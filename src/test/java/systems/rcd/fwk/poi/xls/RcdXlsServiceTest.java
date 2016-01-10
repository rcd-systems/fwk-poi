package systems.rcd.fwk.poi.xls;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.ZoneId;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import systems.rcd.fwk.poi.xls.data.RcdXlsRow;
import systems.rcd.fwk.poi.xls.data.RcdXlsSheet;
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
        System.out.println(xlsWorkbook.toString());

        Assert.assertEquals(3, xlsWorkbook.size());
        final RcdXlsSheet xlsSheet2 = xlsWorkbook.get("Sheet2");
        final RcdXlsSheet xlsSheet1 = xlsWorkbook.get("Sheet1");
        Assert.assertEquals(1, xlsSheet2.size());
        Assert.assertTrue(xlsSheet1.size() >= 5);

        Assert.assertNull(xlsSheet2.get(0));

        final RcdXlsRow xlsRow1 = xlsSheet1.get(0);
        Assert.assertEquals("someText", xlsRow1.get(0).getStringValue());
        Assert.assertEquals(LocalDate.of(2016, 2, 1), xlsRow1.get(1).getInstantValue().atZone(ZoneId.systemDefault())
                .toLocalDate());
        Assert.assertEquals(1.2d, xlsRow1.get(2).getNumericValue(), 0d);
        Assert.assertNull(xlsRow1.get(3));
    }
}
