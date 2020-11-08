package systems.rcd.fwk.poi.xls;

import java.nio.file.Files;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.time.LocalDate;
import java.time.ZoneId;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import systems.rcd.fwk.core.format.xls.RcdXlsService;
import systems.rcd.fwk.core.format.xls.data.RcdXlsRow;
import systems.rcd.fwk.core.format.xls.data.RcdXlsSheet;
import systems.rcd.fwk.core.format.xls.data.RcdXlsWorkbook;
import systems.rcd.fwk.poi.xls.impl.RcdPoiXlsService;

public class RcdXlsServiceTest
{

    private static final String FILE_DIRECTORY = "src/test/resources/systems/rcd/fwk/poi/xls";

    private static final Path INPUT_PATH = Paths.get( FILE_DIRECTORY, "input.xls" );

    @Before
    public void before()
    {
        RcdPoiXlsService.init();
    }

    @Test
    public void test()
        throws Exception
    {
        final RcdXlsWorkbook xlsWorkbook = RcdXlsService.read( Files.newInputStream( INPUT_PATH, StandardOpenOption.READ ) );
        System.out.println( xlsWorkbook.toString() );

        Assert.assertEquals( 3, xlsWorkbook.size() );
        final RcdXlsSheet xlsSheet2 = xlsWorkbook.get( "Sheet2" );
        final RcdXlsSheet xlsSheet1 = xlsWorkbook.get( "Sheet1" );
        Assert.assertEquals( 1, xlsSheet2.size() );
        Assert.assertEquals( 5, xlsSheet1.size() );

        Assert.assertNull( xlsSheet2.get( 0 ) );

        RcdXlsRow xlsRow = xlsSheet1.get( 0 );
        Assert.assertEquals( 4, xlsRow.size() );
        Assert.assertEquals( "someText", xlsRow.get( 0 ).getStringValue() );
        Assert.assertEquals( LocalDate.of( 2016, 2, 1 ), xlsRow.get( 1 ).getInstantValue().atZone( ZoneId.systemDefault() ).toLocalDate() );
        Assert.assertEquals( 1.2d, xlsRow.get( 2 ).getNumericValue(), 0d );
        Assert.assertNull( xlsRow.get( 3 ) );

        xlsRow = xlsSheet1.get( 1 );
        Assert.assertEquals( 4, xlsRow.size() );
        Assert.assertEquals( "someAdditionalText", xlsRow.get( 0 ).getStringValue() );
        Assert.assertNull( xlsRow.get( 1 ) );
        Assert.assertEquals( 4.0d, xlsRow.get( 2 ).getNumericValue(), 0d );
        Assert.assertTrue( xlsRow.get( 3 ).getBooleanValue() );

        xlsRow = xlsSheet1.get( 2 );
        Assert.assertEquals( 3, xlsRow.size() );
        Assert.assertEquals( "someTextsomeAdditionalText", xlsRow.get( 0 ).getStringValue() );
        Assert.assertNull( xlsRow.get( 1 ) );
        Assert.assertEquals( 5.2d, xlsRow.get( 2 ).getNumericValue(), 0d );

        xlsRow = xlsSheet1.get( 3 );
        Assert.assertNull( xlsRow );

        xlsRow = xlsSheet1.get( 4 );
        Assert.assertEquals( 1, xlsRow.size() );
        Assert.assertEquals( "alone", xlsRow.get( 0 ).getStringValue() );
    }
}
