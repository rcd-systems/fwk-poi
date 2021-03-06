package systems.rcd.fwk.poi.xls.impl;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import systems.rcd.fwk.core.ctx.RcdContext;
import systems.rcd.fwk.core.exc.RcdException;
import systems.rcd.fwk.core.format.xls.RcdXlsService;
import systems.rcd.fwk.poi.xls.impl.data.RcdPoiXlsWorkbook;

public class RcdPoiXlsService
    implements RcdXlsService
{
    public static void init()
    {
        RcdContext.setGlobalServiceSupplier( RcdXlsService.class, () -> new RcdPoiXlsService() );
    }

    @Override
    public RcdPoiXlsWorkbook instRead( final InputStream inputStream )
    {
        try
        {
            final Workbook workbook = WorkbookFactory.create( inputStream );
            return new RcdPoiXlsWorkbook( workbook );
        }
        catch ( Exception e )
        {
            throw new RcdException( "Error while reading XLS file", e );
        }
    }
}
