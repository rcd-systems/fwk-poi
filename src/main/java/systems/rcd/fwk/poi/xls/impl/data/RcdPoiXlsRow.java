package systems.rcd.fwk.poi.xls.impl.data;

import java.time.Instant;
import java.time.LocalDateTime;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import systems.rcd.fwk.core.format.xls.data.RcdXlsCell;
import systems.rcd.fwk.core.format.xls.data.RcdXlsRow;

public class RcdPoiXlsRow
    extends ArrayList<RcdXlsCell>
    implements RcdXlsRow
{

    public RcdPoiXlsRow( final Row row )
    {
        super( row.getLastCellNum() );
        for ( int i = 0; i < row.getLastCellNum(); i++ )
        {
            final Cell cell = row.getCell( i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL );
            if ( cell == null || cell.getCellType() == CellType.ERROR )
            {
                add( null );
            }
            else
            {
                RcdPoiXlsCell xlsCell = null;
                try
                {
                    xlsCell = new RcdPoiXlsCell( cell );
                }
                catch ( final Exception e )
                {
                }
                add( xlsCell );
            }
        }
    }

    @Override
    public String getString( final int index )
    {
        final RcdXlsCell cell = get( index );
        return cell == null ? null : cell.getStringValue();
    }

    ;

    @Override
    public LocalDateTime getDateTime( final int index )
    {
        final RcdXlsCell cell = get( index );
        return cell == null ? null : cell.getDateTimeValue();
    }

    @Override
    public Double getNumber( final int index )
    {
        final RcdXlsCell cell = get( index );
        return cell == null ? null : cell.getNumericValue();
    }

    @Override
    public Boolean getBoolean( final int index )
    {
        final RcdXlsCell cell = get( index );
        return cell == null ? null : cell.getBooleanValue();
    }

    @Override
    public RcdXlsCell get( final int index )
    {
        return size() > index ? super.get( index ) : null;
    }
}
