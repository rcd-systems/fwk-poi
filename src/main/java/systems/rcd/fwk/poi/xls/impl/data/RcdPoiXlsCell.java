package systems.rcd.fwk.poi.xls.impl.data;

import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import systems.rcd.fwk.core.format.xls.data.RcdXlsCell;
import systems.rcd.fwk.core.format.xls.data.RcdXlsCellType;

public class RcdPoiXlsCell
    implements RcdXlsCell
{

    private String stringValue;

    private LocalDateTime localDateTimeValue;

    private Double numericValue;

    private Boolean booleanValue;

    public RcdPoiXlsCell( final Cell cell )
    {
        switch ( cell.getCellType() )
        {
            case STRING:
                stringValue = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if ( DateUtil.isCellDateFormatted( cell ) )
                {
                    localDateTimeValue = cell.getLocalDateTimeCellValue();
                }
                else
                {
                    numericValue = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                break;
            case FORMULA:

                try
                {
                    booleanValue = cell.getBooleanCellValue();
                    break;
                }
                catch ( final Exception e )
                {
                }
                try
                {
                    numericValue = cell.getNumericCellValue();
                    break;
                }
                catch ( final Exception e )
                {
                }
                stringValue = cell.getRichStringCellValue().toString();
                break;
        }
    }

    @Override
    public RcdXlsCellType getType()
    {
        if ( stringValue != null )
        {
            return RcdXlsCellType.STRING;
        }
        if ( localDateTimeValue != null )
        {
            return RcdXlsCellType.DATETIME;
        }
        if ( numericValue != null )
        {
            return RcdXlsCellType.NUMBER;
        }
        return RcdXlsCellType.BOOLEAN;
    }

    @Override
    public String getStringValue()
    {
        return stringValue;
    }

    @Override
    public LocalDateTime getDateTimeValue()
    {
        return localDateTimeValue;
    }

    @Override
    public Double getNumericValue()
    {
        return numericValue;
    }

    @Override
    public Boolean getBooleanValue()
    {
        return booleanValue;
    }

    @Override
    public String toString()
    {
        if ( stringValue != null )
        {
            return stringValue;
        }
        if ( localDateTimeValue != null )
        {
            return localDateTimeValue.toString();
        }
        if ( numericValue != null )
        {
            return numericValue.toString();
        }
        return booleanValue.toString();
    }
}
