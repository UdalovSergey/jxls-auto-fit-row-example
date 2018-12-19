import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.CTRowImpl;

import javax.xml.namespace.QName;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new Main().doExport();
    }

    private void doExport() throws IOException, InvalidFormatException {
        try (InputStream is = getClass().getResourceAsStream("wrap.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/plain-wrap.xlsx")) {
                Transformer transformer = PoiTransformer.createTransformer(is, os);
                XlsCommentAreaBuilder.addCommandMapping("autoSize", AutoSizeCommand.class);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                context.putVar("test", "very very very long string which should to be wrapped");
                xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
                context.getConfig().setIsFormulaProcessingRequired(false);
                ((PoiTransformer) transformer).getWorkbook().write(os);
            }
        }
    }

    public static class AutoSizeCommand extends AbstractCommand {

        private Area area;

        @Override
        public String getName() {
            return "autoSize";
        }

        @Override
        public Size applyAt(CellRef cellRef, Context context) {
            Size size = area.applyAt(cellRef, context);

            PoiTransformer transformer = (PoiTransformer) area.getTransformer();
            Workbook workbook = transformer.getWorkbook();
            Row row = workbook.getSheet(cellRef.getSheetName()).getRow(cellRef.getRow());
            row.setHeight((short) -1);
            removeDyDescentAttr(row);
            Cell cell = row.getCell(cellRef.getCol());
            cell.getCellStyle().setWrapText(true);
            return size;
        }

        @Override
        public Command addArea(Area area) {
            super.addArea(area);
            this.area = area;
            return this;
        }

        private void removeDyDescentAttr(Row row) {
            if (row instanceof XSSFRow) {
                XSSFRow xssfRow = (XSSFRow) row;
                CTRowImpl ctRow = (CTRowImpl) xssfRow.getCTRow();
                QName dyDescent = new QName("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                if (ctRow.get_store().find_attribute_user(dyDescent) != null) {
                    ctRow.get_store().remove_attribute(dyDescent);
                }
            } else {
                System.out.println("This method applicable only for xlsx-templates");
            }
        }


    }
}
