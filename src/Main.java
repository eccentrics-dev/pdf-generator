package src;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.*;
import org.apache.pdfbox.pdmodel.graphics.state.PDExtendedGraphicsState;
import org.apache.poi.ss.usermodel.*;
import src.dto.Template01Dto;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        // テンプレート用Excelを元にExcelファイル作成
        Template01Dto data = create();
        Workbook workbook = WorkbookFactory.create(new File("./resources/template01.xlsx"));
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(1);
        if (row == null) throw new Exception("Failed to Create Excel");

        Cell nameCell = row.getCell(2);
        if (nameCell == null) throw new Exception("Failed to Create Excel");

        nameCell.setCellValue(data.getName());

        int baseRow = 6;
        for (int i = 0; i < data.getRoutes().size(); i++) {
            Template01Dto.Route rowData = data.getRoutes().get(i);
            Row currentRow = sheet.getRow(baseRow + i);
            currentRow.getCell(1).setCellValue(rowData.getName());
            currentRow.getCell(2).setCellValue(rowData.getId());
            currentRow.getCell(3).setCellValue(rowData.getStatus());
        }

        int maxRow = 100;
        int maxColumn = 20;
        for(int rowNum = 0; rowNum < maxRow; rowNum++) {
            Row tmpRow = sheet.getRow(rowNum);
            if (tmpRow == null) {
                tmpRow = sheet.createRow(rowNum);
            }
            for(int colNum = 0; colNum < maxColumn; colNum++) {
                Cell tmpCell = tmpRow.getCell(colNum);
                if (tmpCell == null) {
                    tmpRow.createCell(colNum);
                }
            }
        }

        FileOutputStream out = new FileOutputStream("./output/template01.xlsx");
        workbook.write(out);

        // PDFの作成
        PDDocument doc = new PDDocument();
        PDPage pg = new PDPage(PDRectangle.A4);
        doc.addPage(pg);
        PDPageContentStream cs = new PDPageContentStream(doc, pg);

        PDFont font = PDType0Font.load(doc, new FileInputStream("./resources/LINESeedJP_TTF_Rg.ttf"), true);
//        TrueTypeCollection ttc = new TrueTypeCollection(new File("./resources/meiryo.ttc"));
//        TrueTypeFont ttf = ttc.getFontByName("Meiryo");
//        PDFont font = PDType0Font.load(doc, ttf, true);

        PDFontDescriptor descriptor = font.getFontDescriptor();
        // descriptor.setFontWeight(800);
        descriptor.setFontFamily("Meiryo");

        PDExtendedGraphicsState gs = new PDExtendedGraphicsState();
        gs.setLineWidth(0.1f);
        gs.setNonStrokingAlphaConstant(0.5f);
        gs.setStrokingAlphaConstant(0.5f);
        // cs.setGraphicsStateParameters(gs);

        float x = pg.getMediaBox().getWidth();
        float y = pg.getMediaBox().getHeight();

        cs.setStrokingColor(Color.BLACK);
        cs.setNonStrokingColor(Color.BLACK);
        cs.setLineWidth(0.1f);
        cs.moveTo(5, y - 10);
        cs.lineTo(100, y - 10);
        cs.lineTo(100, y - 28);
        cs.lineTo(5, y - 28);
        cs.lineTo(5, y - 10);
        cs.moveTo(5, y - 28);
        cs.lineTo(100, y - 28);
        cs.lineTo(100, y - 48);
        cs.lineTo(5, y - 48);
        cs.lineTo(5, y - 28);
        cs.stroke();

        cs.beginText();
        cs.setFont(font, 11.5f);
        cs.newLineAtOffset(6, y - 24);
        cs.showText("製品名");
        cs.endText();

        cs.beginText();
        cs.newLineAtOffset(6, y - 42);
        cs.showText("XXX/#101");
        cs.endText();

        int rows = sheet.getPhysicalNumberOfRows();
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();
        for(int rowIndex = 0; rowIndex < rows; rowIndex++) {
            Row r = sheet.getRow(rowIndex);
            Cell c = r.getCell(0);
            float cellWidth = sheet.getColumnWidthInPixels(0);
            float cellHeight = r.getHeightInPoints();
        }
        cs.close();

        doc.save("./output/demo.pdf");

    }

    private static Template01Dto create() {
        var routes = List.of(
                new Template01Dto.Route("AAAAA", "ID#001", "2023/07/01"),
                new Template01Dto.Route("BBBBB", "ID#002", "2023/07/02"),
                new Template01Dto.Route("CCCCC", "ID#003", "2023/07/03"),
                new Template01Dto.Route("DDDDD", "ID#004", "2023/07/04"),
                new Template01Dto.Route("EEEEE", "ID#005", "2023/07/05")
        );
        return Template01Dto.builder()
                .name("ABCD/#1001")
                .routes(routes)
                .build();
    }
}
