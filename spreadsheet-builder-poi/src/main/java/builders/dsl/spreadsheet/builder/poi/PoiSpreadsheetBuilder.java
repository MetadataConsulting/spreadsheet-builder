package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder;
import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import builders.dsl.spreadsheet.api.Configurer;

import java.io.*;

public class PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    public static SpreadsheetBuilder create(OutputStream out) {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(), out);
    }

    public static SpreadsheetBuilder create(File file) throws FileNotFoundException {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(), new FileOutputStream(file));
    }

    public static SpreadsheetBuilder create(OutputStream out, InputStream template) throws IOException {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(template), out);
    }

    public static SpreadsheetBuilder create(File file, InputStream template) throws IOException {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(template), new FileOutputStream(file));
    }

    public static SpreadsheetBuilder create(OutputStream out, File template) throws IOException, InvalidFormatException {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(template), out);
    }

    public static SpreadsheetBuilder create(File file, File template) throws IOException, InvalidFormatException {
        return new PoiSpreadsheetBuilder(new XSSFWorkbook(template), new FileOutputStream(file));
    }

    private final XSSFWorkbook workbook;
    private final OutputStream outputStream;

    private PoiSpreadsheetBuilder(XSSFWorkbook workbook, OutputStream outputStream) {
        this.workbook = workbook;
        this.outputStream = outputStream;
    }

    @Override
    public void build(Configurer<WorkbookDefinition> workbookDefinition) {
        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook);
        Configurer.Runner.doConfigure(workbookDefinition, poiWorkbook);
        poiWorkbook.resolve();
        writeTo(outputStream);
    }

    private void writeTo(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException ignored) {
                    // do nothing
                }
            }
        }
    }

}
