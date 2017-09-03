package builders.dsl.spreadsheet.builder.poi

import org.junit.Rule
import org.junit.rules.TemporaryFolder
import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder
import builders.dsl.spreadsheet.builder.tck.AbstractBuilderSpec
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria

import java.awt.Desktop

class PoiExcelBuilderSpec extends AbstractBuilderSpec {

    @Rule TemporaryFolder tmp = new TemporaryFolder()

    File tmpFile

    void setup() {
        tmpFile = tmp.newFile("sample${System.currentTimeMillis()}.xlsx")
    }

    @Override
    protected SpreadsheetCriteria createCriteria() {
        return PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile)
    }

    @Override
    protected SpreadsheetBuilder createSpreadsheetBuilder() {
        return PoiSpreadsheetBuilder.create(tmpFile)
    }

    @Override
    protected void openSpreadsheet() {
        open tmpFile
    }

    /**
     * Tries to open the file in Word. Only works locally on Mac at the moment. Ignored otherwise.
     * Main purpose of this method is to quickly open the generated file for manual review.
     * @param file file to be opened
     */
    private static void open(File file) {
        try {
            if (Desktop.desktopSupported && Desktop.desktop.isSupported(Desktop.Action.OPEN)) {
                Desktop.desktop.open(file)
                Thread.sleep(10000)
            }
        } catch(ignored) {
            // CI
        }
    }
}
