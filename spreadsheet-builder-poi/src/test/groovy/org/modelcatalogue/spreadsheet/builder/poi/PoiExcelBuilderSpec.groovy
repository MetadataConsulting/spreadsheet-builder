package org.modelcatalogue.spreadsheet.builder.poi

import org.junit.Rule
import org.junit.rules.TemporaryFolder
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.api.ForegroundFill
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetQuery
import spock.lang.Specification

import java.awt.Desktop

import static org.modelcatalogue.spreadsheet.builder.poi.PoiCellDefinition.fixName


class PoiExcelBuilderSpec extends Specification {

    @Rule TemporaryFolder tmp

    def "create sample spreadsheet"() {
        when:
        File tmpFile = tmp.newFile("sample${System.currentTimeMillis()}.xlsx")

        SpreadsheetBuilder builder = new PoiSpreadsheetBuilder()

        Date today = new Date()

        tmpFile.withOutputStream { OutputStream out ->
            builder.build(out) {
                style 'red', {
                    font {
                        color red
                    }
                }
                style 'bold', {
                    font {
                        make bold
                    }
                }

                style 'h1', {
                    font {
                        make bold
                    }
                }

                style 'h2', {
                    font {
                        make bold
                    }
                }

                style "borders", {
                    font {
                        color red
                    }
                    border {
                        style thin
                        color black
                    }
                }


                apply MyStyles // or apply(new MyStyles())

                sheet("many rows"){
                    20000.times{
                        row {
                            cell {
                                value '1'
                                styles 'h1', 'red'
                            }
                            cell {
                                value '2'
                                style 'h2'
                            }
                            cell {
                                value '3'
                                style 'h1'
                            }
                            cell {
                                value '4'
                                style 'h2'
                            }
                        }
                    }
                }

                sheet('Sample') {
                    row {
                        cell {
                            value 'Hello'
                            style 'h1'
                        }
                        cell {
                            value 'World'
                            style 'h2'
                        }
                    }
                }

                sheet('Fonts') {
                    row {
                        cell {
                            width auto
                            value 'Bold Red 22'
                            style {
                                font {
                                    make bold
                                    color red
                                    size 22
                                }
                            }
                        }
                        cell {
                            width auto
                            value 'Underline Courier New'
                            style {
                                font {
                                    make underline
                                    name 'Courier New'
                                }
                            }
                        }
                        cell {
                            width auto
                            value 'Italic'
                            style {
                                font {
                                    make italic
                                }
                            }
                        }
                        cell {
                            width auto
                            value 'Strikeout'
                            style {
                                font {
                                    make strikeout
                                }
                            }
                        }
                    }
                }

                sheet('Image') {
                    row (3) {
                        cell ('C') {
                            png image from 'https://goo.gl/UcL1wy'
                        }
                    }
                }

                sheet('Rich Text') {
                    row {
                        cell {
                            text 'Little'
                            text ' '
                            text 'Red', {
                                color red
                                size 22
                                name "Times New Roman"
                            }
                            text ' '
                            text 'Riding', {
                                make italic
                                size 18
                            }
                            text ' '
                            text 'Hood', {
                                make bold
                                size 22
                            }

                        }
                    }
                    row {
                        cell {
                            style {
                                wrap text
                            }
                            text 'First Line'
                            text '\n'
                            text 'Second Line', {
                                make bold
                                size 12
                            }
                            text '\n'
                            for (Map.Entry<String, String> entry in [foo: 'bar', boo: 'cow', empty: '', '':'nothing']) {
                                text entry.key, {
                                    make bold
                                }
                                text ': '
                                text  entry.value
                                text '\n'
                            }
                            text '\n\n'
                            text 'Next line after two spaces'
                            text '\n'
                            text 'Last line', {
                                make italic
                            }
                            text '\n'
                        }
                    }

                    row {
                        cell {
                            text 'Genomics England Consent Withdrawal Options'
                            text '\n\n'
                            text 'Enumerations', {
                                size 12
                                make bold
                            }
                            text '\n'

                            for (Map.Entry<String, String> entry in [
                                    FULL_WITHDRAWAL: 'OPTION 2: FULL WITHDRAWAL: No further use',
                                    PARTIAL_WITHDRAWAL: 'OPTION 1: PARTIAL WITHDRAWAL: No further contact'
                            ]) {
                                text entry.key, {
                                    make bold
                                }
                                text ': '
                                text  entry.value
                                text '\n'
                            }
                        }
                    }
                }

                sheet('Cell Addressing') {
                    row(2) {
                        style {
                            background whiteSmoke
                            border top, bottom, {
                                style thin
                                color black
                            }
                        }
                        cell('B') {
                            value 'A'
                            style {
                                border left, {
                                    style thin
                                    color black
                                }
                            }
                        }
                        cell 'B'
                        cell {
                            value 'C'
                            style {
                                border right, {
                                    style thin
                                    color black
                                }
                            }
                        }
                    }
                    row {
                        cell('B') { value 1 }
                        cell 2
                        cell 3
                    }
                }
                sheet('One') {
                    freeze 1,1
                    row {
                        cell 'First Row'
                    }

                    row {
                        cell 'AC', {
                            value 'AC'
                        }
                        cell 'BE', {
                            value 'BE'
                        }
                    }

                    row {
                        style {
                            align center left
                            border {
                                color '#abcdef'
                                style dashDotDot
                            }
                            border right, {
                                color '#00ff00'
                            }
                        }
                        cell {
                            value 'Hello'
                            name 'Salutation'
                            width auto
                        }
                        cell {
                            style 'bold'
                            value 'World'
                            comment {
                                text 'This cell has some fancy fg/bg'
                                author 'musketyr'
                            }
                            width 50
                        }
                        cell {
                            style {
                                format 'd.m.y'
                                align center center
                            }
                            value today
                            comment 'This is a date!'
                            colspan 5
                            rowspan 2
                        }
                    }
                }
                sheet('Links') {
                    freeze 1,0
                    row {
                        cell {
                            value 'Document (and a very long text)'
                            link to name 'Salutation'
                            width auto
                        }
                        cell {
                            value 'File'
                            link to file 'text.txt'
                        }
                        cell {
                            value 'URL'
                            link to url 'https://www.google.com'
                        }
                        cell {
                            value 'Mail (plain)'
                            link to email 'vladimir@orany.cz'
                        }
                        cell {
                            value 'Mail (with subject)'
                            link to email 'vladimir@orany.cz', subject: 'Testing Excel Builder', body: 'It is really great tools'
                        }
                    }
                }
                sheet ('Groups'){
                    row {
                        cell "Headline 1"
                        group {
                            cell {
                                value "Headline 2"
                                style {
                                    foreground aquamarine
                                    fill solidForeground
                                }
                            }
                            cell "Headline 3"
                            collapse {
                                cell "Headline 4"
                                cell "Headline 5"
                            }
                            cell "Headline 6"
                        }
                    }
                    group {
                        row {
                            cell "Some stuff"
                        }
                        collapse {
                            row {
                                cell "Something"
                            }
                            row {
                                cell "Something other"
                            }
                        }
                        row {
                            cell "Other stuff"
                        }
                    }
                }
                sheet('Fills') {
                    for(ForegroundFill foregroundFill in ForegroundFill.values()) {
                        row {
                            cell {
                                width auto
                                value foregroundFill.name()
                            }
                            cell {
                                style {
                                    background '#FF8C00' // darkOrange
                                    foreground brown
                                    fill foregroundFill
                                }
                            }
                        }
                    }
                }
                sheet('Formula') {
                    row {
                        cell {
                            value 10
                            name 'Cell10'
                        }
                        cell {
                            value 20
                            name 'Cell20'
                        }
                        cell {
                            formula 'SUM(#{Cell10}:#{Cell20})'
                        }
                    }
                }
                // create sheet with same name, it should query the
                sheet('Formula') {
                    row {
                        cell {
                            value 30
                            name 'Cell30'
                        }
                        cell {
                            value 40
                            name 'Cell40'
                        }
                        cell {
                            formula 'SUM(#{Cell30}:#{Cell40})'
                        }
                    }
                }
                sheet('Traversal') {
                    row {
                        cell 'A'
                        cell 'B'
                        cell 'C'
                    }
                    row {
                        cell 'D'
                        cell {
                            value 'E'
                            colspan 2
                        }
                        cell 'F'
                    }
                    row {
                        cell 'G'
                        cell {
                            value 'H'
                            rowspan(2)
                        }
                        cell 'I'
                    }
                    row(5) {
                        cell('B') {
                            value 'J'
                }   }   }
                sheet('Border') {
                    row {
                        style "borders"
                        cell {
                            value 1
                            colspan(2)
                        }
                        cell {
                            value 2
                        }
                    }
                    row {
                        style "borders"
                        cell {
                            value 1
                            colspan(2)
                        }
                        cell {
                            value 2
                        }

                    }
                }
            }

        }

            SpreadsheetCriteria matcher =  PoiSpreadsheetQuery.FACTORY.forFile(tmpFile)

            Collection<Cell> allCells = matcher.all()

        then:
            allCells
            allCells.size() == 80102

        when:
            Collection<Cell> sampleCells = matcher.query({
                sheet('Sample')
            })
        then:
            sampleCells
            sampleCells.size() == 2

        when:
            Collection<Cell> rowCells = matcher.query({
                sheet("many rows") {
                    row(1)
                }
            })
        then:
            rowCells
            rowCells.size() == 4

        when:
            Collection<Cell> someCells = matcher.query({
                sheet {
                    row {
                        cell {
                            date today
                        }
                    }
                }
            })
        then:
            someCells
            someCells.size() == 1

        when:
            Collection<Cell> commentedCells = matcher.query({
                sheet {
                    row {
                        cell {
                            comment 'This is a date!'
                        }
                    }
                }
            })
        then:
            commentedCells
            commentedCells.size() == 1

        when:
            Collection<Cell> namedCells = matcher.query({
                sheet {
                    row {
                        cell {
                            name fixName("Cell10")
                        }
                    }
                }
            })
        then:
            namedCells
            namedCells.size() == 1
        when:
            Collection<Cell> dateCells = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                format 'd.m.y'
                            }
                        }
                    }
                }
            })
        then:
            dateCells
            dateCells.size() == 1
        when:
            Collection<Cell> filledCells = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                fill fineDots
                            }
                        }
                    }
                }
            })
        then:
            filledCells
            filledCells.size() == 1
        when:
            Collection<Cell> magentaCells = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                foreground aquamarine
                            }
                        }
                    }
                }
            })
        then:
            magentaCells
            magentaCells.size() == 1
        when:
            Collection<Cell> redOnes = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                font {
                                    color red
                                }
                            }
                        }
                    }
                }
            })
        then:
            redOnes
            redOnes.size() == 20005

        when:
            Collection<Cell> boldOnes = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                font {
                                    make bold
                                }
                            }
                        }
                    }
                }
            })
        then:
            boldOnes
            boldOnes.size() == 2

        when:
            Collection<Cell> bigOnes = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                font {
                                    size 22
                                }
                            }
                        }
                    }
                }
            })
        then:
            bigOnes
            bigOnes.size() == 40002


        when:
            Collection<Cell> bordered = matcher.query({
                sheet {
                    row {
                        cell {
                            style {
                                border(top) {
                                    style thin
                                }
                            }
                        }
                    }
                }
            })
        then:
            bordered
            bordered.size() == 9
        when:
            Collection<Cell> combined = matcher.query({
                sheet {
                    row {
                        cell {
                            value 'Bold Red 22'
                            style {
                                font {
                                    color red
                                }
                            }
                        }
                    }
                }
            })
        then:
            combined
            combined.size() == 1


        when:
            Collection<Cell> conjunction = matcher.query({
                sheet {
                    row {
                        or {
                            cell {
                                value 'Bold Red 22'
                            }
                            cell {
                                value 'A'
                            }
                        }
                    }
                }
            })
        then:
            conjunction
            conjunction.size() == 3

        when:
            Collection<Cell> traversal = matcher.query({
                sheet('Traversal') {
                    row {
                        cell {
                            value 'E'
            }   }   }   })
        then:
            traversal
            traversal.size() == 1

        when:
            Cell cellE = traversal.first()
        then:
            cellE.row.sheet.name == 'Traversal'
            cellE.row.sheet.previous()
            cellE.row.sheet.previous().name == 'Formula'
            cellE.row.sheet.next()
            cellE.row.sheet.next().name == 'Border'
            cellE.row.number == 2
            cellE.row.above()
            cellE.row.above().number == 1
            cellE.row.bellow()
            cellE.row.bellow().number == 3
            cellE.colspan == 2
            cellE.aboveLeft()
            cellE.aboveLeft().value == 'A'
            cellE.above()
            cellE.above().value == 'B'
            cellE.aboveRight()
            cellE.aboveRight().value == 'C'
            cellE.left()
            cellE.left().value == 'D'
            cellE.right()
            cellE.right().value == 'F'
            cellE.bellowLeft()
            cellE.bellowLeft().value == 'G'
            cellE.bellowRight()
            cellE.bellowRight().value == 'I'
            cellE.bellow()
            cellE.bellow().value == 'H'
            cellE.bellow().bellow().value == 'J'

        when:
            open tmpFile
        then:
            noExceptionThrown()

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
