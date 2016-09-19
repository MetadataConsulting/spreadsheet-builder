package org.modelcatalogue.spreadsheet.builder.poi

import org.junit.Rule
import org.junit.rules.TemporaryFolder
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.CellMatcher
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.api.ForegroundFill
import spock.lang.Specification

import java.awt.Desktop


class PoiExcelBuilderSpec extends Specification {

    @Rule TemporaryFolder tmp

    def "create sample spreadsheet"() {
        when:
        File tmpFile = tmp.newFile("sample${System.currentTimeMillis()}.xlsx")

        SpreadsheetBuilder builder = new PoiSpreadsheetBuilder()

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
                            value new Date()
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
                                    foreground magenta
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
                // create sheet with same name, it should match the
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

            CellMatcher matcher =  new PoiCellMatcher()

            Collection<CellDefinition> allCells = matcher.match(tmpFile) {

            }

        then:
            allCells
            allCells.size() == 80092

//        when:
//            Collection<Cell> sampleCells = matcher.match(tmpFile) {
//                sheet('Sample') {
//
//                }
//            }
//        then:
//            sampleCells
//            sampleCells.size() == 2

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
