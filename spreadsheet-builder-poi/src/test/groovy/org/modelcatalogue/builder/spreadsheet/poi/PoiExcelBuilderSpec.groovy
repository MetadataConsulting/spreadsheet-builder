package org.modelcatalogue.builder.spreadsheet.poi

import org.junit.Rule
import org.junit.rules.TemporaryFolder
import org.modelcatalogue.builder.spreadsheet.api.SpreadsheetBuilder
import org.modelcatalogue.builder.spreadsheet.api.ForegroundFill
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
                        bold
                    }
                }

                sheet('Sample') {
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
            }
        }

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
