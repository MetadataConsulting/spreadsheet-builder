package org.modelcatalogue.spreadsheet.builder.api

import org.modelcatalogue.spreadsheet.api.Configurer
import spock.lang.Specification

class ConfigurerSpec extends Specification {

    void 'set delegate to closure from variable'() {
        given:
            Message message = new Message()
            Configurer<Message> messageConfigurer = {
                text = "Hello World"
            }
        when:
            Configurer.Runner.doConfigure(messageConfigurer, message)
        then:
            message.text == "Hello World"
    }

}

class Message {
    String text
}
