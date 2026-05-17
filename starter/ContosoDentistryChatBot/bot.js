// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

const { QnAMaker, LuisRecognizer } = require('botbuilder-ai');
const DentistScheduler = require('./dentistscheduler');
const IntentRecognizer = require("./intentrecognizer")

class DentaBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        // call the parent constructor
        super();
        if (!configuration) throw new Error('[QnaMakerBot]: Missing parameter. configuration is required');

        // create a QnAMaker connector
        this.QnAMaker = new QnAMaker(configuration.QnAConfiguration, qnaOptions);

        // create a DentistScheduler connector
        this.DentistScheduler = new DentistScheduler(configuration.SchedulerConfiguration);

        // create a IntentRecognizer connector
        this.IntentRecognizer = new IntentRecognizer(configuration.LuisConfiguration);


        this.onMessage(async (context, next) => {
            const qnaResults = await this.QnAMaker.getAnswers(context);
            const luisResults = await this.IntentRecognizer.executeLuisQuery(context);
                     
            // determine which service to respond with based on the results from LUIS //

            const topIntent = LuisRecognizer.topIntent(luisResults);

            if (topIntent === 'GetAvailability' && luisResults.intents[topIntent].score > 0.5) {
                const availability = await this.DentistScheduler.getAvailability();
                await context.sendActivity(MessageFactory.text(availability, availability));
                await next();
                return;
            } else if (topIntent === 'ScheduleAppointment' && luisResults.intents[topIntent].score > 0.5) {
                const time = this.IntentRecognizer.getTimeEntity(luisResults);
                const appointment = await this.DentistScheduler.scheduleAppointment(time);
                await context.sendActivity(MessageFactory.text(appointment, appointment));
                await next();
                return;
            } else {
                if (qnaResults[0]) {
                    await context.sendActivity(MessageFactory.text(qnaResults[0].answer, qnaResults[0].answer));
                } else {
                    await context.sendActivity(MessageFactory.text("I'm sorry, I didn't understand that. Please try asking about availability or scheduling an appointment."));
                }
                await next();
            }

             
            
    });

        this.onMembersAdded(async (context, next) => {
        const membersAdded = context.activity.membersAdded;
        //write a custom greeting
        const welcomeText = 'Welcome to Contoso Dentistry! I can help you find out about our services, check our availability, or schedule an appointment. Try saying "What services do you offer?" or "Let me book an appointment on Monday at 10am." How can I assist you today?';
        for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
            if (membersAdded[cnt].id !== context.activity.recipient.id) {
                await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
            }
        }
        // by calling next() you ensure that the next BotHandler is run.
        await next();
    });
    }
}

module.exports.DentaBot = DentaBot;
