// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// bot.js is your bot's main entry point to handle incoming activities.

const { ActivityTypes } = require('botbuilder');
const { adal } = require('adal-node');

// Turn counter property
const TURN_COUNTER_PROPERTY = 'turnCounterProperty';

class EchoBot {
    /**
     *
     * @param {ConversationState} conversation state object
     */
    constructor(conversationState) {
        // Creates a new state accessor property.
        // See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors
        this.countProperty = conversationState.createProperty(TURN_COUNTER_PROPERTY);
        this.conversationState = conversationState;
    }
    /**
     *
     * Use onTurn to handle an incoming activity, received from a user, process it, and reply as needed
     *
     * @param {TurnContext} on turn context object.
     */
    async onTurn(turnContext) {

        // Handle message activity type. User's responses via text or speech or card interactions flow back to the bot as Message activity.
        // Message activities may contain text, speech, interactive cards, and binary or unknown attachments.
        // see https://aka.ms/about-bot-activity-message to learn more about the message and other activity types
        if (turnContext.activity.type === ActivityTypes.Message) {

            // Update SharePoint list

            var AuthenticationContext = adal.AuthenticationContext;
            var authorityHostUrl = 'https://login.windows.net';
            var tenant = 'sep007.onmicrosoft.com/';
            console.log('ashwin');
            var authoriotyUrl = authorityHostUrl + '/' + tenant;
           // var applicationId = clientID;
           // var clientSecret = appSecret;
            var applicationId = "d2d16b32-5c6e-47c8-bafe-11d0c32dc588";
            var clientSecret = "unIx9CLxX+HOJNbPNHIKO+cuRqBlPD2pDvH4jWjWMEA=";
            const resource = "https://graph.microsoft.com";
           // var messageText = turnContext.activity.text;
            //messageText = messageText.substring(mentionString.length);
            var context = new AuthenticationContext(authoriotyUrl);
            context.acquireTokenWithClientCredentials(
                resource,
                applicationId,
                clientSecret,
                function (err, tokenResponse) {
                    if (err) {
                        console.log('well that did not work: ' + err.stack);
                    } else {
                        let client = MicrosoftGraph.Client.init({
                            authProvider: (done) => {
                                console.log(tokenResponse.accessToken);
                                done(null, tokenResponse.accessToken);
                            }
                        });
                        // client.api('https://graph.microsoft.com/beta/sites/sep007.sharepoint.com,052e0c8e-9e92-4165-8d9a-c4f405aaf2d8/lists/test/items/')
                        client.api('https://graph.microsoft.com/beta/sites/sep007.sharepoint.com:/sites/dev:/lists/test/items/')
                            .version("beta")
                            .header("Content-type", "application/json")
                            .post({
                                "fields": {
                                    "ContentType": "Item",
                                    "Title": "check it out"
                                }
                            }).then((res) => {
                                await turnContext.sendActivity("your message has been posted to SharePoint");
                            }).catch((err) => {
                                console.log(err);
                                await turnContext.sendActivity("Oops ! error ocured");
                            });
                    }
                });


            // read from state.
            let count = await this.countProperty.get(turnContext);
            count = count === undefined ? 1 : ++count;
            await turnContext.sendActivity(`${count}: You said 1 "${turnContext.activity.text}"`);
            // increment and set turn counter.
            await this.countProperty.set(turnContext, count);
        } else {
            // Generic handler for all other activity types.
            await turnContext.sendActivity(`[${turnContext.activity.type} event detected]`);
        }
        // Save state changes
        await this.conversationState.saveChanges(turnContext);
    }
}

exports.EchoBot = EchoBot;
