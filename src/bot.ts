// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { ActivityHandler, MessageFactory, CardFactory, InvokeResponse, AdaptiveCardInvokeResponse, TurnContext, InvokeException } from 'botbuilder';
import * as ACData from "adaptivecards-templating";
import * as cardOnefile from "./cards/cardOne.json";
import * as cardTwofile from "./cards/cardTwo.json";
import * as cardSenderFile from "./cards/cardSender.json";
const CONVERSATION_DATA_PROPERTY = 'conversationData';
export class EchoBot extends ActivityHandler {
    conversationData: any;
    conversationState: any;
    conversationDataAccessor: any;
    constructor(conversationState) {
        super();
        // Create conversation object
        this.conversationState = conversationState;
        //Conversation data accessor
        this.conversationDataAccessor = conversationState.createProperty(CONVERSATION_DATA_PROPERTY);

        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            // const replyText = `Echo: ${ context.activity.text }`;
            // await context.sendActivity(MessageFactory.text(replyText, replyText));
            //empty access for conversation data
            const conversationData = await this.conversationDataAccessor.get(context, {});
            //initial card or the base card   
            var template = new ACData.Template(cardOnefile);
            var userIds = [context.activity.from.id];
            //save the initiator id as the person who called the Bot          
            conversationData.initiator = context.activity.from.id;
            //count of how many turns (to calculate the number of clicks the card got)
            conversationData.turnCount = 0;
           //remove the bot at mentioned from the activity message
            const updatedText = TurnContext.removeRecipientMention(context.activity);
            conversationData.question=updatedText;
            conversationData.voted = [context.activity.from.id];
            //set the conversation data   
            await this.conversationDataAccessor.set(context, conversationData);
            var cardPayload = template.expand({ $root: { question:conversationData.question,userIds: userIds } });
            const cardOne = CardFactory.adaptiveCard(cardPayload);
            cardOne.content.subtitle = "";
            const card = {
                contentType: cardOne.contentType,
                content: cardOne.content
                
            };
            await context.sendActivity({ attachments: [card] });
            
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });



    }
    //UAM onInvokeActivity overriden
    async onInvokeActivity(context: TurnContext) {
        try {
            const request = context.activity.value;
            
            if (request) {
            const conversationData = await this.conversationDataAccessor.get(context); 
            let turnCount = conversationData.turnCount || 0;             
            var responseBody = {}
            var payload={};            
                switch (request.action.verb) {
                    case 'vote': {
                        //on button click , count the number of click and store it in conversation data
                        conversationData.turnCount = ++turnCount;
                        conversationData.voted.push(context.activity.from.id);
                        await this.conversationDataAccessor.set(context, conversationData);
                        payload = await this.processSend(conversationData);                       
                        const cardOne = CardFactory.adaptiveCard(payload);
                        const card = {
                          contentType: cardOne.contentType,
                          content: cardOne.content
                      };
                    const message =MessageFactory.attachment(card);                      
                    message.id = context.activity.replyToId;
                    await context.updateActivity(message);
                    break;    
                    }               
                    case 'refresh': {
                         payload = await this.processRefresh(context.activity.from.id,conversationData);
                        break;                        
                    }              
                    
                    default:
                        throw new InvokeException(404);
                }
                responseBody= { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: payload }
                return this.createInvokeResponse(responseBody);

            }

        } catch (err) {

            throw err;
        } finally {
            this.defaultNextEvent(context)();
        }
    }

    createInvokeResponse(body?: any): InvokeResponse {
        return { status: 200, body };
    }
    //send the base card
    async processSend(conversationData) {
        var template = new ACData.Template(cardOnefile);        
        var cardPayload = template.expand({ $root: { question:conversationData.question,userIds: conversationData.voted } });
        return cardPayload
    }
    //process refresh of cards based on the userIds array
    async processRefresh(userId, conversationData) {
      
        var  template;  
        //initiator
        if(conversationData.initiator===userId)   {
         template = new ACData.Template(cardSenderFile);
         return template.expand({ $root: {question:conversationData.question,turn: conversationData.turnCount, userIds:conversationData.voted } });
           
        } //already voted  
        else if (conversationData.voted.indexOf(userId)>-1&&conversationData.initiator!==userId){
             template = new ACData.Template(cardTwofile);
             return template.expand({ $root: { question:conversationData.question, userIds: conversationData.voted } });
          
         }   
       
    }
    /**
   * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
   */
    async run(context) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);

    }
}



