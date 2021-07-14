// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { ActivityHandler, MessageFactory, CardFactory, InvokeResponse, AdaptiveCardInvokeResponse, TurnContext, InvokeException } from 'botbuilder';
import * as ACData from "adaptivecards-templating";
import * as cardOnefile from "./cards/cardOne.json"
import * as cardTwofile from "./cards/cardTwo.json"
import * as cardSenderFile from "./cards/cardSender.json"
var initiator='8:orgid:5180d0c1-eea1-4882-a884-3f053eda4f15';

export class EchoBot extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            // const replyText = `Echo: ${ context.activity.text }`;
            // await context.sendActivity(MessageFactory.text(replyText, replyText));
            console.log('here on message'+context.activity.from.id)
          
            var template = new ACData.Template(cardOnefile);
            var userIds=[context.activity.from.id];
            var cardPayload = template.expand({$root: {userIds:userIds}}); 
            const cardOne = CardFactory.adaptiveCard(cardPayload);
            cardOne.content.subtitle = "";
            const card = {
            contentType: cardOne.contentType,
            content: cardOne.content,
            preview: cardOne,
            };
            await context.sendActivity({ attachments: [card] });
            // By calling next() you ensure that the next BotHandler is run.
           // await next();
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
            //await next();
        });

        
 
}
async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    try {
      const request=context.activity.value;
      console.log(JSON.stringify(context.activity))
    
      if(request){
        switch (request.action.verb) {
          case 'send': {          
                  
              var responseBody:AdaptiveCardInvokeResponse= await this.processSend(cardTwofile,request.action.data);
              return this.createInvokeResponse(responseBody);
            }
            
            case 'testAgain':{              
              var responseBody:AdaptiveCardInvokeResponse= await this.processRefresh(context.activity.from.aadObjectId,request.action.data);
              return this.createInvokeResponse(responseBody);
            }
             case 'refresh':{
             
              var responseBody:AdaptiveCardInvokeResponse= await this.processRefresh(context.activity.from.aadObjectId,request.action.data);
              return this.createInvokeResponse(responseBody);
            }
         
          default:
              throw new InvokeException(404);
      }

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
async processSend(body,data):Promise<AdaptiveCardInvokeResponse> { 
var template = new ACData.Template(body);
var cardPayload = template.expand({$root: {text:data.text,userIds:[]}}); 

return Promise.resolve({ statusCode:200,type:"application/vnd.microsoft.card.adaptive",value: cardPayload });
         
}
async processRefresh(userId,data):Promise<AdaptiveCardInvokeResponse> { 
  var body = {};
  var cardPayload,template;
  if (initiator.includes(userId) ){  
    cardPayload= cardSenderFile
  }  
  else{ 
    template = new ACData.Template(cardOnefile);
    var userIds=[userId];
    cardPayload = template.expand({$root: {text:data.text,userIds:userIds}}); 
  } 
  return Promise.resolve({ statusCode:200,type:"application/vnd.microsoft.card.adaptive",value: cardPayload });
           
  }
    }
