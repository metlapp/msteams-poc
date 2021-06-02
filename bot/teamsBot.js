const {
  TeamsActivityHandler,
  tokenExchangeOperationName,
  ActionTypes,
  CardFactory,
  TextFormatTypes,
  TurnContext,
  TeamsInfo
} = require("botbuilder");
const { connect } = require("ngrok");

class TeamsBot extends TeamsActivityHandler {
  /**
   *
   * @param {ConversationState} conversationState
   * @param {UserState} userState
   *
   */
  constructor(conversationState, userState) {
    super();
    if (!conversationState) {
      throw new Error("[TeamsBot]: Missing parameter. conversationState is required");
    }
    if (!userState) {
      throw new Error("[TeamsBot]: Missing parameter. userState is required");
    }

    this.conversationState = conversationState;
    this.userState = userState;

    this.onMessage(async (context, next) => {
      //Process incoming messages
      const teamDetails = await TeamsInfo.getTeamDetails(context);
      if (teamDetails) {
          await context.sendActivity(`The group ID is: ${teamDetails.aadGroupId}`);
      } else {
          await context.sendActivity('This message did not come from a channel in a team.');
      }


      await this.processMessage(context);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      //TODO probably remove this?

      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const cardButtons = [
            { type: ActionTypes.ImBack, title: "Show introduction card", value: "intro" },
          ];
          const card = CardFactory.heroCard("Welcome", null, cardButtons, {
            text: `Congratulations! Your hello world Bot 
                            template is running. This bot has default commands to help you modify it.
                            You can reply <strong>intro</strong> to see the introduction card. This bot is built with <a href=\"https://dev.botframework.com/\">Microsoft Bot Framework</a>`,
          });
          await context.sendActivity({ attachments: [card] });
          break;
        }
      }
      await next();
    });
  }

  //Process an incoming message
  async processMessage(context) {
    //Check if we need to remove an @ tag
    const removedMentionText = TurnContext.removeRecipientMention(
      context.activity,
      context.activity.recipient.id
    );
    let text = "";
    //Remove the @ tag if required
    if (removedMentionText) {
      text = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim(); // Remove the line break
    }

    //Ensure the message we received is a text message
    if (context.activity.textFormat !== TextFormatTypes.Plain) {
      console.log("Not a text message");
    }

    //Figure out which command was used
    console.log(`Requested command: ${text}`);
    switch (text) {
      case "demo": {
        await context.sendActivity("Demo command success!");
        break;
      }
      case "breaktime": {
        await context.sendActivity("icebreaker command success!");
        break;
      }

      case "test": {
        var title = "Sample Ice Breaker";
        var description = "This is a sample ice breaker that will hopefully be modifiable later. Maybe we can create a template card and then just send JSON data through to the bot API?";
        const dateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        const dateStr = new Date().toLocaleDateString(undefined, dateOptions);
      
        await context.sendActivity({
          attachments: [
          {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
              {
              "type": "TextBlock",
              "size": "Medium",
              "weight": "Bolder",
              "text": `${title}`
              },
              {
              "type": "ColumnSet",
              "columns": [
                {
                "type": "Column",
                "items": [
                  {
                  "type": "Image",
                  "style": "Person",
                  "url": `https://avatars.slack-edge.com/2021-03-02/1820480857892_f5ff53aaec7a5507e5ad_512.png`,
                  "size": "Small"
                  }
                ],
                "width": "auto"
                },
                {
                "type": "Column",
                "items": [
                  {
                  "type": "TextBlock",
                  "weight": "Bolder",
                  "text": `Metl`,
                  "wrap": true
                  },
                  {
                  "type": "TextBlock",
                  "spacing": "None",
                  "text": `Created ${dateStr}`,
                  "isSubtle": true,
                  "wrap": true
                  }
                ],
                "width": "stretch"
                }
              ]
              },
              {
              "type": "TextBlock",
              "text": `${description}`,
              "wrap": true
              },
            ],
            "actions": [
              {
                "type": "Action.Submit",
                "title": "Dog",
                "data": {cardAction: "update", value: "Dog"}
              },
              {
                "type": "Action.Submit",
                "title": "Cat",
                "data": {cardAction: "update", value: "Cat"}
              },
              {
                "type": "Action.Submit",
                "title": "Fish",
                "data": {cardAction: "update", value: "Fish"}
              }
            ]
            }
          }
          ]
        });
      }


      default: {
        console.log(`Unknown command: ${text}`);
        //Temp just to catch any unrecognized commands
      }
    }
  }
}

module.exports.TeamsBot = TeamsBot;