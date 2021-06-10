const {
  TeamsActivityHandler,
  tokenExchangeOperationName,
  ActionTypes,
  CardFactory,
  TextFormatTypes,
  TurnContext,
  TeamsInfo,
  MessageFactory
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
    var members;
    var channels;
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
      await this.updateMembers(TeamsInfo, context);
      await this.processMessage(context);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
    this.onTeamsChannelCreatedEvent(async (channelInfo, TeamInfo, context, next) => {
      await this.updateChannels(TeamsInfo, context);
      await next();
    });
    this.onTeamsChannelDeletedEvent(async (channelInfo, TeamInfo, context, next) => {
      await this.updateChannels(TeamsInfo, context);
      await next();
    });
    this.onTeamsChannelRenamedEvent(async (channelInfo, TeamInfo, context, next) => {
      await this.updateChannels(TeamsInfo, context);
      await next();
    });
    this.onTeamsChannelRestoredEvent(async (channelInfo, TeamInfo, context, next) => {
      await this.updateChannels(TeamsInfo, context);
      await next();
    });

    this.onMembersAdded(async (context, next) => {

      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const cardButtons = [
            { type: ActionTypes.ImBack, title: "Show introduction card", value: "intro" },
          ];
          await this.updateMembers(TeamsInfo, context);
          await this.updateChannels(TeamsInfo, context);
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

    this.onMembersRemoved(async (context, next) => {
      await this.updateMembers(TeamsInfo, context);
      await next();
    });
  }

  async updateMembers(TeamsInfo, context) {
    this.members = await TeamsInfo.getMembers(context);
  }
  async updateChannels(TeamsInfo, context) {
    this.channels = await TeamsInfo.getTeamChannels(context);
    this.channels[0].name = "General";
  }

  //create channel conversation
  async teamsCreateChannelConversation(context, teamsChannelId, message) {
    const conversationParameters = {
      isGroup: true,
      channelData: {
        channel: {
          id: teamsChannelId
        }
      },
      activity: message
    };

    const connectorClient = context.adapter.createConnectorClient(context.activity.serviceUrl);
    const conversationResourceResponse = await connectorClient.conversations.createConversation(conversationParameters);
    const conversationReference = TurnContext.getConversationReference(context.activity);
    conversationReference.conversation.id = conversationResourceResponse.id;
    return [conversationReference, conversationResourceResponse.activityId];
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
      case "message": {
        await this.updateMembers(TeamsInfo, context);

        const teamMember = this.members[17];//Specific Member to interact with (Tyson:3,Ethan:17)

        const message = `Hello ${teamMember.givenName}. I am Testbot.`;
        //message reference??
        var ref = TurnContext.getConversationReference(context.activity);

        ref.user = teamMember;

        //create conversation (message)
        await context.adapter.createConversation(ref,
          async (t1) => {
            const ref2 = TurnContext.getConversationReference(t1.activity);
            await t1.adapter.continueConversation(ref2, async (t2) => {
              await t2.sendActivity(message);
            });
          });
        break;
      }

      case "channel": {
        const teamsChannelId = this.channels[1].id;//Specific Channel to interact with
        const message = MessageFactory.text('This will be the first message in a new thread');

        //Create and store reference to new conversation
        const newConversation = await this.teamsCreateChannelConversation(context, teamsChannelId, message);

        //send response to conversation
        await context.adapter.continueConversation(newConversation[0],
          async (t) => {
            await t.sendActivity(MessageFactory.text('This will be the first response to the new thread'));
          });

        break;
      }

      default: {
        console.log(`Unknown command: ${text}`);
        //Temp just to catch any unrecognized commands
      }
    }
  }
}

module.exports.TeamsBot = TeamsBot;