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
const {
	connect
} = require("ngrok");
const axios = require('axios');
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
					const cardButtons = [{
						type: ActionTypes.ImBack,
						title: "Show introduction card",
						value: "intro"
					},];
					await this.updateMembers(TeamsInfo, context);
					await this.updateChannels(TeamsInfo, context);

					const card = CardFactory.heroCard("Welcome", null, cardButtons, {
						text: `Congratulations! Your hello world Bot 
                            template is running. This bot has default commands to help you modify it.
                            You can reply <strong>intro</strong> to see the introduction card. This bot is built with <a href=\"https://dev.botframework.com/\">Microsoft Bot Framework</a>`,
					});
					await context.sendActivity({
						attachments: [card]
					});
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
		//console.log(await TeamsInfo.getTeamDetails(context));
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

		//check for non text message
		if (context.activity.textFormat !== TextFormatTypes.Plain) {
			axios
				.post('http://[::]:6969/api/icebreaker-response', {
					id: context.activity.value.id,
					answer: context.activity.value.answer,
					from: context.activity.from.id
				})
				.then(res => {

				})
				.catch(error => {
					console.error(error);
				});

		} else {

			if (context.activity.value) {
				console.log(context.activity.value);
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
				case "refresh": {
					await this.updateMembers(TeamsInfo, context);
					await this.updateChannels(TeamsInfo, context);
					break;
				}
				case "message": {
					await this.updateMembers(TeamsInfo, context);

					const teamMember = this.members[17]; //Specific Member to interact with (Tyson:3,Ethan:17)

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
					const teamsChannelId = this.channels[1].id; //Specific Channel to interact with
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
	validateEmail(email) {
		const re = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
		return re.test(String(email).toLowerCase());
	}

	//search through all channels, match the names and return ID
	getChannelID(channelName) {
		for (var i = 0; i < this.channels.length; i++) {
			if (this.channels[i].name.toString() == channelName) {
				return this.channels[i].id;
			}
		}
		return null;
	}

	//search through all members, match email and return details
	getMemberDetails(email) {
		for (var i = 0; i < this.members.length; i++) {
			if (this.members[i].email.toString() == email) {
				return this.members[i];
			}
		}
		return null;
	}

	//send message to a channel
	async sendChannelMessage(adapter, channelID, message) {

		//conversation parameters
		const conversationParameters = {
			isGroup: true,
			channelData: {
				channel: {
					id: channelID
				}
			},
			activity: message
		};
		const connectorClient = adapter.createConnectorClient("https://smba.trafficmanager.net/ca/");
		const conversationResourceResponse = await connectorClient.conversations.createConversation(conversationParameters);
	}

	//send message to a member
	async sendMemberMessage(adapter, member, message) {

		const connectorClient = adapter.createConnectorClient("https://smba.trafficmanager.net/ca/");
		//conversation parameters
		const conversationParameters = {
			members: [
				member
			],
			channelData: {
				tenant: {
					id: "9f04f85a-8f3c-43e4-887b-549d66d6dab8"
				}
			}
		};
		const conversationResource = await connectorClient.conversations.createConversation(conversationParameters);
		await connectorClient.conversations.sendToConversation(conversationResource.id, message);
	}

	//create and send message based on info from POST request
	async draftMessage(adapter, body) {

		let message;
		switch (body.type) {
			//Number question
			case "Number": {

				if (body.min && body.max) {
					message = this.createNumberIcebreaker(body.text, body.id, body.min, body.max);
				} else {
					message = this.createNumberIcebreaker(body.text, body.id, 1, 10);

				}
				break;
			}
			//Yes/no question
			case "YesNo": {
				message = this.createTwoChoiceIcebreaker(body.text, ["yes", "no"], body.id);
				break;
			}
			//Happy/Sad question
			case "HappySad": {
				message = this.createTwoChoiceIcebreaker(body.text, [":)", ":("], body.id);
				break;
			}
			//TextBlock question
			case "TextBlock": {
				message = this.createTextBlockIcebreaker(body.text, body.id);
				break;
			}
			//Multiple choice question
			case "MultiChoice": {
				message = this.createMultiChoiceIcebreaker(body.text, body.choices, body.id);
				break;
			}
			//Text message
			case "Message": {
				message = MessageFactory.text(body.text);
				break;
			}
		}
		//loop through targets and check for message type
		(body.targets).forEach(element => {
			if (this.validateEmail(element)) {
				//email -> member
				this.sendMemberMessage(adapter, this.getMemberDetails(element), message);
			} else {
				//channelName -> channel
				this.sendChannelMessage(adapter, this.getChannelID(element), message);
			}
		});
	}
	/*
			Create an icebreaker question with multiple response options
		*/
	createNumberIcebreaker(question, id, min, max) {
		const card = CardFactory.adaptiveCard({

			"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
			"type": "AdaptiveCard",
			"version": "1.0",
			"body": [

				{
					"type": "TextBlock",
					"text": question
				}, {
					"type": "Input.Number",
					"id": "answer",
					"placeholder": "Enter a number",
					"min": min,
					"max": max,
					"value": 1
				}
			],
			"actions": [
				{
					"type": "Action.Submit",
					"title": "OK",
					"data": {
						id: id
					}
				}
			]
		}

		);

		return MessageFactory.attachment(card);

	}
	/*
			Create an icebreaker question with multiple response options
		*/
	createTextBlockIcebreaker(question, id) {
		const card = CardFactory.adaptiveCard({
			"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
			"type": "AdaptiveCard",
			"version": "1.0",
			"body": [

				{
					"type": "TextBlock",
					"text": question
				},
				{
					"type": "Input.Text",
					"id": "answer",
					"placeholder": "",
					"maxLength": 500,
					"isMultiline": true
				}
			],
			"actions": [
				{
					"type": "Action.Submit",
					"title": "Submit",
					"data": {
						id: id
					}
				}
			]
		});

		return MessageFactory.attachment(card);

	}
	/*
		Create an icebreaker question with multiple response options
	*/
	createMultiChoiceIcebreaker(question, choices, id) {
		var title = "IceBreaker";
		var description = question;
		const dateOptions = {
			weekday: 'long',
			year: 'numeric',
			month: 'long',
			day: 'numeric'
		};
		const dateStr = new Date().toLocaleDateString(undefined, dateOptions);

		const card = CardFactory.adaptiveCard({
			"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
			"type": "AdaptiveCard",
			"version": "1.0",
			"body": [{
				"type": "TextBlock",
				"size": "Medium",
				"weight": "Bolder",
				"text": `${title}`
			},
			{
				"type": "ColumnSet",
				"columns": [{
					"type": "Column",
					"items": [{
						"type": "Image",
						"style": "Person",
						"url": `https://avatars.slack-edge.com/2021-03-02/1820480857892_f5ff53aaec7a5507e5ad_512.png`,
						"size": "Small"
					}],
					"width": "auto"
				},
				{
					"type": "Column",
					"items": [{
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
			{
				"type": "Input.ChoiceSet",
				"id": "answer",
				"style": "expanded",
				"isMultiSelect": false,
				"value": "null",
				"choices": choices
			}
			],
			"actions": [{
				"type": "Action.Submit",
				"title": "Submit",
				"data": {
					id: id
				}
			}
			]
		});

		return MessageFactory.attachment(card);
	}

	/*
		Create an icebreaker question with only yes or no answers
	*/
	createTwoChoiceIcebreaker(question, choices, id) {
		var title = "IceBreaker";
		var description = question;
		const dateOptions = {
			weekday: 'long',
			year: 'numeric',
			month: 'long',
			day: 'numeric'
		};
		const dateStr = new Date().toLocaleDateString(undefined, dateOptions);

		const card = CardFactory.adaptiveCard({
			"$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
			"type": "AdaptiveCard",
			"version": "1.0",
			"body": [{
				"type": "TextBlock",
				"size": "Medium",
				"weight": "Bolder",
				"text": `${title}`
			},
			{
				"type": "ColumnSet",
				"columns": [{
					"type": "Column",
					"items": [{
						"type": "Image",
						"style": "Person",
						"url": `https://avatars.slack-edge.com/2021-03-02/1820480857892_f5ff53aaec7a5507e5ad_512.png`,
						"size": "Small"
					}],
					"width": "auto"
				},
				{
					"type": "Column",
					"items": [{
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
			"actions": [{
				"type": "Action.Submit",
				"title": choices[0],
				"data": {
					id: id,
					answer: choices[0]
				}
			},
			{
				"type": "Action.Submit",
				"title": choices[1],
				"data": {
					id: id,
					answer: choices[1]
				}
			}
			]
		});

		return MessageFactory.attachment(card);
	}
}

module.exports.TeamsBot = TeamsBot;