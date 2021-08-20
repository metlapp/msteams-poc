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

const TeamsUtils = require("./teamsUtils");

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

		this.serviceURL = "https://smba.trafficmanager.net/ca/";

		//Handler for messages sent within teams
		this.onMessage(async (context, next) => {
			//Process the message that was sent
			await this.processMessage(context);

			// By calling next() you ensure that the next BotHandler is run.
			await next();
		});

		//Handler for when a new channel is created 
		this.onTeamsChannelCreatedEvent(async (channelInfo, TeamInfo, context, next) => {
			await TeamsUtils.updateChannels(context);
			await next();
		});

		//Handler for when a channel is deleted
		this.onTeamsChannelDeletedEvent(async (channelInfo, TeamInfo, context, next) => {
			await TeamsUtils.updateChannels(context);
			await next();
		});

		//Handler for when a channel is renamed
		this.onTeamsChannelRenamedEvent(async (channelInfo, TeamInfo, context, next) => {
			await TeamsUtils.updateChannels(context);
			await next();
		});

		//Handler for when a channel is restored
		this.onTeamsChannelRestoredEvent(async (channelInfo, TeamInfo, context, next) => {
			await TeamsUtils.updateChannels(context);
			await next();
		});

		//Handler for when a new user joins the team (this includes the bot joining)
		this.onMembersAdded(async (context, next) => {
			const membersAdded = context.activity.membersAdded;
			for (let cnt = 0; cnt < membersAdded.length; cnt++) {
				if (membersAdded[cnt].id) {
					try {
						//Try and register the team
						let [success, response] = await TeamsUtils.registerTeam(context);
						if (success) {
							await context.sendActivity(response);
						} else {
							console.error(response);
							await context.sendActivity("There was a problem installing the bot with Metl Solutions. Please get in contact with us.");
						}
					} catch (err) {
						console.error(err);
					}

					break;
				}
			}
			await next();
		});

		//Handler for when a member is removed (or left i think)
		this.onMembersRemoved(async (context, next) => {
			const membersRemoved = context.activity.membersRemoved;
			for (let cnt = 0; cnt < membersRemoved.length; cnt++) {
				if (membersRemoved[cnt].id) {
					await TeamsUtils.deactivateTeam(context);
					return;
				}
			}

			await TeamsUtils.updateMembers(context);
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

		//check for a non text message
		if (context.activity.textFormat !== TextFormatTypes.Plain) {
			//TODO implement proper post request

			console.log(`Value: ${context.activity.value.answer}`);
			console.log(`From: ${context.activity.from.id}`);
		} else {
			//Figure out which command was used
			console.log(`Requested command: ${text}`);
			switch (text) {
				case "awake": {
					await context.sendActivity("Hello! I'm awake!");
					break;
				}
				case "update-channels": {
					let [success, response] = await TeamsUtils.updateChannels(context);
					if (success) {
						await context.sendActivity(response);
					} else {
						await context.sendActivity("There was an error trying to update the channels for the organization. Please get in contact with us.");
						console.error(response);
					}
					break;
				}
				case "update-members": {
					let [success, response] = await TeamsUtils.updateMembers(context);
					if (success) {
						await context.sendActivity(response);
					} else {
						await context.sendActivity("There was an error trying to update the members for the organization. Please get in contact with us.");
						console.error(response);
					}
					break;
				}
				default: {
					console.log(`Unknown command: ${text}`);
					//Temp just to catch any unrecognized commands
				}
			}
		}
	}

	//Used to check if a target sent from the API is a user
	isUser(target) {
		return target.email != null;
	}

	//Send a message to a channel
	async sendChannelMessage(adapter, channel, message) {
		try {
			//conversation parameters
			const conversationParameters = {
				isGroup: true,
				channelData: {
					channel: {
						id: channel.id
					}
				},
				activity: message
			};
			const connectorClient = adapter.createConnectorClient(this.serviceURL);
			const conversationResourceResponse = await connectorClient.conversations.createConversation(conversationParameters);
		} catch (err) {
			console.log(err);
			//TODO somehow tell the front-end there was an issue? Can't without a context.
			//Maybe try storing the latest context for each team in an array then index that?
		}
	}

	//Send a message to a user
	async sendMemberMessage(adapter, member, message) {
		try {
			//conversation parameters
			const conversationParameters = {
				members: [
					member
				],
				channelData: {
					tenant: {
						id: member.tenantId
					}
				}
			};
			const connectorClient = adapter.createConnectorClient(this.serviceURL);
			const conversationResource = await connectorClient.conversations.createConversation(conversationParameters);
			await connectorClient.conversations.sendToConversation(conversationResource.id, message);
		} catch (err) {
			console.log(err);
			//TODO somehow tell the front-end there was an issue? Can't without a context.
			//Maybe try storing the latest context for each team in an array then index that?
		}
	}

	//Create and send a message based on flags from the post request
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
				message = this.createTwoChoiceIcebreaker(body.text, ["Yes", "No"], body.id);
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

		//loop through targets and check for target type
		(body.targets).forEach(target => {
			if (this.isUser(target)) {
				this.sendMemberMessage(adapter, target, message);
			} else {
				this.sendChannelMessage(adapter, target, message);
			}
		});
	}

	//Create an icebreaker question with a number only response
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

	//Create an icebreaker question with a text response
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

	//Create an icebreaker question with multiple options
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

	//Create an icebreaker question with two possible responses
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