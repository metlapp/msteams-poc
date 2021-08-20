TeamsInfo = require("botbuilder");
// index.js is used to setup and configure your bot
const serviceURL = "https://smba.trafficmanager.net/ca/";
// Import required packages
const restify = require("restify");
const path = require("path");
const {
	TurnContext,
	MessageFactory,
	calculateChangeHash
} = require("botbuilder");
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
	BotFrameworkAdapter,
	ConversationState,
	MemoryStorage,
	UserState
} = require("botbuilder");

const {
	TeamsBot
} = require("./teamsBot");
const {
	connect
} = require("ngrok");

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
	appId: process.env.BOT_ID,
	appPassword: process.env.BOT_PASSWORD,
});

adapter.onTurnError = async (context, error) => {
	// This check writes out errors to console log .vs. app insights.
	// NOTE: In production environment, you should consider logging this to Azure
	//       application insights. See https://aka.ms/bottelemetry for telemetry
	//       configuration instructions.
	console.error(`\n [onTurnError] unhandled error: ${error}`);

	// Send a trace activity, which will be displayed in Bot Framework Emulator
	await context.sendTraceActivity(
		"OnTurnError Trace",
		`${error}`,
		"https://www.botframework.com/schemas/error",
		"TurnError"
	);

	//TODO remove this and add a logging system
	//await context.sendActivity(`The bot encountered an unhandled error:\n ${error.message}`);
	//await context.sendActivity("To continue to run this bot, please fix the bot source code.");
	console.error(`The bot encountered an unhandled error:\n ${error.message}`);

	// Clear out state
	await conversationState.delete(context);
};

// Define the state store for your bot.
// See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state storage system to persist the dialog and user state between messages.
const memoryStorage = new MemoryStorage();

// For a distributed bot in production,
// this requires a distributed storage to ensure only one token exchange is processed.
const dedupMemory = new MemoryStorage();

// Create conversation and user state with in-memory storage provider.
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the bot that will handle incoming messages.
const bot = new TeamsBot(conversationState, userState);
const serverURL = "https://smba.trafficmanager.net/ca/";
// Create HTTP server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
	console.log(server.url);
	console.log(`\nBot started, ${server.name} listening to ${server.url}`);
});
server.use(restify.plugins.acceptParser(server.acceptable));
server.use(restify.plugins.queryParser());
server.use(
	restify.plugins.bodyParser({
		mapParams: true
	})
);

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
	await adapter
		.processActivity(req, res, async (context) => {
			await bot.run(context);
		})
		.catch((err) => {
			// Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, sholdn't throw this error.
			if (!err.message.includes("412")) {
				throw err;
			}
		});
});

// Listen for incoming proactive message request.
server.post("/api/proactivemessage", async (req, res, next) => {
	try {
		await bot.draftMessage(adapter, req.body);
		res.status(200);
	} catch (err) {
		console.log(err);
		res.status(500);
	}

	res.send();
});

server.get(
	"/*",
	restify.plugins.serveStatic({
		directory: path.join(__dirname, "public"),
	})
);

// Gracefully shutdown HTTP server
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach((event) => {
	process.on(event, () => {
		server.close();
	});
});