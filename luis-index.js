// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const dotenv = require('dotenv');
const path = require('path');
const restify = require('restify');
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
    BotFrameworkAdapter,
    MemoryStorage,
    ConversationState,
    UserState
} = require('botbuilder');

// Import required bot configuration.
const { BotConfiguration } = require('botframework-config');

// This bot's main dialog.
const { MyBot } = require('./luis');

// Read botFilePath and botFileSecret from .env file
// Note: Ensure you have a .env file and include botFilePath and botFileSecret.
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

// bot endpoint name as defined in .bot file
// See https://aka.ms/about-bot-file to learn more about .bot file its use and bot configuration.
const DEV_ENVIRONMENT = 'development';

// bot name as defined in .bot file
// See https://aka.ms/about-bot-file to learn more about .bot file its use and bot configuration.
const BOT_CONFIGURATION = process.env.NODE_ENV || DEV_ENVIRONMENT;

// Create HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
});

// .bot file path
const BOT_FILE = path.join(__dirname, process.env.botFilePath || '');

// Read bot configuration from .bot file.
let botConfig;
try {
    botConfig = BotConfiguration.loadSync(BOT_FILE, process.env.botFileSecret);
} catch (err) {
    console.error(
        `\nError reading bot file. Please ensure you have valid botFilePath and botFileSecret set for your environment.`
    );
    console.error(
        `\n - The botFileSecret is available under appsettings for your Azure Bot Service bot.`
    );
    console.error(
        `\n - If you are running this bot locally, consider adding a .env file with botFilePath and botFileSecret.`
    );
    console.error(
        `\n - See https://aka.ms/about-bot-file to learn more about .bot file its use and bot configuration.\n\n`
    );
    process.exit();
}

// Get bot endpoint configuration by service name
const endpointConfig = botConfig.findServiceByNameOrId(BOT_CONFIGURATION);
// Language Understanding (LUIS) service name as defined in the .bot file.
const LUIS_CONFIGURATION = process.env.LuisServiceName;
const luisConfig = botConfig.findServiceByNameOrId(LUIS_CONFIGURATION);
// Map the contents to the required format for `LuisRecognizer`.
const luisApplication = {
    applicationId: luisConfig.appId,
    endpointKey: luisConfig.subscriptionKey || luisConfig.authoringKey,
    azureRegion: luisConfig.region
};

// Create configuration for LuisRecognizer's runtime behavior.
const luisPredictionOptions = {
    includeAllIntents: true,
    log: true,
    staging: false
};

// Create adapter.
const adapter = new BotFrameworkAdapter({
    appId: endpointConfig.appId || process.env.microsoftAppID,
    appPassword: endpointConfig.appPassword || process.env.microsoftAppPassword
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    console.error(`\n [onTurnError]: ${error}`);
    // Send a message to the user
    await context.sendActivity(`Oops. Something went wrong!`);
    // Clear out conversation state
    await conversationState.delete(context);
    // Clear out user state
    await userState.delete(context);
};

// Create conversation and user state with in-memory storage provider.
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the main dialog.
const myBot = new MyBot(luisApplication, luisPredictionOptions,conversationState,userState);

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    // console.log(req);
    // console.log(res);
    adapter.processActivity(req, res, async context => {
        // Route to main dialog.
        await myBot.onTurn(context);
    });
});
