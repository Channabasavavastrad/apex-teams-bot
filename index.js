require('dotenv').config();
const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');

// Create adapter using Azure Bot App ID + Secret
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Basic error handling
adapter.onTurnError = async (context, error) => {
    console.error('Bot error:', error);
    await context.sendActivity('Sorry, something went wrong.');
};

// Create web server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Start server
server.listen(process.env.PORT || 3978, () => {
    console.log(`Bot is running on port ${process.env.PORT || 3978}`);
});

// Bot endpoint
server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === 'message') {
            const userMessage = context.activity.text || '';
            await context.sendActivity(`You said: ${userMessage}`);
        }
    });
});
