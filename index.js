require('dotenv').config();
const restify = require('restify');
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication
} = require('botbuilder');

// Authentication config reads these env vars:
// MicrosoftAppType
// MicrosoftAppId
// MicrosoftAppPassword
// MicrosoftAppTenantId
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Error handling
adapter.onTurnError = async (context, error) => {
    console.error('Bot error:', error);
    await context.sendActivity('Sorry, something went wrong.');
};

// Create web server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Optional root route for health check

server.get('/', async (req, res) => {
    res.send({ status: 'Bot is running' });
});

// Start server
server.listen(process.env.PORT || 3978, () => {
    console.log(`Bot is running on port ${process.env.PORT || 3978}`);
});

// Bot endpoint
server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, async (context) => {
        if (context.activity.type === 'message') {
            const userMessage = context.activity.text || '';
            await context.sendActivity(`You said: ${userMessage}`);
        }
    });
});
