require('dotenv').config();

const restify = require('restify');
const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication
} = require('botbuilder');

// Auth config for Azure Bot single-tenant setup.
// These values are read from environment variables:
// MicrosoftAppType
// MicrosoftAppId
// MicrosoftAppPassword
// MicrosoftAppTenantId
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Global error handler
adapter.onTurnError = async (context, error) => {
    console.error('Bot error:', error);

    try {
        await context.sendActivity('Sorry, something went wrong while processing your request.');
    } catch (sendError) {
        console.error('Failed to send error activity:', sendError);
    }
};

// Create web server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Health check route
server.get('/', async (req, res) => {
    res.send({ status: 'Bot is running' });
});

// Start server
const port = process.env.PORT || 3978;
server.listen(port, () => {
    console.log(`Bot is running on port ${port}`);
});

// Azure Bot messaging endpoint
server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, async (context) => {
        if (context.activity.type !== 'message') {
            return;
        }

        const userMessage = context.activity.text || '';
        console.log('User message:', userMessage);

        try {
            // Call Oracle APEX ORDS REST endpoint
            const apexResponse = await fetch(process.env.APEX_ORDS_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                    // If later you secure APEX with a token, add:
                    // 'Authorization': `Bearer ${process.env.APEX_API_TOKEN}`
                },
                body: JSON.stringify({
                    messageText: userMessage
                })
            });

            if (!apexResponse.ok) {
                const errorText = await apexResponse.text();
                console.error('APEX API error:', apexResponse.status, errorText);

                await context.sendActivity(
                    `Oracle APEX API call failed. HTTP ${apexResponse.status}`
                );
                return;
            }

            const apexData = await apexResponse.json();
            console.log('APEX response:', apexData);

            await context.sendActivity(
                apexData.reply || 'No reply returned from Oracle APEX.'
            );
        } catch (err) {
            console.error('Error calling Oracle APEX:', err);
            await context.sendActivity('Unable to connect to Oracle APEX right now.');
        }
    });
});
