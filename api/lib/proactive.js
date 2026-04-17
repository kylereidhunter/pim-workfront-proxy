// Proactive messaging — send a Teams message without the user initiating.
// Uses the same CloudAdapter as the inbound handler.

const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require('botbuilder');

function buildAdapter() {
  const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MICROSOFT_APP_ID,
    MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
    MicrosoftAppType: process.env.MICROSOFT_APP_TYPE || 'SingleTenant',
    MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
  });
  const auth = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
  return new CloudAdapter(auth);
}

async function sendProactive(conversationReference, text) {
  const adapter = buildAdapter();
  const appId = process.env.MICROSOFT_APP_ID;
  await adapter.continueConversationAsync(appId, conversationReference, async (context) => {
    await context.sendActivity(text);
  });
}

module.exports = { sendProactive };
