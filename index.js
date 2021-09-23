const express = require("express");
const ngrok = require('ngrok');
const { BotFrameworkAdapter } = require("botbuilder");


const { ActionMessageExtensionReply } = require("./ActionMessageExtensionReply.js");

const adapter = new BotFrameworkAdapter({
  appId: process.env.AI,
  appPassword: process.env.AP,
});

adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Create bot handlers
const mctionMessageExtensionReply = new ActionMessageExtensionReply();

// Create HTTP server.
const server = express();
const port = process.env.port || process.env.PORT || 3978;

server.listen(port, function () {
  console.log(`\Bot/ME service listening at http://localhost:${port}` )

  //code handle dev mode, exposes app in ngrok
  if (process.env.DEVMODE == 'true')
  {
    ngrok.connect(port, function(err, url) {
      console.log(`Node.js local server is publicly-accessible at ${url}`);
    });
  }
  console.log('%s listening to ', server.name );
});

// This is to test
server.get("/Ping", (req, res) => {
  res.send(`Ping Success`);
});


// Listen for incoming requests.
server.post("/api/messages", (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    // Process bot activity
    await mctionMessageExtensionReply.run(context);
  });
});

server.use(express.static(__dirname+ '/Static'))
