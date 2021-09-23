const { TeamsActivityHandler, CardFactory} = require("botbuilder");

class ActionMessageExtensionReply extends TeamsActivityHandler {

    // Message Extension Code
    // Action.
    handleTeamsMessagingExtensionSubmitAction(context, action) {
      switch (action.commandId) {
        case "Reply":
          return ActionReplyToMessage(context, action);  
        default:
          throw new Error("NotImplemented");
      }
    }
    // Message Extension Code
    // FetchTask.
    async handleTeamsMessagingExtensionFetchTask(context, action) {
      switch (action.commandId) {
        case 'Reply':
            return ReplyToMessage();
        default:
          throw new Error("NotImplemented");
        }
    }

}

function ReplyToMessage() {
      return {
        task: {
            type: 'continue',
            value: {
                width: 150,
                height: 50,
                title: 'Replying...',
                url: `${process.env.BASEURL}/Reply.html`
            }
        }
    };
}

//Prepare Hero card for message reply
async function ActionReplyToMessage(context, action) {
  let userName = "unknown";
    if (
      action.messagePayload &&
      action.messagePayload.from &&
      action.messagePayload.from.user &&
      action.messagePayload.from.user.displayName
    ) {
      userName = action.messagePayload.from.user.displayName;
    }

    const heroCard = CardFactory.heroCard(
      `${userName} `,
      action.messagePayload.body.content
      //images
    );

    const attachment = {
      contentType: heroCard.contentType,
      content: heroCard.content,
      preview: heroCard,
    };

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment],
      },
    };
}


  module.exports.ActionMessageExtensionReply = ActionMessageExtensionReply;

