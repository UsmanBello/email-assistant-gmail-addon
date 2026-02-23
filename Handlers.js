/**
 * Action handlers for the Email Assistant add-on.
 * These functions are triggered by button clicks and universal actions (manifest).
 */

/**
 * Generates AI reply suggestions via the backend API.
 * Called when user clicks "Generate AI Reply".
 */
function onGenerateAIReply(e) {
    var email = Session.getActiveUser().getEmail();
    var idToken = ScriptApp.getIdentityToken();
    Logger.log('Google ID Token (for backend verification): ' + idToken);
    var userInfo = null;
    userInfo = getUserInfo(email, idToken);
    Logger.log('AI Reply: User info received: ' + JSON.stringify(userInfo));

    // Check if user is not registered in our system
    if (userInfo && userInfo.error === "User not registered") {
        Logger.log('AI Reply: User not registered in Email Assistant system');
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Please register with Email Assistant first to use AI reply generation."))
                    .addWidget(
                        CardService.newTextButton()
                            .setText("Register Now")
                            .setOpenLink(CardService.newOpenLink().setUrl(FRONTEND_URL + '/register'))
                    )
            )
            .build();
    }

    // Check for other errors (authentication failures, etc.)
    if (!userInfo || userInfo.error || (!userInfo.organizationId && !userInfo.userId)) {
        Logger.log('AI Reply: User authentication failed - userInfo: ' + JSON.stringify(userInfo));
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Authentication error. Please try again."))
                    .addWidget(
                        CardService.newTextButton()
                            .setText("Try Again")
                            .setOnClickAction(CardService.newAction()
                                .setFunctionName("onGenerateAIReply")
                                .setLoadIndicator(CardService.LoadIndicator.SPINNER))
                    )
            )
            .build();
    }

    // Extract email context
    var subject = '', from = '', body = '', threadId = '', messageId = '';
    try {
        if (e && e.gmail && e.gmail.messageId) {
            messageId = e.gmail.messageId;
            threadId = e.gmail.threadId || '';
            var message = GmailApp.getMessageById(messageId);
            subject = message.getSubject();
            from = message.getFrom();
            body = message.getPlainBody();
        }
    } catch (err) {
        Logger.log('AI Reply: Error extracting email context: ' + err);
    }

    // Call backend to generate AI reply using the add-on specific endpoint
    var aiReplies = [];
    var errorMsg = '';
    try {
        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/responses/generate-addon",
            {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify({
                    emailContent: body,
                    subject: subject,
                    from: from,
                    threadId: threadId,
                    messageId: messageId
                }),
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        Logger.log('AI Reply: Backend response code: ' + code);
        if (code === 200) {
            var data = JSON.parse(response.getContentText());
            Logger.log('AI Reply: Backend response data: ' + JSON.stringify(data));
            if (data.responses && data.responses.length > 0) {
                aiReplies = data.responses.map(function (r) { return r.content || r; });
            } else {
                errorMsg = "No AI reply generated.";
                Logger.log('AI Reply: No AI reply generated in response.');
            }
        } else {
            var errorResponse = response.getContentText();
            Logger.log('AI Reply: Error response from backend: ' + errorResponse);

            try {
                var errorData = JSON.parse(errorResponse);
                if (errorData.error && errorData.error.includes('429')) {
                    errorMsg = "OpenAI quota exceeded. Please check your billing or try again later.";
                } else if (errorData.message) {
                    errorMsg = "AI reply generation failed: " + errorData.message;
                } else {
                    errorMsg = "AI reply generation failed. Please try again.";
                }
            } catch (parseError) {
                errorMsg = "AI reply generation failed. Please try again.";
            }
        }
    } catch (err) {
        errorMsg = "Error generating AI reply: " + err;
        Logger.log('AI Reply: Exception during backend call: ' + err);
    }

    if (aiReplies.length > 0) {
        var cardBuilder = CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("AI Suggested Replies"));

        for (var i = 0; i < aiReplies.length; i++) {
            var replyText = aiReplies[i];
            var composeAction = CardService.newAction()
                .setFunctionName("onUseReply")
                .setParameters({
                    replyText: replyText
                });

            var responseSection = CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(replyText))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Use This Reply")
                        .setComposeAction(composeAction, CardService.ComposedEmailType.REPLY_AS_DRAFT)
                );
            cardBuilder.addSection(responseSection);
        }

        return cardBuilder.build();
    } else {
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText(errorMsg || "AI reply generation failed."))
                    .addWidget(
                        CardService.newTextButton()
                            .setText("Try Again")
                            .setOnClickAction(CardService.newAction()
                                .setFunctionName("onGenerateAIReply")
                                .setLoadIndicator(CardService.LoadIndicator.SPINNER))
                    )
            )
            .build();
    }
}

/**
 * Inserts the selected AI reply as a draft in Gmail.
 * Called when user clicks "Use This Reply".
 */
function onUseReply(e) {
    var replyText = e.parameters.replyText;

    if (!replyText) {
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Error: No reply text provided."))
            )
            .build();
    }

    try {
        var accessToken = e.gmail.accessToken;
        GmailApp.setCurrentMessageAccessToken(accessToken);

        var messageId = e.gmail.messageId;
        if (!messageId) {
            var draft = GmailApp.createDraft('', '', replyText);
            return CardService.newComposeActionResponseBuilder()
                .setGmailDraft(draft)
                .build();
        }

        var message = GmailApp.getMessageById(messageId);
        var draft = message.createDraftReply(replyText);

        return CardService.newComposeActionResponseBuilder()
            .setGmailDraft(draft)
            .build();

    } catch (err) {
        Logger.log('Error in onUseReply: ' + err.toString());
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText(
                        "Error creating reply draft: " + err.toString()
                    ))
            )
            .build();
    }
}

/**
 * Shows add-on settings. Called from contextual trigger.
 */
function onShowSettings(e) {
    var email = Session.getActiveUser().getEmail();

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant Settings"))
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText("<b>User Information:</b>"))
                .addWidget(CardService.newTextParagraph().setText("Email: " + email))
                .addWidget(CardService.newTextParagraph().setText("Status: ✅ Connected"))
        )
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    CardService.newTextButton()
                        .setText("⚙️ Open Full Settings")
                        .setOpenLink(CardService.newOpenLink()
                            .setUrl(FRONTEND_URL + '/settings')
                            .setOpenAs(CardService.OpenAs.FULL_SIZE))
                )
        )
        .build();
}

/**
 * Universal action: Sign out. Called from the add-on's "More actions" menu.
 */
function onUniversalSignOut(e) {
    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("About Sign Out"))
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "Email Assistant stays connected through your Google account, there is no traditional sign out. If you'd like to fully disconnect, you can uninstall the add-on or remove its permissions from your Google Account."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Manage Google Permissions")
                        .setOpenLink(CardService.newOpenLink().setUrl("https://myaccount.google.com/permissions"))
                )
        )
        .build();
}
