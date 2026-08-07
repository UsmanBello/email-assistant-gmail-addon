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

    if (!email || !idToken) {
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText(
                        "⚠️ Unable to verify your Google account identity. This typically happens with custom domain (Google Workspace) accounts where the admin has restricted Apps Script access. Please ask your Google Workspace admin to allow Apps Script identity access, or contact support."
                    ))
                    .addWidget(
                        CardService.newTextButton()
                            .setText("Try Again")
                            .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn"))
                    )
            )
            .build();
    }

    Logger.log('Google ID Token (for backend verification) preview: ' + (idToken ? idToken.substring(0, 20) + '...' : 'null'));
    var userInfo = null;
    userInfo = getUserInfo(email, idToken);
    Logger.log('AI Reply: User info received - keys: ' + (userInfo ? Object.keys(userInfo).join(', ') : 'null'));

    // Check if user is not registered in our system
    if (userInfo && userInfo.error === "User not registered") {
        Logger.log('AI Reply: User not registered in ReplAI - Email Assistant system');
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Please register with ReplAI - Email Assistant first to use AI reply generation."))
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
        Logger.log('AI Reply: User authentication failed - keys: ' + (userInfo ? Object.keys(userInfo).join(', ') : 'null'));
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
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
            if (e.gmail.accessToken) {
                GmailApp.setCurrentMessageAccessToken(e.gmail.accessToken);
            }
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
            Logger.log('AI Reply: Backend response data keys: ' + (data ? Object.keys(data).join(', ') : 'null'));
            if (data.responses && data.responses.length > 0) {
                aiReplies = data.responses.map(function (r) { return r.content || r; });
            } else {
                errorMsg = "No AI reply generated.";
                Logger.log('AI Reply: No AI reply generated in response.');
            }
        } else {
            var errorResponse = response.getContentText();
            Logger.log('AI Reply: Error response from backend (preview): ' + (errorResponse ? errorResponse.substring(0, 100) + '...' : 'empty'));

            // Show only code-driven, generic messages keyed off HTTP status / typed
            // error codes. Never interpolate backend-controlled strings into the UI
            // — a poisoned upstream response could otherwise drop unexpected text
            // (or characters that confuse the card renderer) into the add-on UI.
            if (code === 401 || code === 403) {
                errorMsg = "You need to sign in again to generate replies.";
            } else if (code === 402) {
                errorMsg = "Your free trial has ended. Add billing at replai.us to keep generating replies.";
            } else if (code === 429) {
                errorMsg = "Reply quota reached. Please try again later.";
            } else if (code >= 500) {
                errorMsg = "AI reply service is temporarily unavailable. Please try again.";
            } else {
                errorMsg = "AI reply generation failed. Please try again.";
            }
        }
    } catch (err) {
        // Never show internal error strings in the UI (can leak details / be user-controlled).
        errorMsg = "AI reply generation failed. Please try again.";
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
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText(errorMsg || "AI reply generation failed. Please try again."))
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
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
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
        // Internal error details are logged for debugging only — never surfaced
        // to the user (avoids leaking Apps Script stack frames / file paths).
        Logger.log('Error in onUseReply: ' + err.toString());
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText(
                        "Could not create reply draft. Please try again."
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
        .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant Settings"))
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
 * Calls the backend to revoke the Google OAuth grant and clear stored tokens,
 * then surfaces the result to the user.
 */
function onUniversalSignOut(e) {
    var idToken = null;
    try { idToken = ScriptApp.getIdentityToken(); } catch (err) {
        Logger.log("Sign-out: getIdentityToken failed: " + err);
    }

    var result = revokeBackendSession(idToken);

    var statusText = result.ok
        ? "✅ You have been signed out. Your ReplAI session and Google authorization have been revoked."
        : "⚠️ We couldn't fully revoke your session right now. You can finish signing out by removing ReplAI's access from your Google Account.";

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Sign Out"))
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(statusText))
                .addWidget(CardService.newTextParagraph().setText(
                    "To fully disconnect ReplAI - Email Assistant from your Google Account, you can also uninstall the add-on or remove its permissions below."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Manage Google Permissions")
                        .setOpenLink(CardService.newOpenLink().setUrl("https://myaccount.google.com/permissions"))
                )
        )
        .build();
}
