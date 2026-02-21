var SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
var SERVER_DOMAIN = SCRIPT_PROPERTIES.getProperty('SERVER_DOMAIN');
var ADDON_SECRET = SCRIPT_PROPERTIES.getProperty('ADDON_SECRET');
var FRONTEND_URL = SCRIPT_PROPERTIES.getProperty('FRONTEND_URL') || 'https://email-ai-assistant.netlify.app';

function buildAddOn(e) {
    var cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant"));

    var email = Session.getActiveUser().getEmail();
    var idToken = ScriptApp.getIdentityToken();
    var userInfo = getUserInfo(email, idToken);
    Logger.log('buildAddOn: User info received: ' + JSON.stringify(userInfo));

    // Check if user is not registered in our system
    if (userInfo && userInfo.error === "User not registered") {
        Logger.log('buildAddOn: User not registered in Email Assistant system');
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "Welcome! You need to register with Email Assistant before using this add-on."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Register Now")
                        .setOpenLink(CardService.newOpenLink().setUrl(FRONTEND_URL + '/register'))
                )
        );
        return cardBuilder.build();
    }

    // Check for other errors (authentication failures, etc.)
    if (!userInfo || userInfo.error || (!userInfo.organizationId && !userInfo.userId)) {
        Logger.log('buildAddOn: User authentication failed - userInfo: ' + JSON.stringify(userInfo));
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "Authentication error. Please try refreshing the addon."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Try Again")
                        .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn"))
                )
        );
        return cardBuilder.build();
    }

    Logger.log('buildAddOn: User authentication successful - userType: ' + userInfo.userType);

    // Dynamic header: "Select Email" when no email chosen, "Selected Email" when viewing one
    var headerTitle = (e && e.gmail && e.gmail.messageId) ? "Selected Email" : "Select Email";
    cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(headerTitle));

    // If in email context, show email info and Generate button
    if (e && e.gmail && e.gmail.messageId) {
        var message = GmailApp.getMessageById(e.gmail.messageId);
        var subject = message.getSubject();
        var from = message.getFrom();
        var body = message.getPlainBody();
        var snippet = body.length > 300 ? body.substring(0, 300) + "..." : body;

        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    '<div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 16px;">' +
                    '<br><br>' +
                    '<b>Subject:</b> ' + subject +
                    '<br><br>' +
                    '<b>From:</b> ' + from +
                    '<br><br>' +
                    '<b>Preview:</b><br>' +
                    '<font color="#555" size="2">' + snippet + '</font>' +
                    '</div>'
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("🤖 Generate AI Reply")
                        .setOnClickAction(CardService.newAction()
                            .setFunctionName("onGenerateAIReply")
                            .setLoadIndicator(CardService.LoadIndicator.SPINNER))
                )
        );
    } else {
        // Not in email context - show instruction message
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    '<div style="text-align: center; margin-bottom: 50px;">' +
                    '<font color="#333">To get started, please <b>select an email message</b> from your inbox that you\'d like to generate an AI reply for.</font>' +
                    '</div>'
                ))
                .addWidget(CardService.newTextParagraph().setText(
                    '<table width="100%" cellpadding="12" cellspacing="0" style="background-color: #e8f4fd; border-radius: 6px; border: 1px solid #b3ddf2;">' +
                    '<tr><td>' +
                    '<font color="#0a7ea4"><b>💡 Tip:</b> <i>Click on any email in your inbox, then open this add-on to see the AI reply options.</i></font>' +
                    '</td></tr>' +
                    '</table>'
                ))
        );
    }

    return cardBuilder.build();
}

function getUserInfo(email, idToken) {
    try {
        if (!SERVER_DOMAIN) {
            Logger.log("ERROR: SERVER_DOMAIN is not configured in script properties");
            return { error: "Configuration error", message: "Server domain is not configured. Please check script properties." };
        }

        if (!idToken) {
            Logger.log("ERROR: Google ID token is missing");
            return { error: "Authentication error", message: "Failed to get Google ID token. Please try refreshing the add-on." };
        }

        Logger.log("User lookup request - Email: " + email + ", Server: " + SERVER_DOMAIN);
        Logger.log("ID Token preview: " + (idToken ? idToken.substring(0, 20) + "..." : "null"));

        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/organizations/by-user-email?email=" + encodeURIComponent(email),
            {
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        var responseText = response.getContentText();

        Logger.log("User lookup response code: " + code + " for email: " + email);
        Logger.log("Response text: " + responseText.substring(0, 200));

        if (code === 200) {
            var responseData = JSON.parse(responseText);
            Logger.log("User lookup successful response: " + JSON.stringify(responseData));
            return responseData;
        } else if (code === 404) {
            // User is authenticated but not registered in our system
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("User not registered in system: " + JSON.stringify(errorData));
                return { error: "User not registered", message: errorData.message || "User not found in Email Assistant database" };
            } catch (parseErr) {
                Logger.log("Error parsing 404 response: " + parseErr);
                return { error: "User not registered", message: "User not found in Email Assistant database" };
            }
        } else if (code === 401) {
            // Authentication failed
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("Authentication failed: " + JSON.stringify(errorData));
                return {
                    error: "Authentication failed",
                    message: errorData.message || errorData.error || "Invalid Google ID token. Please check server configuration."
                };
            } catch (parseErr) {
                Logger.log("Error parsing 401 response: " + parseErr);
                return { error: "Authentication failed", message: "Invalid Google ID token" };
            }
        } else if (code === 500) {
            // Server error
            try {
                var errorData = JSON.parse(responseText);
                Logger.log("Server error: " + JSON.stringify(errorData));
                return {
                    error: "Server error",
                    message: errorData.message || errorData.error || "Server configuration error. Please contact support."
                };
            } catch (parseErr) {
                Logger.log("Error parsing 500 response: " + parseErr);
                return { error: "Server error", message: "Internal server error" };
            }
        } else {
            Logger.log("User lookup failed for " + email + " - HTTP " + code + ": " + responseText);
            try {
                var errorData = JSON.parse(responseText);
                return {
                    error: "Request failed",
                    message: errorData.message || errorData.error || "Unexpected error (HTTP " + code + ")"
                };
            } catch (err) {
                return { error: "Request failed", message: "Unexpected error (HTTP " + code + ")" };
            }
        }
    } catch (err) {
        Logger.log("Exception in user lookup for " + email + ": " + err.toString());
        Logger.log("Exception stack: " + (err.stack || "No stack trace"));
        return { error: "Exception", message: "Error connecting to server: " + err.toString() };
    }
}

function onGenerateAIReply(e) {
    var email = Session.getActiveUser().getEmail();
    var idToken = ScriptApp.getIdentityToken();
    Logger.log('Google ID Token (for backend verification): ' + idToken);
    var userInfo = null;
    // Use the existing getUserInfo, but pass idToken instead of jwtToken
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

            // Parse error response to show user-friendly message
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

        // Remove unsupported HTML and use CardSection for each response
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
        // REQUIRED: Set the access token to access the message
        var accessToken = e.gmail.accessToken;
        GmailApp.setCurrentMessageAccessToken(accessToken);

        // Get the current message to reply to
        var messageId = e.gmail.messageId;
        if (!messageId) {
            // Fallback: create a new draft if no message context
            var draft = GmailApp.createDraft('', '', replyText);
            return CardService.newComposeActionResponseBuilder()
                .setGmailDraft(draft)
                .build();
        }

        var message = GmailApp.getMessageById(messageId);

        // Create a draft reply with the AI-generated text
        var draft = message.createDraftReply(replyText);

        // Return a ComposeActionResponse that opens this draft
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