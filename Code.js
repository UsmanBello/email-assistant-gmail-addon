var SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
var SERVER_DOMAIN = SCRIPT_PROPERTIES.getProperty('SERVER_DOMAIN');
var ADDON_SECRET = SCRIPT_PROPERTIES.getProperty('ADDON_SECRET');

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
                        .setOpenLink(CardService.newOpenLink().setUrl("https://email-ai-assistant.netlify.app/register"))
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

    // Show user/organization info based on user type
    var infoSection = CardService.newCardSection();
    if (userInfo.userType === 'organization') {
        infoSection.addWidget(CardService.newTextParagraph().setText(
            '<b>Company:</b> ' + userInfo.organizationName
        ));
    } else {
        infoSection.addWidget(CardService.newTextParagraph().setText(
            '<b>User:</b> ' + userInfo.userName
        ));
    }
    // TODO: Add logo when available
    cardBuilder.addSection(infoSection);

    // Add settings section
    var settingsSection = CardService.newCardSection()
        .addWidget(
            CardService.newTextButton()
                .setText("⚙️ Settings")
                .setOnClickAction(CardService.newAction().setFunctionName("onShowSettings"))
        );
    cardBuilder.addSection(settingsSection);

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
                    '<b>Subject:</b> ' + subject + '<br>' +
                    '<b>From:</b> ' + from + '<br><br>' +
                    '<b>Body Preview:</b><br>' + snippet
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Generate AI Reply")
                        .setOnClickAction(CardService.newAction().setFunctionName("onGenerateAIReply"))
                )
        );
    } else {
        // Not in email context
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText("Welcome to the Email Assistant Gmail Add-on!"))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Generate AI Reply")
                        .setOnClickAction(CardService.newAction().setFunctionName("onGenerateAIReply"))
                )
        );
    }

    return cardBuilder.build();
}

function getUserInfo(email, idToken) {
    try {
        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/organizations/by-user-email?email=" + encodeURIComponent(email),
            {
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        Logger.log("User lookup response code: " + code + " for email: " + email);
        if (code === 200) {
            var responseData = JSON.parse(response.getContentText());
            Logger.log("User lookup successful response: " + JSON.stringify(responseData));
            return responseData;
        } else if (code === 404) {
            // User is authenticated but not registered in our system
            var errorData = JSON.parse(response.getContentText());
            Logger.log("User not registered in system: " + JSON.stringify(errorData));
            return { error: "User not registered", message: errorData.message || "User not found in Email Assistant database" };
        } else {
            Logger.log("User lookup failed for " + email + " - HTTP " + code + ": " + response.getContentText());
            try {
                return JSON.parse(response.getContentText());
            } catch (err) {
                return { error: "Unknown error" };
            }
        }
    } catch (err) {
        Logger.log("Exception in user lookup for " + email + ": " + err);
        return { error: "Exception" };
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
                            .setOpenLink(CardService.newOpenLink().setUrl("https://email-ai-assistant.netlify.app/register"))
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
                            .setOnClickAction(CardService.newAction().setFunctionName("onGenerateAIReply"))
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
            var responseSection = CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(replyText))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Use This Reply")
                        .setOnClickAction(CardService.newAction().setFunctionName("onUseReply").setParameters({
                            replyText: replyText
                        }))
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
                            .setOnClickAction(CardService.newAction().setFunctionName("onGenerateAIReply"))
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

    // Show the reply in a multi-line text input for easy copy-paste
    var inputField = CardService.newTextInput()
        .setFieldName("aiReplyText")
        .setTitle("AI Suggested Reply:")
        .setValue(replyText)
        .setMultiline(true);

    // Add a Copy to Clipboard button (note: Apps Script can't copy to clipboard directly, but user can select/copy easily)
    var instructions = CardService.newTextParagraph().setText(
        "<b>Instructions:</b> Select the text above, copy it (Ctrl+C or Cmd+C), and paste it into your Gmail reply field.");

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
        .addSection(
            CardService.newCardSection()
                .addWidget(inputField)
                .addWidget(instructions)
        )
        .build();
}

function onShowSettings(e) {
    var userProps = PropertiesService.getUserProperties();
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
                .addWidget(CardService.newTextParagraph().setText("<b>Actions:</b>"))
                .addWidget(
                    CardService.newTextButton()
                        .setText("🔄 Refresh Connection")
                        .setOnClickAction(CardService.newAction().setFunctionName("onRefreshConnection"))
                )
                .addWidget(
                    CardService.newTextButton()
                        .setText("🚪 Logout")
                        .setOnClickAction(CardService.newAction().setFunctionName("onLogout"))
                )
        )
        .build();
}

function onRefreshConnection(e) {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('jwtToken');

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText("✅ Connection refreshed! Please log in again."))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Login")
                        .setOnClickAction(CardService.newAction().setFunctionName("onAddonLogin"))
                )
        )
        .build();
}

function onLogout(e) {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('jwtToken');

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText("✅ Logged out successfully!"))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Login Again")
                        .setOnClickAction(CardService.newAction().setFunctionName("onAddonLogin"))
                )
        )
        .build();
} 