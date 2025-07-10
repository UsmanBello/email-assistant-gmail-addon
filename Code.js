function buildAddOn(e) {
    var cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Email Assistant"));

    var userProps = PropertiesService.getUserProperties();
    var jwtToken = userProps.getProperty('jwtToken');
    var email = Session.getActiveUser().getEmail();
    var name = '';
    try {
        // Try to get user's name from the Gmail message if available
        if (e && e.gmail && e.gmail.messageId) {
            var message = GmailApp.getMessageById(e.gmail.messageId);
            var from = message.getFrom();
            // Extract name from "Name <email>" format
            var match = from.match(/^(.*?)\s*<.*?>$/);
            if (match && match[1]) {
                name = match[1];
            }
        }
    } catch (err) {
        // Ignore errors, fallback to empty name
    }

    // Check if JWT exists and is valid (not expired)
    var isTokenValid = false;
    if (jwtToken) {
        try {
            var payload = JSON.parse(Utilities.newBlob(Utilities.base64Decode(jwtToken.split('.')[1])).getDataAsString());
            var now = Math.floor(Date.now() / 1000);
            if (payload.exp && payload.exp > now) {
                isTokenValid = true;
            } else {
                // Token expired, remove it
                userProps.deleteProperty('jwtToken');
                jwtToken = null;
            }
        } catch (err) {
            // Invalid token, remove it
            userProps.deleteProperty('jwtToken');
            jwtToken = null;
        }
    }

    if (!jwtToken || !isTokenValid) {
        // Not authenticated: show Login and Register buttons
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "To use Email Assistant, please log in or register."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Login")
                        .setOnClickAction(CardService.newAction().setFunctionName("onAddonLogin"))
                )
                .addWidget(
                    CardService.newTextButton()
                        .setText("Register")
                        .setOpenLink(CardService.newOpenLink().setUrl("https://fabulous-sundae-79255b.netlify.app/auth?source=addon"))
                )
        );
        return cardBuilder.build();
    }

    // Authenticated: proceed to org lookup and main features
    var orgInfo = getOrganizationInfo(email, jwtToken);
    if (!orgInfo || orgInfo.error) {
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "You need to register with Email Assistant before using this add-on."
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("Register Now")
                        .setOpenLink(CardService.newOpenLink().setUrl("https://fabulous-sundae-79255b.netlify.app/auth?source=addon"))
                )
        );
        return cardBuilder.build();
    }

    // Show organization name (and logo if available)
    var orgSection = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(
            '<b>Company:</b> ' + orgInfo.organizationName
        ));
    // TODO: Add logo when available
    cardBuilder.addSection(orgSection);

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

function onAddonLogin(e) {
    var userProps = PropertiesService.getUserProperties();
    var email = Session.getActiveUser().getEmail();
    Logger.log('Active user email: ' + email);
    if (!email) {
        Logger.log('No email found for active user.');
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Login failed: No Gmail user email found. Please ensure you are logged into Gmail."))
            )
            .build();
    }
    var name = '';
    try {
        if (e && e.gmail && e.gmail.messageId) {
            var message = GmailApp.getMessageById(e.gmail.messageId);
            var from = message.getFrom();
            var match = from.match(/^(.*?)\s*<.*?>$/);
            if (match && match[1]) {
                name = match[1];
            }
        }
    } catch (err) { }
    if (!name) {
        name = email; // fallback to email as name
    }
    Logger.log('Extracted name: ' + name);
    var authResult = authenticateAddonUser(email, name);
    if (authResult && authResult.token) {
        userProps.setProperty('jwtToken', authResult.token);
        Logger.log('Add-on login successful. JWT: ' + authResult.token);
        // Show a success message
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Login successful!"))
            )
            .build();
    } else {
        userProps.deleteProperty('jwtToken');
        var errorMsg = (authResult && authResult.error) ? authResult.error : "Login failed.";
        Logger.log('Add-on login failed: ' + errorMsg);
        // Show an error message
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Login failed: " + errorMsg))
            )
            .build();
    }
}

function authenticateAddonUser(email, name) {
    var ADDON_SECRET = process.env.ADDON_SECRET;
    try {
        var response = UrlFetchApp.fetch(
            "/api/users/auth/addon",
            {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify({ email: email, name: name }),
                muteHttpExceptions: true,
                headers: { 'X-Addon-Secret': ADDON_SECRET },
            }
        );
        var code = response.getResponseCode();
        if (code === 200) {
            return JSON.parse(response.getContentText());
        } else {
            Logger.log("Add-on auth failed for " + email + " - HTTP " + code + ": " + response.getContentText());
            try {
                return { error: JSON.parse(response.getContentText()).message };
            } catch (err) {
                return { error: "Unknown error" };
            }
        }
    } catch (err) {
        Logger.log("Exception in add-on auth for " + email + ": " + err);
        return { error: "Exception" };
    }
}

function getOrganizationInfo(email, token) {
    try {
        var response = UrlFetchApp.fetch(
            `${process.env.SERVER_DOMAIN}/api/organizations/by-user-email?email=${encodeURIComponent(email)}`,
            {
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + token },
            }
        );
        var code = response.getResponseCode();
        if (code === 200) {
            return JSON.parse(response.getContentText());
        } else {
            Logger.log("Org lookup failed for " + email + " - HTTP " + code + ": " + response.getContentText());
            try {
                return JSON.parse(response.getContentText());
            } catch (err) {
                return { error: "Unknown error" };
            }
        }
    } catch (err) {
        Logger.log("Exception in org lookup for " + email + ": " + err);
        return { error: "Exception" };
    }
}

function onGenerateAIReply(e) {
    var userProps = PropertiesService.getUserProperties();
    var jwtToken = userProps.getProperty('jwtToken');
    var email = Session.getActiveUser().getEmail();
    var orgInfo = null;
    if (jwtToken) {
        orgInfo = getOrganizationInfo(email, jwtToken);
    }
    if (!orgInfo || !orgInfo.organizationId) {
        Logger.log('AI Reply: Organization not found for user ' + email);
        return CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("Organization not found. Please register or contact your admin."))
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
    var aiReply = '';
    var errorMsg = '';
    try {
        // Show loading state
        var loadingCard = CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"))
            .addSection(
                CardService.newCardSection()
                    .addWidget(CardService.newTextParagraph().setText("🤖 Generating AI reply... Please wait."))
            )
            .build();

        var response = UrlFetchApp.fetch(
            `${process.env.SERVER_DOMAIN}/api/responses/generate-addon?organizationId=${orgInfo.organizationId}`,
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
                headers: { Authorization: "Bearer " + jwtToken },
            }
        );
        var code = response.getResponseCode();
        Logger.log('AI Reply: Backend response code: ' + code);
        if (code === 200) {
            var data = JSON.parse(response.getContentText());
            Logger.log('AI Reply: Backend response data: ' + JSON.stringify(data));
            if (data.responses && data.responses.length > 0) {
                // The backend returns an array of objects with 'content' property
                var firstResponse = data.responses[0];
                if (firstResponse && firstResponse.content) {
                    aiReply = firstResponse.content;
                } else if (typeof firstResponse === 'string') {
                    aiReply = firstResponse;
                } else {
                    aiReply = JSON.stringify(firstResponse);
                }
                Logger.log('AI Reply: Extracted reply: ' + aiReply.substring(0, 100) + '...');
            } else {
                errorMsg = "No AI reply generated.";
                Logger.log('AI Reply: No AI reply generated in response.');
            }
        } else {
            errorMsg = "AI reply generation failed: " + response.getContentText();
            Logger.log('AI Reply: Error response from backend: ' + response.getContentText());
        }
    } catch (err) {
        errorMsg = "Error generating AI reply: " + err;
        Logger.log('AI Reply: Exception during backend call: ' + err);
    }

    if (aiReply) {
        var cardBuilder = CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("Email Assistant"));

        // Show the first response
        var section = CardService.newCardSection()
            .addWidget(CardService.newTextParagraph().setText("<b>AI Suggested Reply:</b><br>" + aiReply));

        // Add "Use This Reply" button to populate Gmail's reply field
        section.addWidget(
            CardService.newTextButton()
                .setText("Use This Reply")
                .setOnClickAction(CardService.newAction().setFunctionName("onUseReply").setParameters({
                    replyText: aiReply
                }))
        );

        // If there are multiple responses, show them as options
        if (data.responses && data.responses.length > 1) {
            section.addWidget(CardService.newTextParagraph().setText("<br><b>Alternative Options:</b>"));
            for (var i = 1; i < Math.min(data.responses.length, 3); i++) {
                var altResponse = data.responses[i];
                var altContent = altResponse.content || altResponse;
                section.addWidget(CardService.newTextParagraph().setText("<br><b>Option " + (i + 1) + ":</b><br>" + altContent));

                // Add "Use This Reply" button for each alternative
                section.addWidget(
                    CardService.newTextButton()
                        .setText("Use Option " + (i + 1))
                        .setOnClickAction(CardService.newAction().setFunctionName("onUseReply").setParameters({
                            replyText: altContent
                        }))
                );
            }
        }

        cardBuilder.addSection(section);
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