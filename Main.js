/**
 * Main add-on entry point.
 * buildAddOn is the homepage/contextual trigger – called when the add-on loads.
 */
function buildAddOn(e) {
    var cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"));

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

    var userInfo = getUserInfo(email, idToken);
    Logger.log('buildAddOn: User info received - keys: ' + (userInfo ? Object.keys(userInfo).join(', ') : 'null'));

    // Check if user is not registered in our system
    if (userInfo && userInfo.error === "User not registered") {
        Logger.log('buildAddOn: User not registered in ReplAI - Email Assistant system');
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "Welcome! You need to register with ReplAI - Email Assistant before using this add-on."
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
        Logger.log('buildAddOn: User authentication failed - keys: ' + (userInfo ? Object.keys(userInfo).join(', ') : 'null'));
        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    "Authentication error. Please try refreshing the addon." +
                    (userInfo && userInfo.message ? "<br><br><i>Details: " + userInfo.message + "</i>" : "")
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
        if (e.gmail.accessToken) {
            GmailApp.setCurrentMessageAccessToken(e.gmail.accessToken);
        }
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

    cardBuilder.addSection(buildShortcutsSection());

    return cardBuilder.build();
}

/**
 * Quick links into the ReplAI back office (knowledge base + AI response settings).
 * URLs must stay within the manifest's openLinkUrlPrefixes.
 */
function buildShortcutsSection() {
    return CardService.newCardSection()
        .setHeader("Quick links")
        .addWidget(
            CardService.newButtonSet()
                .addButton(
                    CardService.newTextButton()
                        .setText("📚 Knowledge Base")
                        .setOpenLink(CardService.newOpenLink()
                            .setUrl(FRONTEND_URL + '/knowledge-base')
                            .setOpenAs(CardService.OpenAs.FULL_SIZE))
                )
                .addButton(
                    CardService.newTextButton()
                        .setText("⚙️ AI Response Settings")
                        .setOpenLink(CardService.newOpenLink()
                            .setUrl(FRONTEND_URL + '/ai-response-settings')
                            .setOpenAs(CardService.OpenAs.FULL_SIZE))
                )
        );
}
