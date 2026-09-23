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

    // A completely headerless card can render blank in Gmail, so the homepage
    // gets the app name instead of the old "Select Email" label.
    var headerTitle = (e && e.gmail && e.gmail.messageId) ? "Selected Email" : "ReplAI - Email Assistant";
    cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(headerTitle));

    // Quick links always come first, on every card.
    cardBuilder.addSection(buildShortcutsSection());

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
        if (!snippet || !snippet.trim()) {
            snippet = '<i>No message text</i>';
        }

        var previewHtml =
            '<div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 16px;">' +
            '<br><br>' +
            '<b>Subject:</b> ' + subject +
            '<br><br>' +
            '<b>From:</b> ' + from +
            '<br><br>' +
            '<b>Preview:</b><br>' +
            '<font color="#555" size="2">' + snippet + '</font>';

        // Show which attachments were detected — the same auto-selection that
        // will be read for the reply — so the user knows we saw them even when
        // the message body is empty.
        var detected = autoSelectAttachments(message);
        if (detected.length) {
            var detectedNames = [];
            for (var d = 0; d < detected.length; d++) {
                detectedNames.push(String(detected[d].file.getName() || 'unnamed file'));
            }
            previewHtml +=
                '<br><br><font color="#0a7ea4">📎 <b>Attachment' + (detected.length === 1 ? '' : 's') + ' detected:</b> ' +
                detectedNames.join(', ') +
                '<br><i>Will be read and used for the reply.</i></font>';
        }
        previewHtml += '</div>';

        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(previewHtml))
        );

        cardBuilder.addSection(
            CardService.newCardSection()
                .addWidget(
                    CardService.newTextButton()
                        .setText("🤖 Generate AI Reply")
                        .setOnClickAction(CardService.newAction()
                            .setFunctionName("onGenerateAIReply")
                            .setLoadIndicator(CardService.LoadIndicator.SPINNER))
                )
        );
    } else {
        // Not in email context - offer the two entry points.
        cardBuilder.addSection(
            CardService.newCardSection()
                .setHeader("What would you like to do?")
                .addWidget(
                    CardService.newTextButton()
                        .setText("✍️ Compose New Email")
                        .setOnClickAction(CardService.newAction()
                            .setFunctionName("onShowComposeForm"))
                )
                .addWidget(CardService.newTextParagraph().setText(
                    '<font color="#555" size="2">Describe what you want to say and get a polished, ready-to-send draft.</font>'
                ))
                .addWidget(
                    CardService.newTextButton()
                        .setText("↩️ Reply to an Email")
                        .setOnClickAction(CardService.newAction()
                            .setFunctionName("onShowReplyHelp"))
                )
                .addWidget(CardService.newTextParagraph().setText(
                    '<font color="#555" size="2">Open an email from your inbox and get AI reply suggestions for it.</font>'
                ))
        );
    }

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
