/**
 * Action handlers for the Email Assistant add-on.
 * These functions are triggered by button clicks and universal actions (manifest).
 */

// How much of a conversation to send. The backend independently keeps only the
// newest turns, so these bounds exist to keep the request payload sane.
var MAX_THREAD_MESSAGES = 15;
var MAX_MESSAGE_BODY_CHARS = 6000;

/**
 * Collects the whole conversation around the open message.
 *
 * The backend decides which message to reply to (the newest one not sent by this
 * user), so it needs the full thread — not just whichever message happens to be
 * open. Sending only the open message is what previously let the assistant reply
 * to the user's own words.
 *
 * Falls back to the single open message if the thread cannot be read, so a reply
 * is still drafted rather than the whole action failing.
 *
 * @returns {{messages: Array, usedThread: Boolean}}
 */
function collectThreadMessages(message) {
    var toPayload = function (msg) {
        var body = '';
        try {
            body = msg.getPlainBody() || '';
        } catch (err) {
            Logger.log('collectThreadMessages: body read failed: ' + err);
        }
        if (body.length > MAX_MESSAGE_BODY_CHARS) {
            body = body.substring(0, MAX_MESSAGE_BODY_CHARS);
        }
        return {
            id: msg.getId(),
            from: msg.getFrom() || '',
            to: msg.getTo() || '',
            cc: msg.getCc() || '',
            subject: msg.getSubject() || '',
            date: msg.getDate() ? msg.getDate().toUTCString() : '',
            timestamp: msg.getDate() ? msg.getDate().getTime() : null,
            body: body
        };
    };

    try {
        var thread = message.getThread();
        var messages = thread.getMessages();

        // Keep the newest MAX_THREAD_MESSAGES; older context is the first to drop.
        if (messages.length > MAX_THREAD_MESSAGES) {
            messages = messages.slice(messages.length - MAX_THREAD_MESSAGES);
        }

        var payload = [];
        for (var i = 0; i < messages.length; i++) {
            payload.push(toPayload(messages[i]));
        }
        Logger.log('AI Reply: collected ' + payload.length + ' thread message(s)');
        return { messages: payload, usedThread: true };
    } catch (err) {
        // Most likely cause is the granted scopes not permitting a thread read.
        // Degrade to the open message instead of failing the whole action.
        Logger.log('AI Reply: thread read unavailable, using open message only: ' + err);
        try {
            return { messages: [toPayload(message)], usedThread: false };
        } catch (inner) {
            Logger.log('AI Reply: open message read also failed: ' + inner);
            return { messages: [], usedThread: false };
        }
    }
}

/**
 * Builds a simple informational card. Text is always code-supplied, never
 * interpolated from a backend response.
 */
function buildInfoCard(title, bodyText, retryFunctionName) {
    var section = CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(bodyText));

    if (retryFunctionName) {
        section.addWidget(
            CardService.newTextButton()
                .setText("Try Again")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName(retryFunctionName)
                    .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        );
    }

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle(title))
        .addSection(section)
        .build();
}

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
                    .addWidget(CardService.newTextParagraph().setText(
                        "Authentication error. Please try again." +
                        (userInfo && userInfo.message ? "<br><br><i>Details: " + userInfo.message + "</i>" : "")
                    ))
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

    // Optional adjustment typed into the results card's refine box ("make it
    // shorter"). Present only when this call is a Regenerate.
    var instruction = '';
    if (e && e.formInput && e.formInput.refineInstruction) {
        instruction = String(e.formInput.refineInstruction).substring(0, 1000);
    }

    // Extract email context. We send the whole conversation, not just the open
    // message — the backend picks the reply target from it.
    var subject = '', from = '', body = '', threadId = '', messageId = '';
    var threadMessages = [];
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

            var collected = collectThreadMessages(message);
            threadMessages = collected.messages;
        }
    } catch (err) {
        Logger.log('AI Reply: Error extracting email context: ' + err);
    }

    // Call backend to generate AI reply using the add-on specific endpoint
    var aiReplies = [];
    var errorMsg = '';
    var infoMsg = '';
    // The message the backend chose to answer — the newest one not sent by this
    // user. Drafts are threaded onto this, not onto whatever the user had open.
    var replyToMessageId = '';
    try {
        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/responses/generate-addon",
            {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify({
                    // Full conversation: the backend selects the reply target from it.
                    threadMessages: threadMessages,
                    // Retained so an older backend build still behaves as before.
                    emailContent: body,
                    subject: subject,
                    from: from,
                    threadId: threadId,
                    messageId: messageId,
                    // The user's own steering for this generation; empty on first run.
                    instruction: instruction
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

            if (data.replyTo && typeof data.replyTo.messageId === 'string') {
                replyToMessageId = data.replyTo.messageId;
            }

            if (data.status === 'nothing_to_reply_to') {
                // Everything in this conversation was sent by the user, so there is
                // no incoming message to answer. Not an error — explain it instead.
                infoMsg = "There's no message from anyone else in this conversation yet, so there's nothing to reply to.";
                Logger.log('AI Reply: no incoming message in thread.');
            } else if (data.responses && data.responses.length > 0) {
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

    if (infoMsg) {
        // Informational, not a failure — no "Try Again" button.
        return buildInfoCard("ReplAI - Email Assistant", infoMsg, null);
    }

    if (aiReplies.length > 0) {
        var cardBuilder = CardService.newCardBuilder()
            .setHeader(CardService.newCardHeader().setTitle("AI Suggested Replies"));

        // Quick links always come first, on every card.
        cardBuilder.addSection(buildShortcutsSection());

        for (var i = 0; i < aiReplies.length; i++) {
            var replyText = aiReplies[i];
            var composeAction = CardService.newAction()
                .setFunctionName("onUseReply")
                .setParameters({
                    replyText: replyText,
                    // Action parameters must be strings.
                    replyToMessageId: replyToMessageId || ''
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

        // Refine box: type an adjustment and regenerate. The typed value comes
        // back to onGenerateAIReply via e.formInput.refineInstruction.
        var refineSection = CardService.newCardSection()
            .setHeader("Not quite right?")
            .addWidget(
                CardService.newTextInput()
                    .setFieldName("refineInstruction")
                    .setTitle("What should change?")
                    .setHint("e.g. shorter, more formal, mention the refund")
                    .setMultiline(true)
                    .setValue(instruction || '')
            )
            .addWidget(
                CardService.newTextButton()
                    .setText("↻ Regenerate")
                    .setOnClickAction(CardService.newAction()
                        .setFunctionName("onGenerateAIReply")
                        .setLoadIndicator(CardService.LoadIndicator.SPINNER))
            );
        if (instruction) {
            refineSection.addWidget(CardService.newTextParagraph().setText(
                '<font color="#0a7ea4"><i>These replies were adjusted with your request above.</i></font>'
            ));
        }
        cardBuilder.addSection(refineSection);

        return cardBuilder.build();
    } else {
        return buildInfoCard(
            "ReplAI - Email Assistant",
            errorMsg || "AI reply generation failed. Please try again.",
            "onGenerateAIReply"
        );
    }
}

/**
 * Inserts the selected AI reply as a draft in Gmail.
 * Called when user clicks "Use This Reply".
 */
function onUseReply(e) {
    var replyText = e.parameters.replyText;

    if (!replyText) {
        return buildInfoCard("ReplAI - Email Assistant", "Error: No reply text provided.", null);
    }

    try {
        var accessToken = e.gmail.accessToken;
        GmailApp.setCurrentMessageAccessToken(accessToken);

        // Draft against the message the reply was actually written for — the
        // latest message from the other party — falling back to the open message.
        // Without this, opening an older message (or the user's own last message)
        // would thread the draft onto the wrong one.
        var targetId = e.parameters.replyToMessageId || '';
        var openId = e.gmail.messageId;

        if (!targetId && !openId) {
            var blankDraft = GmailApp.createDraft('', '', replyText);
            return CardService.newComposeActionResponseBuilder()
                .setGmailDraft(blankDraft)
                .build();
        }

        var message = null;
        if (targetId) {
            try {
                message = GmailApp.getMessageById(targetId);
            } catch (targetErr) {
                Logger.log('onUseReply: could not open target message, falling back: ' + targetErr);
            }
        }
        if (!message && openId) {
            message = GmailApp.getMessageById(openId);
        }

        // Target unreadable and no open message to fall back to: still hand the
        // user their text as a plain draft rather than losing it.
        if (!message) {
            var looseDraft = GmailApp.createDraft('', '', replyText);
            return CardService.newComposeActionResponseBuilder()
                .setGmailDraft(looseDraft)
                .build();
        }

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
