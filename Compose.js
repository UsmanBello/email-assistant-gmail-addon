/**
 * Compose-from-scratch flow.
 * The user describes what the email should say; the backend returns polished
 * drafts; "Use This Draft" saves a Gmail draft (gmail.compose scope).
 */

/**
 * The compose form. `opts` carries values to prefill and an optional error line.
 */
function buildComposeCard(opts) {
    opts = opts || {};

    var formSection = CardService.newCardSection()
        .setHeader("Compose New Email");

    if (opts.error) {
        formSection.addWidget(CardService.newTextParagraph().setText(
            '<font color="#c0392b"><b>' + opts.error + '</b></font>'
        ));
    }

    formSection
        .addWidget(
            CardService.newTextInput()
                .setFieldName("composeTo")
                .setTitle("To (optional)")
                .setHint("recipient@example.com")
                .setValue(opts.to || '')
        )
        .addWidget(
            CardService.newTextInput()
                .setFieldName("composeSubject")
                .setTitle("Subject (optional)")
                .setValue(opts.subject || '')
        )
        .addWidget(
            CardService.newTextInput()
                .setFieldName("composeIntent")
                .setTitle("What do you want to say?")
                .setHint("e.g. ask Anna for the Q3 invoices before Friday, friendly but firm")
                .setMultiline(true)
                .setValue(opts.intent || '')
        )
        .addWidget(
            CardService.newTextButton()
                .setText("🤖 Generate Draft")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("onGenerateCompose")
                    .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        );

    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("ReplAI - Email Assistant"))
        .addSection(buildShortcutsSection())
        .addSection(formSection)
        .build();
}

function onShowComposeForm(e) {
    return buildComposeCard({});
}

/**
 * Shown when the user picks "Reply to an Email" on the homepage: replying
 * always starts from an open message, so this card just explains that.
 */
function onShowReplyHelp(e) {
    return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Reply to an Email"))
        .addSection(buildShortcutsSection())
        .addSection(
            CardService.newCardSection()
                .addWidget(CardService.newTextParagraph().setText(
                    '<font color="#333">Open the <b>email you want to reply to</b> from your inbox — this add-on will follow along and show the reply options for it.</font>'
                ))
                .addWidget(CardService.newTextParagraph().setText(
                    '<table width="100%" cellpadding="12" cellspacing="0" style="background-color: #e8f4fd; border-radius: 6px; border: 1px solid #b3ddf2;">' +
                    '<tr><td>' +
                    '<font color="#0a7ea4"><b>💡 Tip:</b> <i>Click on any email in your inbox, then open this add-on to see the AI reply options.</i></font>' +
                    '</td></tr>' +
                    '</table>'
                ))
        )
        .build();
}

/**
 * Generates draft variations for the compose form.
 * Also handles Regenerate from the results card: the request fields live on
 * that card too, so e.formInput carries them plus the refine instruction.
 */
function onGenerateCompose(e) {
    var to = (e && e.formInput && e.formInput.composeTo) ? String(e.formInput.composeTo) : '';
    var subject = (e && e.formInput && e.formInput.composeSubject) ? String(e.formInput.composeSubject) : '';
    var intent = (e && e.formInput && e.formInput.composeIntent) ? String(e.formInput.composeIntent) : '';
    var instruction = (e && e.formInput && e.formInput.refineInstruction)
        ? String(e.formInput.refineInstruction).substring(0, 1000)
        : '';

    if (!intent.trim()) {
        return buildComposeCard({
            to: to,
            subject: subject,
            error: "Please describe what the email should say."
        });
    }
    intent = intent.substring(0, 4000);

    var email = Session.getActiveUser().getEmail();
    var idToken = ScriptApp.getIdentityToken();
    if (!email || !idToken) {
        return buildInfoCard(
            "ReplAI - Email Assistant",
            "⚠️ Unable to verify your Google account identity. Please try again or contact support.",
            "onShowComposeForm"
        );
    }

    var drafts = [];
    var errorMsg = '';
    try {
        var response = UrlFetchApp.fetch(
            SERVER_DOMAIN + "/api/responses/compose-addon",
            {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify({
                    intent: intent,
                    to: to,
                    subject: subject,
                    instruction: instruction
                }),
                muteHttpExceptions: true,
                headers: { Authorization: "Bearer " + idToken },
            }
        );
        var code = response.getResponseCode();
        Logger.log('Compose: Backend response code: ' + code);
        if (code === 200) {
            var data = JSON.parse(response.getContentText());
            if (data.responses && data.responses.length > 0) {
                drafts = data.responses.map(function (r) { return r.content || r; });
            } else {
                errorMsg = "No draft generated. Please try again.";
            }
        } else if (code === 401 || code === 403) {
            errorMsg = "You need to sign in again to generate drafts.";
        } else if (code === 402) {
            errorMsg = "Your free trial has ended. Add billing at replai.us to keep generating drafts.";
        } else if (code === 404) {
            // 404 is ambiguous: the backend returns it for an unregistered user
            // (JSON with error field), but Express also 404s (HTML) when the
            // route itself doesn't exist yet — e.g. the server hasn't been
            // deployed with compose support.
            var isUnregistered = false;
            try {
                var errData = JSON.parse(response.getContentText());
                isUnregistered = Boolean(errData && errData.error === "User not registered");
            } catch (ignored) { }
            errorMsg = isUnregistered
                ? "Please register with ReplAI - Email Assistant first at replai.us."
                : "The drafting service isn't available yet. Please try again in a few minutes.";
        } else if (code === 429) {
            errorMsg = "Quota reached. Please try again later.";
        } else if (code >= 500) {
            errorMsg = "The drafting service is temporarily unavailable. Please try again.";
        } else {
            errorMsg = "Draft generation failed. Please try again.";
        }
    } catch (err) {
        errorMsg = "Draft generation failed. Please try again.";
        Logger.log('Compose: Exception during backend call: ' + err);
    }

    if (errorMsg) {
        return buildComposeCard({ to: to, subject: subject, intent: intent, error: errorMsg });
    }

    var cardBuilder = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("AI Suggested Drafts"));

    // Quick links always come first, on every card.
    cardBuilder.addSection(buildShortcutsSection());

    // The request stays on the card (collapsed) so the user can tweak it and
    // Regenerate — the fields are read back via e.formInput.
    cardBuilder.addSection(
        CardService.newCardSection()
            .setHeader("Your request")
            .setCollapsible(true)
            .setNumUncollapsibleWidgets(0)
            .addWidget(
                CardService.newTextInput()
                    .setFieldName("composeTo")
                    .setTitle("To (optional)")
                    .setValue(to)
            )
            .addWidget(
                CardService.newTextInput()
                    .setFieldName("composeSubject")
                    .setTitle("Subject (optional)")
                    .setValue(subject)
            )
            .addWidget(
                CardService.newTextInput()
                    .setFieldName("composeIntent")
                    .setTitle("What do you want to say?")
                    .setMultiline(true)
                    .setValue(intent)
            )
    );

    for (var i = 0; i < drafts.length; i++) {
        // A compose action (not a plain click action): Gmail opens the draft in
        // a compose window, exactly like "Use This Reply" does for replies.
        var composeAction = CardService.newAction()
            .setFunctionName("onUseComposeDraft")
            .setParameters({
                draftText: drafts[i],
                draftTo: to,
                draftSubject: subject
            });
        var draftSection = CardService.newCardSection()
            .addWidget(CardService.newTextParagraph().setText(drafts[i]))
            .addWidget(
                CardService.newTextButton()
                    .setText("Use This Draft")
                    .setComposeAction(composeAction, CardService.ComposedEmailType.STANDALONE_DRAFT)
            );
        // First draft always fully visible; later ones collapse to a preview
        // header (cards can't be collapsible AND initially open).
        if (i > 0) {
            draftSection
                .setHeader('Draft ' + (i + 1) + ': "' + previewSnippet(drafts[i]) + '"')
                .setCollapsible(true)
                .setNumUncollapsibleWidgets(0);
        }
        cardBuilder.addSection(draftSection);
    }

    var refineSection = CardService.newCardSection()
        .setHeader("Not quite right?")
        .addWidget(
            CardService.newTextInput()
                .setFieldName("refineInstruction")
                .setTitle("What should change?")
                .setHint("e.g. shorter, more formal, add a deadline")
                .setMultiline(true)
                .setValue(instruction)
        )
        .addWidget(
            CardService.newTextButton()
                .setText("↻ Regenerate")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("onGenerateCompose")
                    .setLoadIndicator(CardService.LoadIndicator.SPINNER))
        );
    if (instruction) {
        refineSection.addWidget(CardService.newTextParagraph().setText(
            '<font color="#0a7ea4"><i>These drafts were adjusted with your request above.</i></font>'
        ));
    } else {
        // Optional feature: collapsed to its header on a first generation.
        // Stays fully visible after a regenerate so the instruction is in view.
        refineSection.setCollapsible(true).setNumUncollapsibleWidgets(0);
    }
    cardBuilder.addSection(refineSection);

    return cardBuilder.build();
}

/**
 * Compose-action callback for "Use This Draft": creates the standalone Gmail
 * draft and hands it back, so Gmail opens it in a compose window where the
 * user can adjust the recipient, add attachments, and send.
 */
function onUseComposeDraft(e) {
    var draftText = e && e.parameters ? e.parameters.draftText : '';
    var to = e && e.parameters ? (e.parameters.draftTo || '') : '';
    var subject = e && e.parameters ? (e.parameters.draftSubject || '') : '';

    if (!draftText) {
        return buildInfoCard("ReplAI - Email Assistant", "Error: No draft text provided.", "onShowComposeForm");
    }

    try {
        var draft = GmailApp.createDraft(to, subject, draftText);
        return CardService.newComposeActionResponseBuilder()
            .setGmailDraft(draft)
            .build();
    } catch (err) {
        Logger.log('Compose: createDraft failed: ' + err);
        return buildInfoCard(
            "ReplAI - Email Assistant",
            "Couldn't open the draft. Please try again.",
            "onShowComposeForm"
        );
    }
}
