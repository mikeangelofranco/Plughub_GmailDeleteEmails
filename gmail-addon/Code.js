/**
 * Called whenever the add-on is opened from an email.
 * e is the event object containing message metadata.
 */
function buildAddOn(e) {
  return [createCleanerCard(e, null)];
}

/**
 * Optional: what to show on homepage (no email selected).
 */
function buildHomePage(e) {
  return [createCleanerCard(e, null)];
}

/**
 * Builds the UI card that mirrors the provided mockup.
 */
function createCleanerCard(e, context) {
  var state = getFormState(e);
  context = context || {};

  var scopeInput = CardService.newSelectionInput()
    .setFieldName("deleteScope")
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .addItem("Current thread only", "thread", state.deleteScope === "thread")
    .addItem("Inbox", "inbox", state.deleteScope === "inbox")
    .addItem("All mail", "all", state.deleteScope === "all");

  var starredInput = CardService.newSelectionInput()
    .setFieldName("saveStarred")
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .addItem("Keep starred emails safe", "starred", state.saveStarred);

  var subjectInput = CardService.newTextInput()
    .setFieldName("subjectContains")
    .setTitle("Delete emails where Subject contains")
    .setHint('e.g. "promo", "newsletter", "facebook"')
    .setValue(state.subjectContains);

  var trashInput = CardService.newSelectionInput()
    .setFieldName("includeTrash")
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .addItem("Include Trash", "includeTrash", state.includeTrash);

  var previewAction = CardService.newAction()
    .setFunctionName("onPreviewEmails");
  var deleteAction = CardService.newAction()
    .setFunctionName("onDeleteNow");

  var previewButton = CardService.newTextButton()
    .setText("Preview emails")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#DB4437")
    .setOnClickAction(previewAction);

  var deleteButton = CardService.newTextButton()
    .setText("Delete now")
    .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
    .setOnClickAction(deleteAction);

  var section = CardService.newCardSection()
    .addWidget(
      CardService.newDecoratedText()
        .setText("<b>PlugHub Gmail Cleaner</b>")
        .setWrapText(true)
    )
    .addWidget(
      CardService.newTextParagraph()
        .setText("Bulk delete emails in your account. Use with caution.")
    )
    .addWidget(
      CardService.newTextParagraph()
        .setText("<font color='#DB4437'>Deleted emails may be unrecoverable.</font>")
    )
    .addWidget(scopeInput)
    .addWidget(starredInput)
    .addWidget(subjectInput)
    .addWidget(
      CardService.newTextParagraph()
        .setText("Leave empty to match all subjects")
    )
    .addWidget(trashInput)
    .addWidget(
      CardService.newButtonSet()
        .addButton(previewButton)
        .addButton(deleteButton)
    )
    .addWidget(
      CardService.newTextParagraph()
        .setText("Preview is recommended before deleting.")
    );

  if (context.error || context.preview) {
    section.addWidget(CardService.newDivider());

    if (context.error) {
      section.addWidget(
        CardService.newTextParagraph()
          .setText("<font color='#DB4437'>" + context.error + "</font>")
      );
    } else {
      section.addWidget(
        CardService.newTextParagraph()
          .setText("<b>Preview</b>")
      );

      var totalMatched = (typeof context.preview.total === "number")
        ? context.preview.total
        : (context.preview.items ? context.preview.items.length : 0);
      section.addWidget(
        CardService.newTextParagraph()
          .setText("Total matched emails: " + totalMatched)
      );

      if (context.preview.items && context.preview.items.length) {
        context.preview.items.forEach(function (item) {
          section.addWidget(
            CardService.newKeyValue()
              .setTopLabel(item.from + " • " + item.date)
              .setContent(item.subject)
              .setBottomLabel(item.snippet)
          );
        });
      } else {
        section.addWidget(
          CardService.newTextParagraph()
            .setText("No emails matched the current filters.")
        );
      }

      section.addWidget(
        CardService.newTextParagraph()
          .setText(context.preview.summary)
      );
    }
  }

  section
    .addWidget(
      CardService.newTextParagraph()
        .setText("<b>PlugHub</b>")
    );

  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Plughub Email Cleaner")
    )
    .addSection(section)
    .build();
}

/**
 * Optional: called when new authorization is needed.
 */
function onAuthorize(e) {
  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Authorization Required")
    )
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText("Please grant the requested permissions to use Plughub Email Cleaner.")
        )
    )
    .build();
}

/**
 * Placeholder handlers for UI buttons. They simply notify the user for now.
 */
function onPreviewEmails(e) {
  var state = getFormState(e);
  try {
    var previewResult;
    if (state.deleteScope === "thread") {
      previewResult = previewCurrentThread(e, state);
    } else {
      previewResult = previewMailbox(state);
    }

    return CardService.newActionResponseBuilder()
      .setNavigation(
        CardService.newNavigation().updateCard(createCleanerCard(e, { preview: previewResult }))
      )
      .setNotification(
        CardService.newNotification()
          .setText("Preview generated for the current filters.")
      )
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNavigation(
        CardService.newNavigation().updateCard(createCleanerCard(e, { error: err.message }))
      )
      .setNotification(
        CardService.newNotification()
          .setText("Preview failed: " + err.message)
      )
      .build();
  }
}

function onDeleteNow(e) {
  var summary = summarizeState(e);
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().updateCard(createCleanerCard(e, null))
    )
    .setNotification(
      CardService.newNotification()
        .setText("Delete action is not implemented yet. " + summary)
    )
    .build();
}

/**
 * Extracts current form inputs so UI reflects user selections.
 */
function getFormState(e) {
  var state = {
    deleteScope: "thread",
    subjectContains: "",
    subjectTokens: [],
    saveStarred: false,
    includeTrash: false
  };

  if (e && e.commonEventObject && e.commonEventObject.formInputs) {
    var inputs = e.commonEventObject.formInputs;
    state.deleteScope = getFirstValue(inputs, "deleteScope", state.deleteScope);
    state.subjectContains = getFirstValue(inputs, "subjectContains", state.subjectContains);
    state.saveStarred = hasSelection(inputs, "saveStarred", "starred");
    state.includeTrash = hasSelection(inputs, "includeTrash", "includeTrash");
  }

  state.subjectTokens = parseSubjectTokens(state.subjectContains);
  return state;
}

function getFirstValue(inputs, name, fallback) {
  var input = inputs[name];
  if (!input || !input.stringInputs || !input.stringInputs.value || !input.stringInputs.value.length) {
    return fallback;
  }
  return input.stringInputs.value[0];
}

function hasSelection(inputs, name, matchValue) {
  var input = inputs[name];
  if (!input || !input.stringInputs || !input.stringInputs.value) {
    return false;
  }
  var values = input.stringInputs.value;
  for (var i = 0; i < values.length; i++) {
    if (values[i] === matchValue) {
      return true;
    }
  }
  return false;
}

function summarizeState(e) {
  var state = getFormState(e);
  var parts = [
    "Scope: " + state.deleteScope,
    "Subject: " + (state.subjectContains || "any"),
    "Starred safe: " + (state.saveStarred ? "yes" : "no"),
    "Include trash: " + (state.includeTrash ? "yes" : "no")
  ];
  return parts.join(" | ");
}

var PREVIEW_MESSAGE_LIMIT = 5;
var PREVIEW_THREAD_SEARCH_LIMIT = 8;

function previewCurrentThread(e, state) {
  if (!e || !e.gmail || !e.gmail.threadId) {
    throw new Error("Open an email to preview the current thread or switch scope to Inbox/All mail.");
  }

  var thread = GmailApp.getThreadById(e.gmail.threadId);
  if (!thread) {
    throw new Error("Unable to load the selected thread.");
  }

  var matches = collectMessages(thread.getMessages(), state, PREVIEW_MESSAGE_LIMIT);
  var summary = matches.total === 0
    ? "No emails matched in the current thread."
    : "Showing top " + Math.min(matches.entries.length, PREVIEW_MESSAGE_LIMIT) + " of "
      + matches.total + " matching email(s) in this thread.";

  return {
    summary: summary,
    items: matches.entries,
    total: matches.total
  };
}

function previewMailbox(state) {
  var query = buildSearchQuery(state);
  var threads = GmailApp.search(query, 0, PREVIEW_THREAD_SEARCH_LIMIT);

  var previewItems = [];
  var scannedMatches = 0;

  for (var i = 0; i < threads.length; i++) {
    var remaining = PREVIEW_MESSAGE_LIMIT - previewItems.length;
    if (remaining <= 0) {
      break;
    }

    var threadMatches = collectMessages(threads[i].getMessages(), state, remaining);
    scannedMatches += threadMatches.total;
    previewItems = previewItems.concat(threadMatches.entries);
  }

  var totalMatches = getTotalMatchCount(query);
  if (totalMatches === null) {
    totalMatches = scannedMatches;
  }

  var scopeLabel = state.deleteScope === "inbox" ? "Inbox" : "All mail";
  var summary;
  if (totalMatches === 0) {
    summary = "No emails matched the current filters in your " + scopeLabel + ".";
  } else if (!previewItems.length) {
    summary = "Found " + totalMatches + " matching email(s) in " + scopeLabel + ", but none were available to preview.";
  } else {
    summary = "Showing top " + previewItems.length + " of " + totalMatches + " matching email(s) in " + scopeLabel + ".";
  }

  return {
    summary: summary,
    items: previewItems,
    total: totalMatches
  };
}

function buildSearchQuery(state) {
  var clauses = [];
  if (state.deleteScope === "inbox") {
    clauses.push("in:inbox");
  } else if (state.deleteScope === "all") {
    clauses.push("in:anywhere");
  }

  if (state.subjectTokens.length) {
    var subjectClause = state.subjectTokens.map(function (token) {
      return 'subject:"' + token.replace(/"/g, '\\"') + '"';
    }).join(" OR ");
    clauses.push("(" + subjectClause + ")");
  }

  if (state.saveStarred) {
    clauses.push("-is:starred");
  }

  if (!state.includeTrash) {
    clauses.push("-in:trash");
  } else if (state.deleteScope !== "inbox") {
    clauses.push("in:anywhere");
  }

  var query = clauses.join(" ").trim();
  return query || "in:anywhere";
}

function collectMessages(messages, state, limit) {
  var entries = [];
  var total = 0;
  for (var i = 0; i < messages.length; i++) {
    var message = messages[i];
    if (!state.includeTrash && message.isInTrash && message.isInTrash()) {
      continue;
    }
    if (state.saveStarred && message.isInStarred && message.isInStarred()) {
      continue;
    }
    if (!matchesSubject(message.getSubject(), state.subjectTokens)) {
      continue;
    }

    total++;
    if (!limit || entries.length < limit) {
      entries.push(buildPreviewEntry(message));
    }
  }

  return {
    total: total,
    entries: entries
  };
}

function parseSubjectTokens(value) {
  if (!value) {
    return [];
  }
  return value.split(/[,\\n]/)
    .map(function (token) { return token.trim().toLowerCase(); })
    .filter(function (token) { return token.length; });
}

function matchesSubject(subject, tokens) {
  if (!tokens || !tokens.length) {
    return true;
  }
  var haystack = (subject || "").toLowerCase();
  for (var i = 0; i < tokens.length; i++) {
    if (haystack.indexOf(tokens[i]) !== -1) {
      return true;
    }
  }
  return false;
}

function buildPreviewEntry(message) {
  var snippet = formatSnippet(message.getPlainBody());
  if (!snippet) {
    snippet = "(No preview content)";
  }
  return {
    subject: message.getSubject() || "(no subject)",
    from: message.getFrom() || "",
    date: formatDate(message.getDate()),
    snippet: snippet
  };
}

function formatSnippet(body) {
  if (!body) {
    return "";
  }
  var text = body.replace(/\\s+/g, " ").trim();
  if (text.length > 100) {
    return text.substring(0, 97) + "...";
  }
  return text;
}

function formatDate(date) {
  if (!date) {
    return "";
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a");
}

function getTotalMatchCount(query) {
  if (!query) {
    query = "in:anywhere";
  }

  try {
    if (typeof Gmail === "undefined" || !Gmail.Users || !Gmail.Users.Messages) {
      return null;
    }
    var response = Gmail.Users.Messages.list("me", {
      q: query,
      maxResults: 1,
      fields: "resultSizeEstimate"
    });
    if (response && typeof response.resultSizeEstimate === "number") {
      return response.resultSizeEstimate;
    }
  } catch (err) {
    Logger.log("Failed to determine total match count: " + err);
  }

  return null;
}


