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

      var hasCount = typeof context.preview.total === "number";
      var totalMatched = hasCount
        ? context.preview.total
        : (context.preview.items ? context.preview.items.length : 0);
      var totalLabel = hasCount
        ? (context.preview.truncated ? totalMatched + "+" : String(totalMatched))
        : String(totalMatched);
      section.addWidget(
        CardService.newTextParagraph()
          .setText("Total matched emails: " + totalLabel)
      );

      if (context.preview.items && context.preview.items.length) {
        context.preview.items.forEach(function (item) {
          section.addWidget(
            CardService.newKeyValue()
              .setTopLabel(item.from + " | " + item.date)
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
  var state = getFormState(e);
  try {
    var result = deleteMessagesForState(e, state);
    var summary = buildDeleteSummary(result, state);

    return CardService.newActionResponseBuilder()
      .setNavigation(
        CardService.newNavigation().updateCard(createCleanerCard(e, null))
      )
      .setNotification(
        CardService.newNotification()
          .setText(summary)
      )
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNavigation(
        CardService.newNavigation().updateCard(createCleanerCard(e, { error: err.message }))
      )
      .setNotification(
        CardService.newNotification()
          .setText("Delete failed: " + err.message)
      )
      .build();
  }
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
var COUNT_PAGE_SIZE = 500;
var COUNT_HARD_LIMIT = 20000;
var DELETE_OPERATION_LIMIT = 20000;

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
    total: matches.total,
    truncated: false
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

  var countInfo = getTotalMatchCount(query, state.includeTrash);
  var totalMatches = countInfo.count;
  var countTruncated = countInfo.truncated;
  if (totalMatches === null) {
    totalMatches = scannedMatches;
    countTruncated = previewItems.length >= PREVIEW_MESSAGE_LIMIT;
  }

  var scopeLabel = state.deleteScope === "inbox" ? "Inbox" : "All mail";
  var summary;
  if (totalMatches === 0) {
    summary = "No emails matched the current filters in your " + scopeLabel + ".";
  } else if (!previewItems.length) {
    summary = "Found " + totalMatches + " matching email(s) in " + scopeLabel + ", but none were available to preview.";
  } else {
    summary = "Showing top " + previewItems.length + " of " + totalMatches + " matching email(s) in " + scopeLabel + ".";
    if (countTruncated) {
      summary += " (Count truncated after " + COUNT_HARD_LIMIT + " messages.)";
    }
  }

  return {
    summary: summary,
    items: previewItems,
    total: totalMatches,
    truncated: countTruncated
  };
}

function deleteMessagesForState(e, state) {
  if (state.deleteScope === "thread") {
    return deleteThreadMessages(e, state);
  }
  return deleteMailboxMessages(state);
}

function deleteThreadMessages(e, state) {
  if (!e || !e.gmail || !e.gmail.threadId) {
    throw new Error("Open an email to delete the current thread or choose Inbox/All mail.");
  }

  var thread = GmailApp.getThreadById(e.gmail.threadId);
  if (!thread) {
    throw new Error("Unable to load the selected thread.");
  }

  var messages = thread.getMessages();
  var total = 0;
  var deleted = 0;
  var errors = 0;

  for (var i = 0; i < messages.length; i++) {
    var message = messages[i];
    if (!messageMatchesFilters(message, state)) {
      continue;
    }
    total++;
    if (trashMessage(message, state)) {
      deleted++;
    } else {
      errors++;
    }
  }

  return {
    scope: "thread",
    total: total,
    deleted: deleted,
    errors: errors,
    truncated: false
  };
}

function deleteMailboxMessages(state) {
  if (typeof Gmail === "undefined" || !Gmail.Users || !Gmail.Users.Messages) {
    throw new Error("Enable the Gmail advanced service in Apps Script to delete outside the current thread.");
  }

  var query = buildSearchQuery(state);
  var pageToken = null;
  var deleted = 0;
  var processed = 0;
  var errors = 0;
  var truncated = false;
  var stop = false;

  try {
    do {
    var params = {
        q: query,
        maxResults: COUNT_PAGE_SIZE,
        includeSpamTrash: !!state.includeTrash,
        fields: "messages(id,labelIds),nextPageToken"
      };
      if (pageToken) {
        params.pageToken = pageToken;
      }

      var response = Gmail.Users.Messages.list("me", params);
      var messages = (response && response.messages) ? response.messages : [];

      for (var i = 0; i < messages.length; i++) {
        if (processed >= DELETE_OPERATION_LIMIT) {
          truncated = true;
          stop = true;
          break;
        }
        var messageResource = messages[i];
        var labelIds = messageResource.labelIds || [];
        var isInTrash = labelIds.indexOf("TRASH") !== -1;
        var hardDelete = state.includeTrash && isInTrash;

        processed++;
        if (trashMessageById(messageResource.id, hardDelete)) {
          deleted++;
        } else {
          errors++;
        }
      }

      if (stop) {
        break;
      }

      pageToken = response && response.nextPageToken;
      if (!pageToken) {
        break;
      }
    } while (true);
  } catch (err) {
    Logger.log("Failed to delete Gmail messages: " + err);
    throw new Error("Unable to delete emails for the current filters.");
  }

  if (pageToken) {
    truncated = true;
  }

  return {
    scope: state.deleteScope,
    total: processed,
    deleted: deleted,
    errors: errors,
    truncated: truncated
  };
}

function buildSearchQuery(state) {
  var clauses = [];
  if (state.deleteScope === "inbox") {
    if (state.includeTrash) {
      clauses.push("(in:inbox OR in:trash)");
    } else {
      clauses.push("in:inbox");
    }
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
    if (!messageMatchesFilters(message, state)) {
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

function messageMatchesFilters(message, state) {
  if (!state.includeTrash && typeof message.isInTrash === "function" && message.isInTrash()) {
    return false;
  }
  if (state.saveStarred && typeof message.isStarred === "function" && message.isStarred()) {
    return false;
  }
  return matchesSubject(message.getSubject(), state.subjectTokens);
}

function trashMessage(message, state) {
  if (!message) {
    return false;
  }

  var hardDelete = false;
  if (state.includeTrash && typeof message.isInTrash === "function") {
    hardDelete = message.isInTrash();
  }

  if (!hardDelete && typeof message.moveToTrash === "function") {
    try {
      message.moveToTrash();
      return true;
    } catch (err) {
      Logger.log("Failed to move message to trash via GmailApp: " + err);
    }
  }

  if (typeof message.getId !== "function") {
    return false;
  }
  return trashMessageById(message.getId(), hardDelete);
}

function trashMessageById(messageId, hardDelete) {
  if (!messageId) {
    return false;
  }

  try {
    if (hardDelete && typeof Gmail !== "undefined" && Gmail.Users && Gmail.Users.Messages && Gmail.Users.Messages.delete) {
      Gmail.Users.Messages.delete("me", messageId);
    } else if (typeof Gmail !== "undefined" && Gmail.Users && Gmail.Users.Messages && Gmail.Users.Messages.trash) {
      Gmail.Users.Messages.trash("me", messageId);
    } else {
      GmailApp.getMessageById(messageId).moveToTrash();
    }
    return true;
  } catch (err) {
    Logger.log("Failed to remove message " + messageId + ": " + err);
    return false;
  }
}

function buildDeleteSummary(result, state) {
  if (!result.total) {
    return "No matching emails were found to delete.";
  }

  var scopeText;
  if (state.deleteScope === "thread") {
    scopeText = "this thread";
  } else if (state.deleteScope === "inbox") {
    scopeText = "your Inbox";
  } else {
    scopeText = "All mail";
  }

  var message = "Deleted " + result.deleted + " email(s) from " + scopeText;
  if (result.deleted !== result.total) {
    message += " out of " + result.total;
  }
  message += ".";

  if (result.errors) {
    message += " " + result.errors + " email(s) could not be deleted.";
  }
  if (result.truncated) {
    message += " Stopped early due to execution limits.";
  }

  return message;
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

function getTotalMatchCount(query, includeTrash) {
  if (!query) {
    query = "in:anywhere";
  }

  if (typeof Gmail === "undefined" || !Gmail.Users || !Gmail.Users.Messages) {
    return { count: null, truncated: false };
  }

  var total = 0;
  var truncated = false;
  var pageToken = null;

  try {
    do {
      var params = {
        q: query,
        maxResults: COUNT_PAGE_SIZE,
        includeSpamTrash: !!includeTrash,
        fields: "messages/id,nextPageToken"
      };
      if (pageToken) {
        params.pageToken = pageToken;
      }
      var response = Gmail.Users.Messages.list("me", params);
      if (response.messages) {
        total += response.messages.length;
      }
      pageToken = response.nextPageToken;
      if (!pageToken) {
        break;
      }
      if (total >= COUNT_HARD_LIMIT) {
        truncated = true;
        break;
      }
    } while (true);
  } catch (err) {
    Logger.log("Failed to determine total match count: " + err);
    return { count: null, truncated: false };
  }

  if (pageToken) {
    truncated = true;
  }

  return { count: total, truncated: truncated };
}


