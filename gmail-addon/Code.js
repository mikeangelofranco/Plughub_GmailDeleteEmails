/**
 * Called whenever the add-on is opened from an email.
 * e is the event object containing message metadata.
 */
function buildAddOn(e) {
  return [createCleanerCard(e)];
}

/**
 * Optional: what to show on homepage (no email selected).
 */
function buildHomePage(e) {
  return [createCleanerCard(e)];
}

/**
 * Builds the UI card that mirrors the provided mockup.
 */
function createCleanerCard(e) {
  var state = getFormState(e);

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
    )
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
  var summary = summarizeState(e);
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().updateCard(createCleanerCard(e))
    )
    .setNotification(
      CardService.newNotification()
        .setText("Preview is not implemented yet. " + summary)
    )
    .build();
}

function onDeleteNow(e) {
  var summary = summarizeState(e);
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().updateCard(createCleanerCard(e))
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
    saveStarred: false,
    includeTrash: false
  };

  if (!e || !e.commonEventObject || !e.commonEventObject.formInputs) {
    return state;
  }

  var inputs = e.commonEventObject.formInputs;
  state.deleteScope = getFirstValue(inputs, "deleteScope", state.deleteScope);
  state.subjectContains = getFirstValue(inputs, "subjectContains", state.subjectContains);
  state.saveStarred = hasSelection(inputs, "saveStarred", "starred");
  state.includeTrash = hasSelection(inputs, "includeTrash", "includeTrash");

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


