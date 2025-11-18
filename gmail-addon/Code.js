/**
 * Called whenever the add-on is opened from an email.
 * e is the event object containing message metadata.
 */
function buildAddOn(e) {
  var accessToken = e.gmail.accessToken;
  var messageId = e.gmail.messageId;

  // Get message details
  var message = GmailApp.getMessageById(messageId);
  var subject = message.getSubject();
  var from = message.getFrom();

  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText("👋 Hello from Plughub Email Cleaner!"))
    .addWidget(CardService.newTextParagraph()
      .setText("**From:** " + from))
    .addWidget(CardService.newTextParagraph()
      .setText("**Subject:** " + subject));

  var card = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Plughub Email Cleaner")
        .setSubtitle("Quick glance at this email")
    )
    .addSection(section)
    .build();

  return [card];
}

/**
 * Optional: what to show on homepage (no email selected).
 */
function buildHomePage(e) {
  var section = CardService.newCardSection()
    .addWidget(
      CardService.newTextParagraph()
        .setText("Welcome to Plughub Email Cleaner.\nSelect an email to see details here.")
    );

  var card = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Plughub Email Cleaner")
        .setSubtitle("Homepage")
    )
    .addSection(section)
    .build();

  return [card];
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
