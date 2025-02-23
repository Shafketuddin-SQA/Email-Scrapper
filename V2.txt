function processWebsites() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();

    // Retrieve the starting row from script properties
    var scriptProperties = PropertiesService.getScriptProperties();
    var startRow = parseInt(scriptProperties.getProperty('START_ROW'), 10) || 1;
    var batchSize = 50;  // Batch size for processing URLs

    if (startRow > lastRow) {
      Logger.log("All rows processed.");
      return;
    }

    var endRow = Math.min(startRow + batchSize - 1, lastRow);
    processBatch(sheet, startRow, endRow);

    // Update the starting row for the next execution
    scriptProperties.setProperty('START_ROW', (endRow + 1).toString());

    // Set up a trigger to continue processing after a delay
    if (endRow < lastRow) {
      ScriptApp.getProjectTriggers().forEach(function(trigger) {
        if (trigger.getHandlerFunction() === 'processWebsites') {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      ScriptApp.newTrigger('processWebsites')
        .timeBased()
        .after(1000 * 60 * 2) // Wait for 2 minutes
        .create();
    } else {
      // Clear the script properties if processing is complete
      scriptProperties.deleteProperty('START_ROW');
      Logger.log("Processing complete.");
    }

  } catch (error) {
    Logger.log("Error in processWebsites: " + error.toString());
  }
}

function processBatch(sheet, startRow, endRow) {
  var websiteRange = sheet.getRange("A" + startRow + ":A" + endRow);
  var emailsRange = sheet.getRange("B" + startRow + ":B" + endRow);

  var websites = websiteRange.getValues();
  var results = [];

  websites.forEach(function(website, index) {
    var url = website[0];
    var rowIndex = startRow + index;
    var existingEmail = emailsRange.getCell(index + 1, 1).getValue();

    if (!existingEmail) {
      var extractedEmail = fetchEmailsFromWebsite(url);
      if (extractedEmail) {
        sheet.getRange(rowIndex, 2).setValue(extractedEmail);
        results.push({ url: url, email: extractedEmail });
      }
    }
  });

  Logger.log("Batch processed from row " + startRow + " to " + endRow);
}

function fetchEmailsFromWebsite(url) {
  try {
    var domain = extractDomainFromUrl(url);
    var emails = crawlAndFetchEmails(url, domain);

    if (emails.length > 0) {
      return removeDuplicates(emails).join(", ");
    }
    return searchSocialMediaForEmail(domain);
  } catch (e) {
    Logger.log("Error in fetchEmailsFromWebsite for URL " + url + ": " + e.toString());
    return null;
  }
}

function crawlAndFetchEmails(url, domain, depth = 1) {
  var emails = [];
  var visitedUrls = new Set();

  function crawl(url, currentDepth) {
    if (currentDepth > depth || visitedUrls.has(url)) {
      return;
    }

    visitedUrls.add(url);

    try {
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var content = response.getContentText();

      // Extract emails
      var emailPattern = new RegExp("[a-zA-Z0-9._%+-]+@" + domain.replace(/\./g, '\\.'), 'g');
      var matches = content.match(emailPattern);

      if (matches) {
        emails = emails.concat(matches.filter(validateEmail));
      }

      // Extract links within the same domain
      var linkPattern = new RegExp("https?://(www\\.)?" + domain.replace(/\./g, '\\.') + "[^\"']*", 'g');
      var links = content.match(linkPattern);

      if (links) {
        links.forEach(function(link) {
          crawl(link, currentDepth + 1);
        });
      }
    } catch (e) {
      Logger.log("Error crawling URL " + url + ": " + e.toString());
    }
  }

  crawl(url, 1);
  return emails;
}

function searchSocialMediaForEmail(domain) {
  var socialMediaSites = ['linkedin.com', 'facebook.com', 'twitter.com', 'instagram.com', 'plus.google.com'];
  var emails = [];

  socialMediaSites.forEach(function(site) {
    try {
      var searchQuery = site + " " + domain + " email";
      var searchUrl = "https://www.google.com/search?q=" + encodeURIComponent(searchQuery);

      var response = UrlFetchApp.fetch(searchUrl, {muteHttpExceptions: true});
      var content = response.getContentText();

      var emailPattern = new RegExp("[a-zA-Z0-9._%+-]+@" + domain.replace(/\./g, '\\.'), 'g');
      var matches = content.match(emailPattern);

      if (matches) {
        emails = emails.concat(matches.filter(validateEmail));
      }
    } catch (e) {
      Logger.log("Error searching " + site + " for " + domain + ": " + e.toString());
    }
  });

  return emails.length > 0 ? removeDuplicates(emails).join(", ") : null;
}

function removeDuplicates(array) {
  return [...new Set(array)];
}

function validateEmail(email) {
  var emailParts = email.split("@");
  if (emailParts.length !== 2) {
    return false;
  }
  var domainPart = emailParts[1];
  var validDomainPattern = /^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return validDomainPattern.test(domainPart);
}

function extractDomainFromUrl(url) {
  var domainRegex = /^(?:https?:\/\/)?(?:www\.)?([^\/:]+)/i;
  var match = url.match(domainRegex);
  return match ? match[1].toLowerCase() : '';
}
