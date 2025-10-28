function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var command = data.message.text;
  
  if (command === '/fetchNotices') {
    try {
      fetchAndShareNotices();
      return ContentService.createTextOutput("Notices fetched and shared successfully!");
    } catch (error) {
      return ContentService.createTextOutput("An error occurred while fetching and sharing notices: " + error.message);
    }
  } else {
    return ContentService.createTextOutput("Invalid command. Please use /fetchNotices to fetch notices.");
  }
}

function fetchAndShareNotices() {
  var url = "https://vulms.vu.edu.pk/NoticeBoard/NoticeBoard2.aspx";

  // Fetch the webpage content
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to fetch webpage. Response code: " + response.getResponseCode());
  }
  var content = response.getContentText();

  var vuStateGenerator = getInputValueById(content, "__VIEWSTATEGENERATOR");
  var vuEventValidation = getInputValueById(content, "__EVENTVALIDATION");
  var vuViewState = getInputValueById(content, "__VIEWSTATE");

  // Extract notices
  var notices = extractNotices(content);
  var currentDate = new Date();
  var filteredNotices = filterNoticesByDate(notices, currentDate);

  Logger.log("Filtered Notices: " + JSON.stringify(filteredNotices));

  // Check for new notices before sending
  var newNotices = getNewNotices(filteredNotices);
  Logger.log("New Notices to Send: " + JSON.stringify(newNotices));

  if (newNotices.length > 0) {
    sendNoticesToChat(newNotices);
    saveSentNotices(newNotices);
    Logger.log("Notices sent and saved.");
  } else {
    Logger.log("No new notices to send.");
  }
}

// Function to get notices that haven't been sent
function getNewNotices(notices) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sentNotices = JSON.parse(scriptProperties.getProperty("sentNotices") || "[]");

  Logger.log("Previously Sent Notices: " + JSON.stringify(sentNotices));

  var sentNoticesSet = new Set(sentNotices);

  var newNotices = notices.filter(function (notice) {
    var noticeKey = notice.title + "::" + notice.date.toDateString();
    return !sentNoticesSet.has(noticeKey);
  });

  return newNotices;
}

// Function to save sent notices properly
function saveSentNotices(notices) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sentNotices = JSON.parse(scriptProperties.getProperty("sentNotices") || "[]");

  notices.forEach(function (notice) {
    var noticeKey = notice.title + "::" + notice.date.toDateString();
    sentNotices.push(noticeKey);
  });

  scriptProperties.setProperty("sentNotices", JSON.stringify(sentNotices));

  Logger.log("Updated Sent Notices: " + JSON.stringify(sentNotices));
}


function parseDoPostBackParameter(href) {
  // Regular expression to match the __doPostBack function call
  var regex = /__doPostBack\(&#39;([^']+)&#39;,/;

  // Execute the regular expression on the href
  var match = regex.exec(href);

  // If a match is found, return the first parameter
  if (match && match.length >= 2) {
    return match[1];
  } else {
    return null; // No match found
  }
}

function extractNotices(htmlContent) {
  var notices = [];
  var regex = /<div class="m-timeline-3__item-desc">\s*<div>\s*<span class="m-timeline-3__item-text">\s*<a.*?title="([^"]+)"\s*class="newstext m--font-bolder"\s*href="([^"]+)">(?:<span.*?>)?([^<]+)(?:<\/span>)?<\/a>[\s\S]*?<span class="m-link m-link--metal m-timeline-3__item-link">(.*?)<\/span>/g;
  var match;

  while ((match = regex.exec(htmlContent)) !== null) {
    var title = match[3];
    var href = match[2];
    var dateString = match[4]; // Extracted date string

    // Parse the date string into a Date object
    var date = new Date(dateString);

    // Push the extracted data into the notices array
    notices.push({
      title: title,
      href: href,
      date: date,
      target: parseDoPostBackParameter(href)
    });
  }

  return notices;
}


function filterNoticesByDate(notices, currentDate) {
  var filteredNotices = [];

  notices.forEach(function(notice) {
    if (
      notice.date.getDate() === currentDate.getDate() &&
      notice.date.getMonth() === currentDate.getMonth() &&
      notice.date.getFullYear() === currentDate.getFullYear()
    ) {
      filteredNotices.push(notice);
    }
  });

  return filteredNotices;
}

// function sendNoticesToChat(notices) {
//   var formattedNotices = notices.map(function(notice) {
//     var noticeDetails = fetchNoticeDetails(notice.target); // Get full notice text
//     return "📢 *" + notice.title + "*\n📅 Date: " + notice.date.toDateString() + "\n\n";
//   }).join("\n\n");

//   var message = "📝 *New Notices:*\n\n" + formattedNotices;

//   var payload = {
//     "text": message
//   };
  
//   var options = {
//     "method": "post",
//     "contentType": "application/json",
//     "payload": JSON.stringify(payload)
//   };
  
//   var url = "https://chat.googleapis.com/v1/spaces/AAAAyQNIeHk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=BKe63wStVhpJoq668Ao9NBKGHFZEjAsPJmeOWXdFc2A"; // Replace with your Google Chat API URL

//   var response = UrlFetchApp.fetch(url, options);
//   if (response.getResponseCode() !== 200) {
//     throw new Error("Failed to send message to Google Chat. Response code: " + response.getResponseCode());
//   }
// }
function sendNoticesToChat(notices) {
  var formattedNotices = notices.map(function(notice) {
    var noticeDetails = fetchNoticeDetails(notice.target);
    return "📢 *" + notice.title + "*\n📅 Date: " + notice.date.toDateString() + "\n\n";
  }).join("\n\n");

  var message = "📝 *New Notice:*\n\n" + formattedNotices;

  var payload = {
    "text": message
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  var webhookUrls = [
    "[ADD THE WEBHOOK URL HERE]",
    "[ADD ANOTHER WEBHOOK URL HERE IF ANY]"
    // Add more webhook URLs here if any
  ];

  webhookUrls.forEach(function(url) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) {
        Logger.log("Failed to send message to: " + url + " | Code: " + response.getResponseCode());
      } else {
        Logger.log("Message sent to: " + url);
      }
    } catch (error) {
      Logger.log("Error sending to: " + url + " | " + error.message);
    }
  });
}


// Function to fetch notice details
function fetchNoticeDetails(noticeTarget) {
  var detailsUrl = "https://vulms.vu.edu.pk/NoticeBoard/NoticeBoard2.aspx"; // Find the correct URL
  var payload = {
    "__EVENTTARGET": noticeTarget,
    "__EVENTARGUMENT": ""
  };

  var options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": payload
  };

  var response = UrlFetchApp.fetch(detailsUrl, options);
  if (response.getResponseCode() === 200) {
    var html = response.getContentText();
    var details = extractNoticeText(html); // Extract text from the response
    return details;
  }

  return "⚠️ Unable to retrieve details.";
}



// Function to extract notice text from HTML response
// function extractNoticeText(html) {
//   var pattern = /<div id="noticeContent">(.*?)<\/div>/; // Adjust based on actual HTML
//   var match = html.match(pattern);
//   return match ? match[1].replace(/<[^>]+>/g, '') : "No details found.";
// }
function extractNoticeText(html) {
  var pattern = /<div[^>]+class=["']paraGraphtext[^"']*["'][^>]*>([\s\S]*?)<\/div>/i;
  var match = html.match(pattern);
  
  if (match) {
    var rawText = match[1]
      .replace(/<p[^>]*>/g, '\n')  // Replace <p> with new lines for readability
      .replace(/<\/p>/g, '')  // Remove closing <p> tags
      .replace(/<[^>]+>/g, '')  // Remove remaining HTML tags
      .trim();  // Trim unnecessary spaces

    return rawText || "⚠️ No details found.";
  }
  
  return "⚠️ No details found.";
}

function getNoticesFromFilteredNotices(notices, vuStateGenerator, vuEventValidation, vuViewState) {
  // var message = "New Notices:\n\n" + notices.join("\n\n");
  
  notices.forEach(function(notice) {
     var payload = {
      "__VIEWSTATEGENERATOR": vuStateGenerator,
      "__EVENTVALIDATION": vuEventValidation,
      "__VIEWSTATE": vuViewState,
      "__EVENTTARGET": notice.target
    };
    // Convert payload object to query string format
    var queryString = Object.keys(payload).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]);
    }).join('&');

    var options = {
      "method": "post",
      "contentType": "application/x-www-form-urlencoded",
      "payload": queryString // Use queryString instead of formData
    };
    var response = UrlFetchApp.fetch("https://vulms.vu.edu.pk/NoticeBoard/NoticeBoard2.aspx", options);
    var text = response.getContentText();
    if (response.getResponseCode() !== 200) {
      throw new Error("Failed to send message to Google Chat. Response code: " + response.getResponseCode());
    } 
  })
}


function getInputValueById(htmlContent, id) {
  var regex = new RegExp('<input[^>]*id="' + id + '"[^>]*value="([^"]+)"');
  var match = htmlContent.match(regex);
  
  if (match) {
    return match[1]; // Return the value captured by the first group in the regex
  } else {
    return null; // Return null if the input element with the specified ID is not found
  }
}
