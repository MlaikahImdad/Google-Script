var NOTICE_LIST_URL = "https://vulms.vu.edu.pk/NoticeBoard/NoticeBoard2.aspx";
var WEBHOOK_PROPERTY = "GOOGLE_CHAT_WEBHOOK_URL";
var SENT_NOTICES_PROPERTY = "sentNotices";

function doPost(e) {
  var data = JSON.parse(e.postData.contents || "{}");
  var command = ((data.message && data.message.text) || "").trim();

  if (command.indexOf("/fetchNotices") === 0) {
    try {
      var count = fetchAndShareNotices();
      return ContentService.createTextOutput(count ? "Notice shared." : "No new notice found.");
    } catch (error) {
      Logger.log(error.stack || error.message);
      return ContentService.createTextOutput("Error: " + error.message);
    }
  }

  return ContentService.createTextOutput("Invalid command. Use /fetchNotices.");
}

function testFetchLatestNoticeOnly() {
  var response = UrlFetchApp.fetch(NOTICE_LIST_URL, { muteHttpExceptions: true });
  assertOk(response, "fetch notice board");

  var html = response.getContentText();
  var cookies = getResponseCookies(response);
  var hidden = getAspNetHiddenFields(html);
  var notices = extractNotices(html);

  if (!notices.length) {
    throw new Error("No notices parsed from notice board.");
  }

  notices.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  var latestNotice = notices[0];
  var details = fetchNoticeDetails(latestNotice, hidden, cookies);

  Logger.log("LATEST NOTICE TITLE: " + latestNotice.title);
  Logger.log("LATEST NOTICE DATE: " + formatDate(latestNotice.date));
  Logger.log("LATEST NOTICE TARGET: " + latestNotice.target);
  Logger.log("DETAILS PREVIEW: " + details.text.slice(0, 1000));
  Logger.log("IMAGES: " + JSON.stringify(details.images));
}

function testGoogleChatWebhookOnly() {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty(WEBHOOK_PROPERTY);
  if (!webhookUrl) {
    throw new Error("Set script property " + WEBHOOK_PROPERTY + " to your Google Chat webhook URL.");
  }

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      text: "VU notice webhook test: Apps Script can post to this Google Chat space."
    }),
    muteHttpExceptions: true
  });

  assertOk(response, "send Google Chat test message");
  Logger.log("Google Chat webhook test succeeded.");
}

function testGoogleChatCardOnly() {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty(WEBHOOK_PROPERTY);
  if (!webhookUrl) {
    throw new Error("Set script property " + WEBHOOK_PROPERTY + " to your Google Chat webhook URL.");
  }

  var payload = {
    text: "\u200B",
    cardsV2: [
      {
        cardId: "vu-card-test",
        card: {
          header: {
            title: "VU Notice Card Test",
            subtitle: "Formatting check"
          },
          sections: [
            {
              widgets: [
                {
                  textParagraph: {
                    text: "<b>Full notice title appears here</b><br><br>Published On: 16-05-26 9:20 AM<br>Demo link: <a href=\"https://datesheet.vu.edu.pk\">datesheet.vu.edu.pk</a>"
                  }
                },
                {
                  buttonList: {
                    buttons: [
                      {
                        text: "Open notice board",
                        onClick: { openLink: { url: NOTICE_LIST_URL } }
                      }
                    ]
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log("CARD TEST response code: " + response.getResponseCode());
  Logger.log("CARD TEST response body: " + response.getContentText().slice(0, 1000));
  assertOk(response, "send Google Chat card test");
}

function resetSentNoticesForTesting() {
  PropertiesService.getScriptProperties().deleteProperty(SENT_NOTICES_PROPERTY);
  Logger.log("Cleared sent notice history for testing.");
}

function debugLatestNoticeDetailHtml() {
  var response = UrlFetchApp.fetch(NOTICE_LIST_URL, { muteHttpExceptions: true });
  assertOk(response, "fetch notice board");

  var html = response.getContentText();
  var cookies = getResponseCookies(response);
  var hidden = getAspNetHiddenFields(html);
  var notices = extractNotices(html);

  notices.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  var latestNotice = notices[0];
  var detailsResponse = fetchRawNoticeDetailsResponse(latestNotice, hidden, cookies);
  var detailHtml = detailsResponse.getContentText();
  var titleIndex = detailHtml.indexOf(latestNotice.title);
  var snippetStart = Math.max(0, titleIndex - 1500);
  var snippetEnd = Math.min(detailHtml.length, titleIndex + 5000);

  Logger.log("DEBUG TITLE: " + latestNotice.title);
  Logger.log("DEBUG HREF: " + latestNotice.href);
  Logger.log("DEBUG TARGET: " + latestNotice.target);
  Logger.log("DEBUG COOKIES: " + cookies);
  Logger.log("DEBUG HIDDEN FIELD COUNT: " + Object.keys(hidden).length);
  Logger.log("DEBUG TITLE INDEX: " + titleIndex);
  Logger.log("DEBUG HTML SNIPPET: " + detailHtml.substring(snippetStart, snippetEnd));
}

function fetchAndShareNotices() {
  var response = UrlFetchApp.fetch(NOTICE_LIST_URL, { muteHttpExceptions: true });
  assertOk(response, "fetch notice board");

  var html = response.getContentText();
  var cookies = getResponseCookies(response);
  var hidden = getAspNetHiddenFields(html);
  var notices = extractNotices(html);

  if (!notices.length) {
    Logger.log("No notices parsed from notice board.");
    return 0;
  }

  notices.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  var latestNotice = notices[0];
  if (!isNewNotice(latestNotice)) {
    Logger.log("Latest notice already sent: " + latestNotice.title);
    return 0;
  }

  var details = fetchNoticeDetails(latestNotice, hidden, cookies);
  latestNotice.details = details.text;
  latestNotice.images = details.images;
  latestNotice.url = details.url || NOTICE_LIST_URL;
  Logger.log("Notice details preview: " + latestNotice.details.slice(0, 1000));

  sendNoticeToChat(latestNotice);
  saveSentNotice(latestNotice);
  return 1;
}

function fetchRawNoticeDetailsResponse(notice, hidden, cookies) {
  if (/NewsDetails2\.aspx/i.test(notice.href)) {
    return UrlFetchApp.fetch(absolutizeUrl(notice.href, NOTICE_LIST_URL), {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: cookies ? { Cookie: cookies } : {}
    });
  }

  var payload = {};
  Object.keys(hidden || {}).forEach(function (key) {
    payload[key] = hidden[key];
  });
  payload.__EVENTTARGET = notice.target;
  payload.__EVENTARGUMENT = "";

  return UrlFetchApp.fetch(NOTICE_LIST_URL, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: payload,
    followRedirects: true,
    muteHttpExceptions: true,
    headers: cookies ? {
      Cookie: cookies,
      Referer: NOTICE_LIST_URL
    } : {
      Referer: NOTICE_LIST_URL
    }
  });
}

function extractNotices(html) {
  var notices = [];
  var itemPattern = /<div class="m-timeline-3__item-desc">([\s\S]*?)(?=<div class="m-timeline-3__item-desc">|<div class="m-portlet__foot|$)/g;
  var match;

  while ((match = itemPattern.exec(html)) !== null) {
    var block = match[1];
    var link = block.match(/<a\b[^>]*class="[^"]*newstext[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
    var dateMatch = block.match(/<span class="m-link m-link--metal m-timeline-3__item-link">([\s\S]*?)<\/span>/i);

    if (!link || !dateMatch) continue;

    var href = decodeHtml(link[1]);
    var title = cleanText(link[2]);
    var date = new Date(cleanText(dateMatch[1]));

    if (!title || isNaN(date.getTime())) continue;

    notices.push({
      title: title,
      href: href,
      target: parseDoPostBackTarget(href),
      date: date
    });
  }

  return notices;
}

function fetchNoticeDetails(notice, hidden, cookies) {
  var response;
  var url = NOTICE_LIST_URL;

  if (/NewsDetails2\.aspx/i.test(notice.href)) {
    url = absolutizeUrl(notice.href, NOTICE_LIST_URL);
    response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: cookies ? { Cookie: cookies } : {}
    });
  } else if (notice.target) {
    var payload = {};
    Object.keys(hidden || {}).forEach(function (key) {
      payload[key] = hidden[key];
    });
    payload.__EVENTTARGET = notice.target;
    payload.__EVENTARGUMENT = "";

    response = UrlFetchApp.fetch(NOTICE_LIST_URL, {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: payload,
      followRedirects: true,
      muteHttpExceptions: true,
      headers: cookies ? {
        Cookie: cookies,
        Referer: NOTICE_LIST_URL
      } : {
        Referer: NOTICE_LIST_URL
      }
    });
  } else {
    url = absolutizeUrl(notice.href, NOTICE_LIST_URL);
    response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: cookies ? { Cookie: cookies } : {}
    });
  }

  assertOk(response, "fetch notice details for " + notice.title);

  url = getFinalUrl(response, url);
  var html = response.getContentText();
  Logger.log("Notice href: " + notice.href);
  Logger.log("Notice target: " + notice.target);
  Logger.log("Detail final URL: " + url);
  Logger.log("Detail HTML title check for " + notice.title + ": " + html.indexOf(notice.title));
  return {
    text: extractNoticeText(html, notice.title),
    images: extractNoticeImages(html, notice.title),
    url: url
  };
}

function sendNoticeToChat(notice) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty(WEBHOOK_PROPERTY);
  if (!webhookUrl) {
    throw new Error("Set script property " + WEBHOOK_PROPERTY + " to your Google Chat webhook URL.");
  }

  var detailText = truncateText(notice.details || "No details found.", 6000);

  var widgets = buildNoticeTextWidgets(notice, detailText);
  widgets.push({
    buttonList: {
      buttons: [
        {
          text: "Open notice board",
          onClick: { openLink: { url: NOTICE_LIST_URL } }
        }
      ]
    }
  });

  var payload = {
    cardsV2: [
      {
        cardId: "vu-notice",
        card: {
          header: {
            title: notice.title,
            subtitle: formatDate(notice.date)
          },
          sections: [{ widgets: widgets }]
        }
      }
    ]
  };

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log("Google Chat response code: " + response.getResponseCode());
  Logger.log("Google Chat response body: " + response.getContentText().slice(0, 1000));

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    Logger.log("Rich card failed. Sending plain-text fallback.");
    sendPlainTextNoticeToChat(webhookUrl, notice, detailText);
    return;
  }
}

function extractNoticeText(html, title) {
  var candidates = extractDetailBlocks(html);
  candidates.push(extractAroundTitle(html, title));
  candidates.push(extractMainContentAfterTitle(html, title));
  var bestText = "";
  var bestScore = 0;

  for (var i = 0; i < candidates.length; i++) {
    var text = htmlToReadableText(candidates[i]);
    text = trimNoticeText(text, title);

    var score = scoreNoticeText(text, title);
    if (isUsefulNoticeText(text, title) && score > bestScore) {
      bestText = text;
      bestScore = score;
    }
  }

  return bestText || "No details found.";
}

function sendPlainTextNoticeToChat(webhookUrl, notice, detailText) {
  var body = stripLinksForPlainText(removeLeadingTitleAndDate(detailText, notice.title));
  var message = "*" + notice.title + "*\n" +
    "Date: " + formatDate(notice.date) + "\n\n" +
    truncateText(body, 3500) + "\n\n" +
    "Notice board: " + NOTICE_LIST_URL;

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: message }),
    muteHttpExceptions: true
  });

  Logger.log("Plain fallback response code: " + response.getResponseCode());
  Logger.log("Plain fallback response body: " + response.getContentText().slice(0, 1000));
  assertOk(response, "send plain-text notice to Google Chat");
}

function stripLinksForPlainText(text) {
  return stripEmbeddedImages(String(text || "")).replace(/<[^>]+>/g, "");
}

function extractNoticeImages(html, title) {
  var blocks = extractDetailBlocks(html);
  blocks.push(extractAroundTitle(html, title));
  blocks.push(extractMainContentAfterTitle(html, title));

  var block = blocks.join("\n") || html;
  var images = [];
  var imgPattern = /<img\b[^>]*src="([^"]+)"[^>]*>/gi;
  var match;

  while ((match = imgPattern.exec(block)) !== null) {
    var imageUrl = absolutizeUrl(decodeHtml(match[1]), NOTICE_LIST_URL);
    if (isUsefulNoticeImage(imageUrl) && images.indexOf(imageUrl) === -1) {
      images.push(imageUrl);
    }
  }

  return images;
}

function isUsefulNoticeImage(imageUrl) {
  if (!/^https:\/\//i.test(imageUrl)) return false;
  if (/\.(svg|ico)(\?|$)/i.test(imageUrl)) return false;
  if (/\/(logo|favicon|loader|spinner|avatar|profile|icon)[^\/]*\./i.test(imageUrl)) return false;
  return true;
}

function stripEmbeddedImages(value) {
  return String(value || "")
    .replace(/<img\b[^>]*>/gi, " ")
    .replace(/data:\s*image\/[a-z0-9.+-]+;base64,[a-z0-9+/=\s]+/gi, " ");
}

function extractDetailBlock(html) {
  var blocks = extractDetailBlocks(html);
  return blocks.length ? blocks[0] : "";
}

function extractDetailBlocks(html) {
  var patterns = [
    /<div[^>]+class=["'][^"']*paraGraphtext[^"']*["'][^>]*>([\s\S]*?)<\/div>/i,
    /<div[^>]+id=["']noticeContent["'][^>]*>([\s\S]*?)<\/div>/i,
    /<div[^>]+class=["'][^"']*news-detail[^"']*["'][^>]*>([\s\S]*?)<\/div>/i,
    /<span[^>]+id=["'][^"']*(?:lbl|Label)[^"']*(?:Detail|Description|Text|News)[^"']*["'][^>]*>([\s\S]*?)<\/span>/i,
    /<div[^>]+id=["'][^"']*(?:Detail|Description|Text|News)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i
  ];
  var blocks = [];

  for (var i = 0; i < patterns.length; i++) {
    var match = html.match(patterns[i]);
    if (match) blocks.push(match[1]);
  }

  return blocks;
}

function extractBodyHtml(html) {
  var match = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return match ? match[1] : html;
}

function extractAroundTitle(html, title) {
  var lowerHtml = String(html || "").toLowerCase();
  var lowerTitle = String(title || "").toLowerCase();
  var titleIndex = lowerHtml.indexOf(lowerTitle);
  if (titleIndex < 0) return "";

  var start = Math.max(0, lowerHtml.lastIndexOf("<div", titleIndex));
  var endMarkers = [
    '<div class="m-portlet__foot',
    '<footer'
  ];
  var end = -1;

  for (var i = 0; i < endMarkers.length; i++) {
    var markerIndex = lowerHtml.indexOf(endMarkers[i], titleIndex + lowerTitle.length);
    if (markerIndex >= 0 && (end < 0 || markerIndex < end)) {
      end = markerIndex;
    }
  }

  if (end < 0) {
    end = Math.min(html.length, titleIndex + 20000);
  }

  return html.substring(start, end);
}

function extractMainContentAfterTitle(html, title) {
  var lowerHtml = String(html || "").toLowerCase();
  var lowerTitle = String(title || "").toLowerCase();
  var titleIndex = lowerHtml.indexOf(lowerTitle);
  if (titleIndex < 0) return "";

  var start = Math.max(0, lowerHtml.lastIndexOf("<div", titleIndex));
  var end = lowerHtml.indexOf("important links & notifications", titleIndex);

  if (end < 0) {
    end = lowerHtml.indexOf("<footer", titleIndex);
  }

  if (end < 0) {
    end = Math.min(html.length, titleIndex + 30000);
  }

  return html.substring(start, end);
}

function htmlToReadableText(html) {
  var withoutNoise = stripEmbeddedImages(String(html || ""))
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, function (_, href, label) {
      var text = cleanText(label);
      var decodedHref = decodeHtml(href);
      if (!text) return "";
      if (/^\s*javascript:/i.test(decodedHref)) return text;
      return text + " (" + absolutizeUrl(decodedHref, NOTICE_LIST_URL) + ")";
    })
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ");

  return cleanText(withoutNoise).replace(/\n{3,}/g, "\n\n");
}

function trimNoticeText(text, title) {
  text = String(text || "");

  var titleIndex = text.toLowerCase().indexOf(String(title || "").toLowerCase());
  if (titleIndex >= 0) {
    text = text.substring(titleIndex);
  }

  text = text
    .replace(/^.*?Published On:\s*/i, "Published On: ")
    .replace(/\bPlease enter Text to search\.\b/gi, "")
    .replace(/\bBack\b\s*/i, "")
    .replace(/\bNews & Events\b\s*/i, "")
    .replace(/\bImportant Links & Notifications\b[\s\S]*$/i, "")
    .trim();

  return text;
}

function isUsefulNoticeText(text, title) {
  if (!text || text === "No details found.") return false;
  if (text.indexOf("WebForm_DoPostBackWithOptions") !== -1) return false;
  if (text.length < String(title || "").length + 30) return false;
  return true;
}

function scoreNoticeText(text, title) {
  var score = String(text || "").length;

  if (/\bTable of Contents\b/i.test(text)) {
    score -= 800;
  }

  if (/\b(guidelines|instructions|schedule|eligibility|procedure|students|examination|support|important|note)\b/i.test(text)) {
    score += 1200;
  }

  if (normalizeForCompare(text).indexOf(normalizeForCompare(title)) >= 0) {
    score += 500;
  }

  return score;
}

function parseDoPostBackTarget(href) {
  var decoded = decodeHtml(href || "");
  var match = decoded.match(/__doPostBack\('([^']+)'/i);
  if (match) return match[1];

  match = decoded.match(/WebForm_PostBackOptions\("([^"]+)"/i);
  return match ? match[1] : "";
}

function getAspNetHiddenFields(html) {
  var fields = {};
  var inputPattern = /<input\b[^>]*type=["']hidden["'][^>]*>/gi;
  var input;

  while ((input = inputPattern.exec(html)) !== null) {
    var tag = input[0];
    var nameMatch = tag.match(/\bname=["']([^"']+)["']/i);
    if (!nameMatch) continue;

    var valueMatch = tag.match(/\bvalue=["']([^"']*)["']/i);
    fields[nameMatch[1]] = valueMatch ? decodeHtml(valueMatch[1]) : "";
  }

  fields.__VIEWSTATE = fields.__VIEWSTATE || getInputValueByName(html, "__VIEWSTATE");
  fields.__VIEWSTATEGENERATOR = fields.__VIEWSTATEGENERATOR || getInputValueByName(html, "__VIEWSTATEGENERATOR");
  fields.__EVENTVALIDATION = fields.__EVENTVALIDATION || getInputValueByName(html, "__EVENTVALIDATION");
  return fields;
}

function getResponseCookies(response) {
  var headers = response.getAllHeaders ? response.getAllHeaders() : {};
  var setCookie = headers["Set-Cookie"] || headers["set-cookie"];

  if (!setCookie) return "";

  if (Object.prototype.toString.call(setCookie) !== "[object Array]") {
    setCookie = [setCookie];
  }

  return setCookie.map(function (cookie) {
    return String(cookie).split(";")[0];
  }).join("; ");
}

function getInputValueByName(html, name) {
  var inputPattern = /<input\b[^>]*>/gi;
  var input;

  while ((input = inputPattern.exec(html)) !== null) {
    var tag = input[0];
    var nameMatch = tag.match(/\b(?:name|id)=["']([^"']+)["']/i);
    if (!nameMatch || nameMatch[1] !== name) continue;

    var valueMatch = tag.match(/\bvalue=["']([^"']*)["']/i);
    return valueMatch ? decodeHtml(valueMatch[1]) : "";
  }

  return "";
}

function isNewNotice(notice) {
  var sent = JSON.parse(PropertiesService.getScriptProperties().getProperty(SENT_NOTICES_PROPERTY) || "[]");
  return sent.indexOf(noticeKey(notice)) === -1;
}

function saveSentNotice(notice) {
  var props = PropertiesService.getScriptProperties();
  var sent = JSON.parse(props.getProperty(SENT_NOTICES_PROPERTY) || "[]");
  sent.push(noticeKey(notice));
  props.setProperty(SENT_NOTICES_PROPERTY, JSON.stringify(sent.slice(-100)));
}

function noticeKey(notice) {
  return notice.title + "::" + formatDate(notice.date);
}

function cleanText(value) {
  return decodeHtml(String(value || "").replace(/<[^>]+>/g, " "))
    .replace(/[ \t]+/g, " ")
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .trim();
}

function decodeHtml(value) {
  return String(value || "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function escapeForChat(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function formatDetailsForChat(text) {
  return String(text || "")
    .split(/\n/)
    .map(function (line) {
      return formatLineLinksForChat(line);
    })
    .join("<br>");
}

function formatNoticeBodyForChat(notice, details) {
  var title = String(notice.title || "").trim();
  var body = String(details || "").trim();
  var escapedTitle = escapeForChat(title);

  body = removeLeadingTitleAndDate(body, title).trim();

  return "<b>" + escapedTitle + "</b><br><br>" + formatDetailsForChat(body);
}

function buildNoticeTextWidgets(notice, details) {
  var title = String(notice.title || "").trim();
  var body = removeLeadingTitleAndDate(String(details || "").trim(), title).trim();
  body = removeTableOfContentsAnchorLinks(body);

  var chunks = splitTextForChatWidgets(body, 900);
  var widgets = [];

  if (!chunks.length) {
    chunks = ["No details found."];
  }

  for (var i = 0; i < chunks.length; i++) {
    var text = formatDetailsForChat(chunks[i]);
    if (i === 0) {
      text = "<b>" + escapeForChat(title) + "</b><br><br>" + text;
    }

    widgets.push({
      textParagraph: {
        text: text
      }
    });
  }

  return widgets;
}

function splitTextForChatWidgets(text, maxChunkLength) {
  var lines = String(text || "").split(/\n/);
  var chunks = [];
  var current = "";

  lines.forEach(function (line) {
    var next = current ? current + "\n" + line : line;

    if (next.length > maxChunkLength && current) {
      chunks.push(current);
      current = line;
    } else {
      current = next;
    }
  });

  if (current) {
    chunks.push(current);
  }

  return chunks;
}

function removeTableOfContentsAnchorLinks(text) {
  var lines = String(text || "").split(/\n/);
  var cleaned = [];
  var insideToc = false;

  lines.forEach(function (line) {
    if (/^\s*Table of Contents\s*$/i.test(line)) {
      insideToc = true;
      cleaned.push(line);
      return;
    }

    if (insideToc && /^\s*\d+\.\s+/.test(line)) {
      insideToc = false;
    }

    if (insideToc) {
      line = line.replace(/\s*\(https?:\/\/[^)]*#.*?\)\s*$/i, "");
    }

    cleaned.push(line);
  });

  return cleaned.join("\n");
}

function addNoticeImageWidgets(widgets, notice) {
  var images = (notice.images || []).slice(0, 1);

  images.forEach(function (imageUrl) {
    widgets.push({
      image: {
        imageUrl: imageUrl,
        altText: notice.title
      }
    });
  });
}

function removeLeadingTitleAndDate(text, title) {
  var lines = String(text || "").split(/\n/);
  var normalizedTitle = normalizeForCompare(title);

  while (lines.length && !lines[0].trim()) {
    lines.shift();
  }

  if (lines.length && normalizeForCompare(lines[0]) === normalizedTitle) {
    lines.shift();
  }

  while (lines.length && !lines[0].trim()) {
    lines.shift();
  }

  if (lines.length && /^date:\s*/i.test(lines[0].trim())) {
    lines.shift();
  }

  return lines.join("\n");
}

function normalizeForCompare(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function formatLineLinksForChat(line) {
  line = String(line || "");

  var fullLineLink = line.match(/^(\s*-\s*)?(.+?)\s*\((https?:\/\/[^)]+)\)\s*$/i);
  if (fullLineLink) {
    var prefix = fullLineLink[1] || "";
    var label = fullLineLink[2];
    var url = fullLineLink[3];

    if (isUnsafeCardUrl(url)) {
      return escapeForChat(prefix + label);
    }

    url = sanitizeChatUrl(url);
    return escapeForChat(prefix) + '<a href="' + escapeAttributeForChat(url) + '">' + escapeForChat(label) + "</a>";
  }

  return escapeForChat(line).replace(/(https?:\/\/[^\s<)]+)/g, function (url) {
    if (isUnsafeCardUrl(url)) {
      return escapeForChat(url);
    }

    var safeUrl = sanitizeChatUrl(url);
    return '<a href="' + escapeAttributeForChat(safeUrl) + '">' + escapeForChat(url) + "</a>";
  });
}

function sanitizeChatUrl(url) {
  url = String(url || "").trim();
  try {
    return encodeURI(url);
  } catch (error) {
    return url.replace(/ /g, "%20");
  }
}

function isUnsafeCardUrl(url) {
  url = String(url || "");
  if (/#/.test(url)) return true;
  if (/[<>"']/.test(url)) return true;
  return false;
}

function escapeAttributeForChat(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function absolutizeUrl(url, baseUrl) {
  if (!url) return "";
  if (/^https?:\/\//i.test(url)) return url;
  if (url.indexOf("//") === 0) return "https:" + url;

  var base = baseUrl.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!base) return url;

  if (url.charAt(0) === "/") return base[1] + url;
  return base[1] + baseUrl.replace(/\/[^\/]*$/, "/").replace(/^https?:\/\/[^\/]+/i, "") + url;
}

function getFinalUrl(response, fallbackUrl) {
  var headers = response.getAllHeaders ? response.getAllHeaders() : {};
  return headers["X-Final-Url"] || headers["x-final-url"] || fallbackUrl;
}

function truncateText(text, maxLength) {
  text = String(text || "");
  return text.length > maxLength ? text.substring(0, maxLength - 20) + "\n\n[details truncated]" : text;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM dd, yyyy");
}

function assertOk(response, action) {
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Failed to " + action + ". HTTP " + code + ": " + response.getContentText().slice(0, 300));
  }
}
