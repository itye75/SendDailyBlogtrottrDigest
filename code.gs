/**
 * Blogtrottr digest with AI page summaries only.
 * Columns: Email date | Source name | Post title | Link to post | Post digest
 * Requires a Script property GEMINI_API_KEY.
 */

var SENDER_FILTER = 'from:(busybee@blogtrottr.com)';
var PROCESSED_LABEL = 'Blogtrottr/Digested';
var MAX_THREADS = 200;
var MAX_LINKS_PER_MSG = 10;
var RECIPIENT = Session.getActiveUser().getEmail();

// New configuration variable for the number of days to process
var DAYS_TO_PROCESS = 7;

// Fetch and summarization controls
var HTTP_TIMEOUT_MS = 15000;
var MAX_PAGES = 120;
var MAX_HTML_BYTES = 1200000;
var EXTRACT_TEXT_LIMIT = 9999999;
var MODEL_NAME = 'gemini-1.5-flash';
// New variable to control summary length
var SUMMARY_SENTENCE_COUNT = 3;

function sendDailyBlogtrottrDigest() {
  var processed = getOrCreateLabel_(PROCESSED_LABEL);

  // Calculate the date range for the search query
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  var afterDate = new Date(today);
  afterDate.setDate(today.getDate() - DAYS_TO_PROCESS);

  // Format dates for Gmail search
  var formattedAfterDate = Utilities.formatDate(afterDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  // Updated query to fetch emails from the last X days that are not already processed.
  var query = SENDER_FILTER + ' after:' + formattedAfterDate;

  var threads = GmailApp.search(query, 0, MAX_THREADS);
  if (threads.length === 0) return;

  var items = [];
  var seen = {};

  for (var i = 0; i < threads.length; i++) {
  var th = threads[i];
  var messages = th.getMessages();
  for (var j = 0; j < messages.length; j++) {
    var m = messages[j];
    var subject = m.getSubject() || 'No subject';
    var html = m.getBody();
    var text = m.getPlainBody() || '';

    var links = extractLinksFromHtml_(html);
    if (links.length === 0) links = links.concat(extractLinksFromText_(text));

    var filtered = links
      .map(function(u) { return cleanUrl_(u); })
      .filter(function(u) { return isLikelyArticleUrl_(u); })
      .slice(0, MAX_LINKS_PER_MSG);

    // Skip messages with no relevant links
    if (filtered.length === 0) {
      continue;
    }

    // Use the subject as the key for 'seen' to ensure each email message is processed once
    if (seen[m.getId()]) continue;
    seen[m.getId()] = true;

    var date = m.getDate()
    var source = tryExtractSourceFromEmail_(subject, html, filtered)

    items.push({
      emailDate: date,
      subject: subject,
      url: filtered,
      source: source, // Use the first link to determine the source
      digest: '',
      title: ''
    });
  }
  th.addLabel(processed);
}

  if (items.length === 0) return;

  buildPageSummaries_(items);

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var subject = 'Blogtrottr Digest · ' + today;
  var html = renderHtmlTable_(items);
  var text = renderTextTable_(items);

logItems(items);

  GmailApp.sendEmail(RECIPIENT, subject, text, { htmlBody: html });
}

function logItems(items)
{
  for (var i = 0; i < items.length; i++) {
    Logger.log(
      'Adding post to digest:\n' +
      '  date: %s\n' +
      '  subject: %s\n' +
      '  url: %s\n' +
      '  source: %s\n' + 
      '  digest: %s\n' +
      '  title: %s',
      items[i].emailDate, items[i].subject, items[i].url, items[i].source, items[i].digest, items[i].title
    );
  }

}


// --- Page fetch and AI summarization ---
// --- Page fetch and AI summarization ---

function buildPageSummaries_(items) {
  var apiKey = getGeminiKey_();

  var fetched = 0;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];

    if (fetched >= MAX_PAGES) {
      it.digest = 'Skipped due to fetch limit.';
      it.title = 'N/A';
      continue;
    }

    var page = fetchArticleHtmlWithStrategies_(it.url);
    fetched++;

    if (!page || !page.html) {
      it.digest = 'Could not fetch the page.';
      it.title = 'N/A';
      continue;
    }

    // This is the key line. It attempts to extract the title from the fetched HTML.
    it.title = extractTitleFromHtml_(page.html) || 'No Title Found';

    var mainText = extractMainReadableText_(page.html, EXTRACT_TEXT_LIMIT);
    if (!mainText) {
      it.digest = 'Could not extract article text.';
      continue;
    }

    var summary = '';
    if (apiKey) {
      summary = callGeminiSummarize_(apiKey, it.url, it.source, mainText);
    }

    if (!summary) {
      summary = simpleFallbackSummary_(mainText);
    }

    it.digest = summary || 'No summary available.';
    Utilities.sleep(40);
  }
}

// --- Multi strategy fetch to improve success on common sites ---

function fetchArticleHtmlWithStrategies_(url) {
  var res = safeFetchHtml_(url);
  if (res) return res;

  res = safeFetchHtml_(url, { referer: 'https://www.google.com/' });
  if (res) return res;

  var head = safeFetchHtml_(url, { method: 'headlike' });
  if (head && head.html) {
    var redir = extractMetaRefreshUrl_(head.html, url);
    if (redir) {
      res = safeFetchHtml_(redir, { referer: url });
      if (res) return res;
    }
  }

  var cacheUrl = 'https://webcache.googleusercontent.com/search?q=cache:' + encodeURIComponent(url);
  res = safeFetchHtml_(cacheUrl);
  if (res) return res;

  return null;
}

function extractMetaRefreshUrl_(html, baseUrl) {
  var m = /<meta\s+http-equiv=["']refresh["'][^>]*content=["'][^;]+;\s*url=([^"']+)["']/i.exec(html || '');
  if (!m) return '';
  var raw = m[1].trim();
  try {
    return new URL(raw, baseUrl).toString();
  } catch (e) {
    return raw;
  }
}

function safeFetchHtml_(url, opts) {
  opts = opts || {};
  try {
    var headers = {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
      'Accept-Language': 'en-US,en;q=0.9',
      'Sec-Fetch-Dest': 'document',
      'Sec-Fetch-Mode': 'navigate',
      'Sec-Fetch-Site': 'none',
      'Sec-Fetch-User': '?1',
      'Upgrade-Insecure-Requests': '1',
      'Connection': 'keep-alive'
    };
    if (opts.referer) headers['Referer'] = opts.referer;

    var params = {
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      headers: headers
    };

    if (opts.method === 'headlike') {
      params.headers.Range = 'bytes=0-8191';
    }

    var resp = UrlFetchApp.fetch(url, params);
    var code = resp.getResponseCode();
    if (code < 200 || code >= 400) return null;

    var blob = resp.getBlob();
    var html = '';
    if (blob.getBytes().length > MAX_HTML_BYTES) {
      html = Utilities.newBlob(blob.getBytes().slice(0, MAX_HTML_BYTES)).getDataAsString();
    } else {
      var ctype = resp.getHeaders()['Content-Type'] || '';
      var charset = (ctype.match(/charset=([^;]+)/i) || [, ''])[1];
      html = charset ? Utilities.newBlob(blob.getBytes()).getDataAsString(charset) : resp.getContentText();
    }
    return { html: html };
  } catch (e) {
    Logger.log('Fetch error for URL %s: %s', url, e.message);
    return null;
  }
}

// --- Readability style main text extraction (local, no proxy) ---

function extractMainReadableText_(html, limit) {
  var h = String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '');  // remove HTML comments

  var plainText = normalizeWhitespace_(stripHtml_(h));
  return plainText.slice(0, limit);
}

function extractTitleFromHtml_(html) {
  try {
    var apiKey = getGeminiKey_();
    if (!apiKey || !html) return '';

    // Use a modest slice to keep the prompt efficient
    var textForAi = extractMainReadableText_(html, 4000);
    if (!textForAi) return '';

    var aiTitle = callGeminiTitle_(apiKey, textForAi);
    return safeTrim_(aiTitle || '');
  } catch (e) {
    Logger.log('extractTitleFromHtml_ error: %s', e.message);
    return '';
  }
}

function callGeminiTitle_(apiKey, pageText) {
  try {
    var prompt =
      'From the page text below, return the best concise page title. ' +
      'Return only the title text with no quotes or extra words. ' +
      'Keep it under 120 characters. Prefer the main headline. ' +
      'Ignore navigation, banners, cookie notices, and boilerplate. ' +
      'Text follows:\n\n' + pageText;

    var payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}]
    };

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' +
        MODEL_NAME +
        ':generateContent?key=' +
        encodeURIComponent(apiKey),
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      var data = JSON.parse(resp.getContentText());
      var out =
        (data &&
         data.candidates &&
         data.candidates[0] &&
         data.candidates[0].content &&
         data.candidates[0].content.parts &&
         data.candidates[0].content.parts
           .map(function(p) { return p.text; })
           .filter(Boolean)
           .join(' ')) || '';
      return safeTrim_(out);
    }

    Logger.log('Gemini title call failed: %s %s', resp.getResponseCode(), resp.getContentText());
    return '';
  } catch (e) {
    Logger.log('Gemini title call error: %s', e.message);
    return '';
  }
}

// --- Gemini client ---

function getGeminiKey_() {
  var key = null; // PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return key ? key.trim() : '';
}

function callGeminiSummarize_(apiKey, url, source, text) {
  try {
    var prompt =
      'Provide a concise summary of the following article in ' + SUMMARY_SENTENCE_COUNT + ' bullet points. ' +
      'Return only the bullet points, without a title or any other introductory text. ' +
      'Base the summary solely on the article text provided. ' +
      'Article text follows:\n' + text;

    var payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }]}]
    };

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_NAME + ':generateContent?key=' + encodeURIComponent(apiKey),
      { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
    );

    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      var data = JSON.parse(resp.getContentText());
      var out =
        (data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts.map(function(p) { return p.text; }).filter(Boolean).join(' ')) || '';
      return out; // Don't normalize whitespace here, we need the newlines for bullets
    }
    Logger.log('Gemini API call failed with status code: %s', resp.getResponseCode());
    Logger.log('Gemini API response: %s', resp.getContentText());
    return '';
  } catch (e) {
    Logger.log('Gemini API call error: %s', e.message);
    return '';
  }
}

// --- Rendering ---

function renderHtmlTable_(items) {
  var tz = Session.getScriptTimeZone();
  var rows = items.map(function(it) {
    var d = Utilities.formatDate(it.emailDate, tz, 'yyyy-MM-dd HH:mm');
    var safeSource = escapeHtml_(it.source);
    var safeTitle = escapeHtml_(it.title);

    // Create a new URL block to handle the array of links
    var urlsBlock = '';
    if (Array.isArray(it.url)) {
      urlsBlock = '<ul>' + it.url.map(function(url) {
        var safeUrl = escapeHtml_(url);
        return '<li><a href="' + safeUrl + '">' + safeUrl + '</a></li>';
      }).join('\n') + '</ul>';
    } else {
      // Fallback for non-array URLs
      var safeUrl = escapeHtml_(it.url);
      urlsBlock = '<a href="' + safeUrl + '">' + safeUrl + '</a>';
    }

    // Create the bulleted list HTML from the digest text
    var bulletPoints = it.digest.split('\n')
      .filter(function(line) { return line.trim().length > 0; })
      .map(function(line) { return '<li>' + escapeHtml_(line.trim().replace(/^[\*\-•] /, '')) + '</li>'; })
      .join('\n');
    var digestHtml = '<ul>' + bulletPoints + '</ul>';

    // The new row with the added title column
    return '<tr><td style="white-space:nowrap; vertical-align:top; padding:6px 8px">' + d + '</td><td style="vertical-align:top; padding:6px 8px">' + safeSource + '</td><td style="vertical-align:top; padding:6px 8px">' + safeTitle + '</td><td style="vertical-align:top; padding:6px 8px">' + urlsBlock + '</td><td style="vertical-align:top; padding:6px 8px">' + digestHtml + '</td></tr>';
  }).join('\n');

  // The new header with the added title column
  return '<div style="font:14px system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif"><h2 style="margin:0 0 12px 0">Blogtrottr Digest</h2><table style="border-collapse:collapse; width:100%"><thead><tr><th style="text-align:left; border-bottom:1px solid #ddd; padding:8px">Email date</th><th style="text-align:left; border-bottom:1px solid #ddd; padding:8px">Source name</th><th style="text-align:left; border-bottom:1px solid #ddd; padding:8px">Post title</th><th style="text-align:left; border-bottom:1px solid #ddd; padding:8px">Link to post</th><th style="text-align:left; border-bottom:1px solid #ddd; padding:8px">Post digest</th></tr></thead><tbody>' + rows + '</tbody></table><p style="color:#666; margin-top:12px">Summaries are generated from the fetched page content only.</p></div>';
}

function renderTextTable_(items) {
  var tz = Session.getScriptTimeZone();

  var lines = items.map(function(it) {
    var d = Utilities.formatDate(it.emailDate, tz, 'yyyy-MM-dd HH:mm');

    // Flatten digest bullets to a single line
    var digestText = String(it.digest || '')
      .replace(/[\*\u2022\-]\s*/g, '')
      .replace(/\n/g, ' ')
      .trim();

    // it.url may contain multiple links joined with commas
    var urls = String(it.url || '')
      .split(/\s*,\s*/)
      .filter(Boolean);

    var urlsBlock = joinAsLines(urls);

    return 'Email date: ' + d +
           '\nSource name: ' + it.source +
           '\nPost title: ' + it.title +
           '\nLink to post(s):' + urlsBlock +
           '\nPost digest: ' + digestText;
  });

  // Blank line between emails
  return 'Blogtrottr Digest\n\n' + lines.join('\n\n') + '\n';
}

function joinAsLines(urls, opts) {
  urls = Array.isArray(urls) ? urls : [];
  opts = opts || {};
  var indent = opts.indent || '';
  var eol = opts.eol || '\n';
  var trim = opts.trim !== false; // default true

  return urls.map(function(s) {
    s = s == null ? '' : String(s);
    if (trim) s = s.trim();
    return indent + s;
  }).join(eol);
}

// --- Shared helpers ---

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function extractLinksFromHtml_(html) {
  var links = [];
  var re = /<a\s+[^>]*href=["'](https?:\/\/[^"']+)["'][^>]*>(.*?)<\/a>/gi;
  var m;
  while ((m = re.exec(html || '')) !== null) links.push(m[1]);
  return links;
}

function extractLinksFromText_(text) {
  var links = [];
  var re = /(https?:\/\/[^\s<>"']+)/g;
  var m;
  while ((m = re.exec(text || '')) !== null) links.push(m[1]);
  return links;
}

function cleanUrl_(u) {
  if (typeof u !== 'string' || u.trim() === '') {
    return u;
  }

  var parts = u.split('?');
  if (parts.length > 1) {
    var baseUrl = parts[0];
    var params = parts[1].split('&');
    var filteredParams = params.filter(function(p) {
      return !['utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'utm_content', 'utm_id', 'gclid', 'fbclid', 'mc_cid', 'mc_eid'].some(function(tracker) { return p.indexOf(tracker + '=') === 0; });
    });
    return baseUrl + (filteredParams.length > 0 ? '?' + filteredParams.join('&') : '');
  }
  return u;
}

function isLikelyArticleUrl_(u) {
  var url = (u || '').toLowerCase();
  var junk =
    url.includes('blogtrottr.com') ||
    url.includes('unsubscribe') ||
    url.includes('feedburner') ||
    url.includes('feedproxy') ||
    url.includes('feedsportal') ||
    url.includes('safelinks.protection') ||
    url.includes('facebook.com') ||
    url.includes('twitter.com') ||
    url.indexOf('mailto:') === 0 ||
    url.indexOf('.jpg') === url.length - 4 ||
    url.indexOf('.jpeg') === url.length - 5 ||
    url.indexOf('.png') === url.length - 4 ||
    url.indexOf('.gif') === url.length - 4 ||
    url.indexOf('.svg') === url.length - 4 ||
    url.indexOf('.webp') === url.length - 5 ||
    url.indexOf('.bmp') === url.length - 4;
  return !junk;
}

function tryExtractSourceFromEmail_(subject, html, url) {
  try {
    if (subject) {
      // Remove square brackets if Blogtrottr includes them
      var cleaned = subject.replace(/^\[.*?\]\s*/, '');

      // Split on common separators
      var parts = cleaned.split(/[\|\-–—:»]/);

      // Take the first part as the likely source name
      if (parts.length > 1) {
        return safeTrim_(parts[0]);
      }

      // If no separator, just return the subject itself
      return safeTrim_(cleaned);
    }

    // Fallback to hostname
    var host = getHostnameFromUrl_(url);
    return host.replace(/^www\./, '');
  } catch (e) {
    return 'unknown';
  }
}



function getHostnameFromUrl_(u) {
    if (typeof u !== 'string') return '';
    var match = u.match(/^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)/i);
    return (match && match[1]) ? match[1] : '';
}

function safeTrim_(s) {
  return (s || '').replace(/\s+/g, ' ').trim();
}

function stripHtml_(html) {
  return String(html || '').replace(/<[^>]*>/g, '');
}

function normalizeWhitespace_(s) {
  return (s || '').replace(/\s+/g, ' ').trim();
}

function truncate_(s, max) {
  if ((s || '').length <= max) return s || '';
  return (s || '').slice(0, Math.max(0, max - 1)).trimEnd() + '…';
}

function escapeHtml_(html) {
  html = String(html || '');
  return html.replace(/&/g, '&amp;')
             .replace(/</g, '&lt;')
             .replace(/>/g, '&gt;')
             .replace(/"/g, '&quot;')
             .replace(/'/g, '&#039;');
}

function simpleFallbackSummary_(text) {
  if (!text) return '';
  var sentences = text.match(/[^.!?]+[.!?]/g) || [];
  return sentences.slice(0, 2).join(' ').trim();
}
