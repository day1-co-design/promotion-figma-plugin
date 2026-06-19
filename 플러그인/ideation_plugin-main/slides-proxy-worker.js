export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    try {
      var requestUrl = new URL(request.url);
      var sourceUrl = requestUrl.searchParams.get("url") || "";
      if (!sourceUrl) {
        return jsonResponse({ error: "Missing url parameter" }, 400);
      }

      var parsedSource = new URL(sourceUrl);
      if (parsedSource.hostname !== "docs.google.com" || !/^\/presentation\//.test(parsedSource.pathname)) {
        return jsonResponse({ error: "Only Google Slides URLs are allowed" }, 400);
      }

      var result = await readSlidesText(parsedSource);
      return jsonResponse(result, 200);
    } catch (error) {
      return jsonResponse({ error: error && error.message ? error.message : String(error) }, 500);
    }
  }
};

const PROXY_VERSION = "2026-06-12-coordinate-preview-v2";

async function readSlidesText(parsedSource) {
  var publicMatch = parsedSource.pathname.match(/\/presentation(?:\/u\/\d+)?\/d\/e\/([^/]+)/i);
  if (publicMatch) {
    var html = await fetchText(parsedSource.toString());
    var publishedText = extractPublishedSlidesText(html);
    var publishedWireframeSlides = extractPublishedWireframeSlides(html);
    if (!publishedText) throw new Error("Published Slides text was not found");
    return {
      text: publishedText,
      method: "published-html-proxy",
      wireframeSlides: publishedWireframeSlides,
      wireframeDataStatus: publishedWireframeSlides.length ? "coordinate" : "text-only",
      proxyVersion: PROXY_VERSION
    };
  }

  var documentMatch = parsedSource.pathname.match(/\/presentation(?:\/u\/\d+)?\/d\/([^/]+)/i);
  if (!documentMatch) {
    throw new Error("Google Slides presentation id was not found");
  }

  var documentId = documentMatch[1];
  var exportUrl = "https://docs.google.com/presentation/d/" + documentId + "/export/txt";
  var exportedText = normalizeFetchedText(await fetchText(exportUrl));
  if (!exportedText || looksLikeAccessPage(exportedText)) {
    throw new Error("Slides text export was blocked or empty");
  }
  var pptxBase64 = await fetchPptxBase64(documentId);

  return {
    text: exportedText,
    method: pptxBase64 ? "text-export-proxy+pptx-base64" : "text-export-proxy",
    wireframeSlides: [],
    wireframeDataStatus: pptxBase64 ? "pptx-base64" : "text-only",
    pptxBase64,
    proxyVersion: PROXY_VERSION
  };
}

async function fetchPptxBase64(documentId) {
  var urls = [
    "https://docs.google.com/presentation/d/" + documentId + "/export/pptx",
    "https://docs.google.com/presentation/export/pptx?id=" + documentId + "&exportFormat=pptx"
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var response = await fetch(urls[i], {
        redirect: "follow",
        headers: {
          "user-agent": "Mozilla/5.0 SlidesProxy/1.0"
        }
      });
      if (!response.ok) continue;
      var buffer = await response.arrayBuffer();
      var bytes = new Uint8Array(buffer);
      if (bytes.length && bytes[0] === 0x50 && bytes[1] === 0x4b) {
        return arrayBufferToBase64(buffer);
      }
    } catch (error) {
    }
  }
  return "";
}

function arrayBufferToBase64(buffer) {
  var bytes = new Uint8Array(buffer);
  var parts = [];
  var chunkSize = 0x8000;
  for (var i = 0; i < bytes.length; i += chunkSize) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize)));
  }
  return btoa(parts.join(""));
}

async function fetchText(url) {
  var response = await fetch(url, {
    redirect: "follow",
    headers: {
      "user-agent": "Mozilla/5.0 SlidesProxy/1.0"
    }
  });
  if (!response.ok) {
    throw new Error("Fetch failed: " + response.status);
  }
  return response.text();
}

function extractPublishedSlidesText(html) {
  var source = String(html || "");
  var parts = [];
  var seen = {};
  var metaText = extractPublishedMetaText(source);
  if (metaText) addUniquePart(parts, seen, metaText);

  var actualLabelPattern = /\saria-label=(["'])([\s\S]*?)\1/gi;
  var actualMatch;
  while ((actualMatch = actualLabelPattern.exec(source))) {
    addUniquePart(parts, seen, decodePublishedAttribute(actualMatch[2]));
  }

  var escapedLabelPattern = /aria-label\\x3d\\x22([\s\S]*?)\\x22/gi;
  var escapedMatch;
  while ((escapedMatch = escapedLabelPattern.exec(source))) {
    addUniquePart(parts, seen, decodePublishedAttribute(escapedMatch[1]));
  }

  var readableParts = parts.filter(function (part) {
    return /[가-힣A-Za-z0-9]/.test(part) && !/Google Slides|자바스크립트가 브라우저/i.test(part);
  });

  return normalizeFetchedText(readableParts.join("\n\n"));
}

function extractPublishedWireframeSlides(html) {
  var svgSlides = extractPublishedSvgSlides(html);
  if (!svgSlides.length) return [];

  var docPages = extractPublishedDocDataPages(html);
  var pageTextById = {};
  var pageOrderById = {};
  docPages.forEach(function (page, index) {
    pageTextById[page.pageId] = page.text;
    pageOrderById[page.pageId] = index;
  });

  if (docPages.length) {
    svgSlides.sort(function (a, b) {
      var orderA = Object.prototype.hasOwnProperty.call(pageOrderById, a.pageId) ? pageOrderById[a.pageId] : 9999;
      var orderB = Object.prototype.hasOwnProperty.call(pageOrderById, b.pageId) ? pageOrderById[b.pageId] : 9999;
      return orderA - orderB;
    });
  }

  return svgSlides.map(function (slide, index) {
    var text = pageTextById[slide.pageId] || extractTextFromSvg(slide.svg);
    var lines = wireframeTextLines(text);
    return {
      number: index + 1,
      title: lines[0] || "Slide " + (index + 1),
      text: lines.join("\n"),
      lines: lines,
      width: slide.width || 960,
      height: slide.height || 540,
      elements: [{
        type: "svg",
        x: 0,
        y: 0,
        w: slide.width || 960,
        h: slide.height || 540,
        svg: slide.svg,
        zIndex: 0
      }]
    };
  });
}

function extractPublishedSvgSlides(html) {
  var source = String(html || "");
  var slides = [];
  var pattern = /SK_svgData\s*=\s*'([\s\S]*?)';[\s\S]*?SK_viewerApp\.setPageData\('([^']+)'/g;
  var match;
  while ((match = pattern.exec(source))) {
    var svg = decodePublishedJsString(match[1]);
    if (svg.indexOf("<svg") < 0) continue;
    var size = svgViewBoxSize(svg);
    slides.push({
      pageId: match[2],
      svg: sanitizePublishedSvg(svg),
      width: size.width,
      height: size.height
    });
  }
  return slides;
}

function extractPublishedDocDataPages(html) {
  var source = String(html || "");
  var marker = source.indexOf("docData:");
  if (marker < 0) return [];
  var start = source.indexOf("[", marker);
  if (start < 0) return [];
  var snippet = extractBalancedJsArray(source, start);
  if (!snippet) return [];

  try {
    var docData = Function("\"use strict\"; return (" + snippet + ");")();
    var pages = docData && Array.isArray(docData[1]) ? docData[1] : [];
    return pages.map(function (page) {
      return {
        pageId: String(page && page[0] ? page[0] : ""),
        text: normalizeFetchedText(page && page[2] ? page[2] : "")
      };
    }).filter(function (page) {
      return page.pageId;
    });
  } catch (error) {
    return [];
  }
}

function extractBalancedJsArray(source, start) {
  var depth = 0;
  var quote = "";
  var escaped = false;
  for (var i = start; i < source.length; i++) {
    var char = source[i];
    if (quote) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === quote) {
        quote = "";
      }
      continue;
    }
    if (char === "\"" || char === "'") {
      quote = char;
      continue;
    }
    if (char === "[") {
      depth++;
    } else if (char === "]") {
      depth--;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  return "";
}

function svgViewBoxSize(svg) {
  var match = String(svg || "").match(/viewBox=(["'])([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\1/i);
  if (match) {
    return {
      width: Number(match[4]) || 960,
      height: Number(match[5]) || 540
    };
  }
  return { width: 960, height: 540 };
}

function extractTextFromSvg(svg) {
  var source = String(svg || "");
  var parts = [];
  var seen = {};
  var pattern = /\saria-label=(["'])([\s\S]*?)\1/gi;
  var match;
  while ((match = pattern.exec(source))) {
    addUniquePart(parts, seen, decodeXml(match[2]));
  }
  return normalizeFetchedText(parts.join("\n"));
}

function sanitizePublishedSvg(svg) {
  return String(svg || "")
    .replace(/<script\b[\s\S]*?<\/script>/gi, "")
    .replace(/\s+on[a-z]+\s*=\s*(["'])[\s\S]*?\1/gi, "");
}

function decodePublishedJsString(value) {
  return String(value || "")
    .replace(/\\x([0-9a-f]{2})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\u([0-9a-f]{4})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\\//g, "/")
    .replace(/\\'/g, "'")
    .replace(/\\"/g, "\"")
    .replace(/\\n/g, "\n")
    .replace(/\\r/g, "\r")
    .replace(/\\t/g, "\t")
    .replace(/\\\\/g, "\\");
}

function wireframeTextLines(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u2028|\u2029/g, "\n")
    .split("\n")
    .map(function (line) {
      return line.replace(/[ \t]+/g, " ").trim();
    })
    .filter(Boolean);
}

function extractPublishedMetaText(html) {
  var fields = [];
  var metaPattern = /<meta[^>]+(?:property|name)=(["'])(og:title|og:description|description)\1[^>]+content=(["'])([\s\S]*?)\2[^>]*>/gi;
  var match;
  while ((match = metaPattern.exec(html))) {
    fields.push(decodePublishedAttribute(match[3]));
  }
  return normalizeFetchedText(fields.join("\n"));
}

function addUniquePart(parts, seen, value) {
  var text = normalizeFetchedText(value);
  if (!text || text.length < 3) return;
  if (seen[text]) return;
  seen[text] = true;
  parts.push(text);
}

function decodePublishedAttribute(value) {
  return decodeXml(String(value || "")
    .replace(/\\x([0-9a-f]{2})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\u([0-9a-f]{4})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\\//g, "/"));
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#x([0-9a-f]+);/gi, function (_, hex) { return String.fromCharCode(parseInt(hex, 16)); })
    .replace(/&#(\d+);/g, function (_, number) { return String.fromCharCode(parseInt(number, 10)); });
}

function looksLikeAccessPage(text) {
  return /<html|로그인|권한|액세스|access denied|sign in|request access|doctype html/i.test(text);
}

function normalizeFetchedText(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u2028|\u2029/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function jsonResponse(value, status) {
  return new Response(JSON.stringify(value), {
    status: status,
    headers: {
      ...corsHeaders(),
      "content-type": "application/json; charset=utf-8"
    }
  });
}

function corsHeaders() {
  return {
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET, OPTIONS",
    "access-control-allow-headers": "content-type"
  };
}
