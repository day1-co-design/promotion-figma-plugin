function doGet(e) {
  try {
    var sourceUrl = e && e.parameter && e.parameter.url ? e.parameter.url : "";
    if (!sourceUrl) {
      return jsonResponse_({ error: "Missing url parameter" });
    }

    var result = readSlidesText_(sourceUrl);
    return jsonResponse_(result);
  } catch (error) {
    return jsonResponse_({
      error: error && error.message ? error.message : String(error)
    });
  }
}

var PROXY_VERSION_ = "2026-06-12-coordinate-preview-v2";

function readSlidesText_(sourceUrl) {
  var publishedMatch = sourceUrl.match(/\/presentation(?:\/u\/\d+)?\/d\/e\/([^/]+)/i);
  if (publishedMatch) {
    var html = UrlFetchApp.fetch(sourceUrl, {
      followRedirects: true,
      muteHttpExceptions: true
    }).getContentText();
    var publishedText = extractPublishedSlidesText_(html);
    var publishedWireframeSlides = extractPublishedWireframeSlides_(html);
    if (!publishedText) {
      throw new Error("Published Slides text was not found");
    }
    return {
      text: publishedText,
      method: "published-html-apps-script",
      wireframeSlides: publishedWireframeSlides,
      wireframeDataStatus: publishedWireframeSlides && publishedWireframeSlides.length ? "coordinate" : "text-only",
      proxyVersion: PROXY_VERSION_
    };
  }

  var documentMatch = sourceUrl.match(/\/presentation(?:\/u\/\d+)?\/d\/([^/?#]+)/i);
  if (!documentMatch) {
    throw new Error("Google Slides presentation id was not found");
  }

  var documentId = documentMatch[1];
  var presentation = SlidesApp.openById(documentId);
  var slides = presentation.getSlides();
  var text = readSlidesAppTextFromPresentation_(presentation, slides);
  if (!text) {
    throw new Error("SlidesApp could not read text. Check that this Apps Script account can access the deck.");
  }

  var method = "slides-app";
  var wireframeSlides = buildWireframeSlides_(presentation, slides);
  var pptxBase64 = "";
  if (!wireframeSlides.length) {
    var previewWireframeSlides = fetchDocumentPreviewWireframeSlides_(documentId);
    if (previewWireframeSlides.length) {
      wireframeSlides = previewWireframeSlides;
      method = "slides-app+preview-svg";
    } else {
      pptxBase64 = exportPresentationPptxBase64_(documentId);
      if (pptxBase64) method = "slides-app+pptx-base64";
    }
  }
  return {
    text: text,
    method: method,
    wireframeSlides: wireframeSlides,
    wireframeDataStatus: wireframeSlides && wireframeSlides.length ? "coordinate" : (pptxBase64 ? "pptx-base64" : "text-only"),
    pptxBase64: pptxBase64,
    proxyVersion: PROXY_VERSION_
  };
}

function exportPresentationPptxBase64_(documentId) {
  var urls = [
    "https://docs.google.com/presentation/d/" + documentId + "/export/pptx",
    "https://docs.google.com/presentation/export/pptx?id=" + documentId + "&exportFormat=pptx"
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var response = UrlFetchApp.fetch(urls[i], {
        followRedirects: true,
        muteHttpExceptions: true,
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken()
        }
      });
      var blob = response.getBlob();
      var bytes = blob && blob.getBytes && blob.getBytes();
      if (bytes && bytes.length && bytes[0] === 0x50 && bytes[1] === 0x4b) {
        return Utilities.base64Encode(bytes);
      }
    } catch (error) {
      logWireframeTextOmission_("pptx-base64-export-failed", documentId, "", error);
    }
  }

  try {
    var fileBlob = DriveApp.getFileById(documentId).getAs(MimeType.MICROSOFT_POWERPOINT);
    var fileBytes = fileBlob && fileBlob.getBytes && fileBlob.getBytes();
    if (fileBytes && fileBytes.length && fileBytes[0] === 0x50 && fileBytes[1] === 0x4b) {
      return Utilities.base64Encode(fileBytes);
    }
  } catch (error) {
    logWireframeTextOmission_("drive-pptx-export-failed", documentId, "", error);
  }

  return "";
}

function fetchDocumentPreviewWireframeSlides_(documentId) {
  var urls = [
    "https://docs.google.com/presentation/d/" + documentId + "/preview",
    "https://docs.google.com/presentation/d/" + documentId + "/present"
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var response = UrlFetchApp.fetch(urls[i], {
        followRedirects: true,
        muteHttpExceptions: true
      });
      var html = response.getContentText();
      var slides = extractPublishedWireframeSlides_(html);
      if (slides && slides.length) return slides;
    } catch (error) {
      logWireframeTextOmission_("preview-svg-fallback-failed", documentId, "", error);
    }
  }
  return [];
}

function readSlidesAppText_(documentId) {
  var presentation = SlidesApp.openById(documentId);
  var slides = presentation.getSlides();
  return readSlidesAppTextFromPresentation_(presentation, slides);
}

function readSlidesAppTextFromPresentation_(presentation, slides) {
  var parts = [
    presentation.getName()
  ];

  for (var i = 0; i < slides.length; i++) {
    var slideParts = [];
    collectSlideElementsText_(slides[i].getPageElements(), slideParts);
    var notes = readSpeakerNotes_(slides[i]);
    if (notes) {
      slideParts.push("Speaker notes\n" + notes);
    }
    if (slideParts.length) {
      parts.push("Slide " + (i + 1) + "\n" + slideParts.join("\n"));
    }
  }

  return normalizeText_(parts.join("\n\n"));
}

function collectSlideElementsText_(elements, parts) {
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    var type = element.getPageElementType();

    if (type === SlidesApp.PageElementType.SHAPE) {
      addTextPart_(parts, element.asShape().getText().asString());
    } else if (type === SlidesApp.PageElementType.TABLE) {
      collectTableText_(element.asTable(), parts);
    } else if (type === SlidesApp.PageElementType.GROUP) {
      collectSlideElementsText_(element.asGroup().getChildren(), parts);
    }
  }
}

function collectTableText_(table, parts) {
  for (var row = 0; row < table.getNumRows(); row++) {
    var cells = [];
    for (var col = 0; col < table.getNumColumns(); col++) {
      cells.push(normalizeText_(table.getCell(row, col).getText().asString()));
    }
    addTextPart_(parts, cells.filter(Boolean).join(" | "));
  }
}

function readSpeakerNotes_(slide) {
  try {
    var notesPage = slide.getNotesPage();
    if (!notesPage) return "";
    var notesShape = notesPage.getSpeakerNotesShape();
    if (!notesShape) return "";
    return normalizeText_(notesShape.getText().asString());
  } catch (error) {
    return "";
  }
}

function buildWireframeSlides_(presentation, slides) {
  var slideSize = presentationSlideSize_(presentation);
  var slideData = [];

  for (var i = 0; i < slides.length; i++) {
    var textParts = [];
    var elements = [];
    collectSlideElementsText_(slides[i].getPageElements(), textParts);
    collectWireframeElements_(slides[i].getPageElements(), elements, "");
    var text = normalizeText_(textParts.join("\n"));
    slideData.push({
      number: i + 1,
      text: text,
      elements: elements
    });
  }

  var markerIndex = -1;
  for (var m = 0; m < slideData.length; m++) {
    if (isExactMainPageText_(slideData[m].text)) {
      markerIndex = m;
      break;
    }
  }

  var markerTargetSlides = markerIndex >= 0 ? slideData.slice(markerIndex + 1) : slideData;
  var targetSlides = markerTargetSlides.some(function (slide) {
    return slide.elements && slide.elements.length;
  }) ? markerTargetSlides : slideData;
  var result = [];
  for (var s = 0; s < targetSlides.length; s++) {
    var lines = wireframeTextLines_(targetSlides[s].text);
    result.push({
      number: targetSlides[s].number,
      title: lines[0] || "Slide " + targetSlides[s].number,
      text: lines.join("\n"),
      lines: lines,
      width: slideSize.width,
      height: slideSize.height,
      elements: targetSlides[s].elements
    });
  }

  return result.filter(function (slide) {
    return slide.elements && slide.elements.length;
  });
}

function collectWireframeElements_(elements, parts, groupId) {
  for (var i = 0; i < elements.length; i++) {
    try {
      collectWireframeElement_(elements[i], parts, groupId);
    } catch (error) {
      logWireframeTextOmission_("element-collect-failed", safeObjectId_(elements[i]), groupId || "", error);
    }
  }
}

function collectWireframeElement_(element, parts, groupId) {
  var type = safePageElementType_(element);
  var bounds = safeElementBounds_(element);
  var zIndex = parts.length;
  var objectId = safeObjectId_(element);
  var elementGroupId = groupId || "";
  var transform = safeElementTransform_(element);

  if (type === SlidesApp.PageElementType.GROUP) {
    var before = parts.length;
    try {
      collectWireframeElements_(element.asGroup().getChildren(), parts, objectId || elementGroupId);
    } catch (error) {
      logWireframeTextOmission_("group-children-read-failed", objectId, elementGroupId, error);
    }
    if (parts.length === before && bounds) {
      parts.push(fallbackWireframeElement_(bounds, type, objectId, elementGroupId, transform, zIndex, readPageElementText_(element)));
    }
    return;
  }

  if (!bounds) return;

  if (type === SlidesApp.PageElementType.LINE) {
    parts.push({
      type: "line",
      x: bounds.x,
      y: bounds.y,
      w: bounds.w,
      h: bounds.h,
      text: "",
      fill: "",
      stroke: "#111111",
      fontSize: 18,
      fontSizeUnit: "pt",
      zIndex: zIndex,
      objectId: objectId,
      groupId: elementGroupId,
      transform: transform
    });
    return;
  }

  if (type === SlidesApp.PageElementType.IMAGE) {
    var imageInfo = imageInfo_(element.asImage());
    parts.push({
      type: "image",
      x: bounds.x,
      y: bounds.y,
      w: bounds.w,
      h: bounds.h,
      text: "",
      fill: "#E6E6E6",
      stroke: "",
      imageDataUrl: imageInfo.imageDataUrl,
      imageUrl: imageInfo.imageUrl,
      contentUrl: imageInfo.contentUrl,
      fontSize: 18,
      fontSizeUnit: "pt",
      zIndex: zIndex,
      objectId: objectId,
      groupId: elementGroupId,
      transform: transform
    });
    return;
  }

  if (type === SlidesApp.PageElementType.TABLE) {
    var tableParts = [];
    collectTableText_(element.asTable(), tableParts);
    if (isDesignCommentText_(tableParts.join("\n"))) return;
    var tableLines = wireframeTextLines_(tableParts.join("\n"));
    parts.push({
      type: "shape",
      x: bounds.x,
      y: bounds.y,
      w: bounds.w,
      h: bounds.h,
      text: tableLines.join("\n"),
      fill: "#FFFFFF",
      stroke: "#D9D9D9",
      fontSize: 16,
      fontSizeUnit: "pt",
      zIndex: zIndex,
      objectId: objectId,
      groupId: elementGroupId,
      transform: transform,
      shapeType: "TABLE",
      textRuns: tableLines.map(function (line) {
        return { text: line, fontSize: 16, textColor: "#111111" };
      })
    });
    return;
  }

  if (type === SlidesApp.PageElementType.SHAPE) {
    var shape = element.asShape();
    var text = "";
    try {
      text = normalizeText_(shape.getText().asString());
    } catch (error) {
      logWireframeTextOmission_("shape-text-read-failed", objectId, elementGroupId, error);
      text = "";
    }
    if (isDesignCommentText_(text)) return;
    var lines = wireframeTextLines_(text);
    var fill = shapeFillHex_(shape);
    var stroke = shapeBorderHex_(shape);
    if (!lines.length && text) logWireframeTextOmission_("shape-text-empty-after-normalize", objectId, elementGroupId, text);
    if (!lines.length && !fill && !stroke) return;
    parts.push({
      type: shapeElementType_(shape, lines),
      x: bounds.x,
      y: bounds.y,
      w: bounds.w,
      h: bounds.h,
      text: lines.join("\n"),
      fill: fill,
      stroke: stroke,
      textColor: shapeTextColorHex_(shape),
      fontSize: shapeFontSize_(shape),
      fontSizeUnit: "pt",
      align: shapeTextAlign_(shape),
      verticalAlign: shapeTextVerticalAlign_(shape),
      zIndex: zIndex,
      objectId: objectId,
      groupId: elementGroupId,
      transform: transform,
      shapeType: safeShapeType_(shape),
      textRuns: shapeTextRuns_(shape),
      radius: shapeRadius_(shape)
    });
    return;
  }

  parts.push(fallbackWireframeElement_(bounds, type, objectId, elementGroupId, transform, zIndex, readPageElementText_(element)));
}

function fallbackWireframeElement_(bounds, pageElementType, objectId, groupId, transform, zIndex, rawText) {
  var lines = wireframeTextLines_(rawText || "");
  return {
    type: lines.length ? "text" : "shape",
    x: bounds.x,
    y: bounds.y,
    w: bounds.w,
    h: bounds.h,
    text: lines.join("\n"),
    fill: lines.length ? "" : "#E6E6E6",
    stroke: lines.length ? "" : "#D9D9D9",
    textColor: "#111111",
    fontSize: 18,
    fontSizeUnit: "pt",
    zIndex: zIndex,
    objectId: objectId,
    groupId: groupId,
    transform: transform,
    shapeType: String(pageElementType || "UNKNOWN")
  };
}

function safePageElementType_(element) {
  try {
    return element.getPageElementType();
  } catch (error) {
    return "UNKNOWN";
  }
}

function readPageElementText_(element) {
  var readers = [
    function () {
      return element.asShape().getText().asString();
    },
    function () {
      var wordArt = element.asWordArt && element.asWordArt();
      return wordArt && wordArt.getText ? wordArt.getText() : "";
    },
    function () {
      return element.getTitle && element.getTitle();
    },
    function () {
      return element.getDescription && element.getDescription();
    }
  ];
  for (var i = 0; i < readers.length; i++) {
    try {
      var text = normalizeText_(readers[i]() || "");
      if (text) return text;
    } catch (error) {
    }
  }
  return "";
}

function imageInfo_(image) {
  var result = { imageDataUrl: "", imageUrl: "", contentUrl: "" };
  try {
    result.imageUrl = String(image.getSourceUrl && image.getSourceUrl() || "");
  } catch (error) {
    result.imageUrl = "";
  }
  try {
    result.contentUrl = String(image.getContentUrl && image.getContentUrl() || "");
  } catch (error) {
    result.contentUrl = "";
  }
  try {
    var blob = image.getBlob && image.getBlob();
    var contentType = blob && blob.getContentType && blob.getContentType();
    var bytes = blob && blob.getBytes && blob.getBytes();
    if (contentType && bytes && bytes.length) {
      result.imageDataUrl = "data:" + contentType + ";base64," + Utilities.base64Encode(bytes);
    }
  } catch (error) {
    result.imageDataUrl = "";
  }
  return result;
}

function logWireframeTextOmission_(reason, objectId, groupId, detail) {
  try {
    console.log("[wireframe] text omitted: " + JSON.stringify({
      reason: reason,
      objectId: objectId || "",
      groupId: groupId || "",
      detail: String(detail && detail.message ? detail.message : detail || "").slice(0, 160)
    }));
  } catch (error) {
  }
}

function safeObjectId_(element) {
  try {
    return String(element.getObjectId ? element.getObjectId() : "");
  } catch (error) {
    return "";
  }
}

function safeElementTransform_(element) {
  try {
    var transform = element.getTransform && element.getTransform();
    if (!transform) return null;
    return {
      scaleX: safeNumber_(transform.getScaleX && transform.getScaleX()),
      scaleY: safeNumber_(transform.getScaleY && transform.getScaleY()),
      shearX: safeNumber_(transform.getShearX && transform.getShearX()),
      shearY: safeNumber_(transform.getShearY && transform.getShearY()),
      translateX: safeNumber_(transform.getTranslateX && transform.getTranslateX()),
      translateY: safeNumber_(transform.getTranslateY && transform.getTranslateY()),
      unit: String(transform.getUnit && transform.getUnit() || "")
    };
  } catch (error) {
    return null;
  }
}

function shapeElementType_(shape, lines) {
  var shapeType = safeShapeType_(shape);
  if (/ELLIPSE|ARC|PIE|DONUT/i.test(shapeType)) return "ellipse";
  return lines.length ? "text" : "shape";
}

function safeShapeType_(shape) {
  try {
    return String(shape.getShapeType());
  } catch (error) {
    return "";
  }
}

function safeElementBounds_(element) {
  try {
    var x = safeNumber_(element.getLeft());
    var y = safeNumber_(element.getTop());
    var w = safeNumber_(element.getWidth());
    var h = safeNumber_(element.getHeight());
    if (w <= 0 && h <= 0) return null;
    if (w <= 0) w = 1;
    if (h <= 0) h = 1;
    return { x: x, y: y, w: w, h: h };
  } catch (error) {
    return null;
  }
}

function presentationSlideSize_(presentation) {
  try {
    return {
      width: safeNumber_(presentation.getPageWidth()) || 960,
      height: safeNumber_(presentation.getPageHeight()) || 540
    };
  } catch (error) {
    return { width: 960, height: 540 };
  }
}

function shapeFillHex_(shape) {
  try {
    return colorToHex_(shape.getFill().getSolidFill().getColor());
  } catch (error) {
    return "";
  }
}

function shapeBorderHex_(shape) {
  try {
    return colorToHex_(shape.getBorder().getLineFill().getSolidFill().getColor());
  } catch (error) {
    return "";
  }
}

function colorToHex_(color) {
  try {
    var rgb = color.asRgbColor();
    if (rgb && rgb.asHexString) return rgb.asHexString();
  } catch (error) {
    return "";
  }
  return "";
}

function shapeFontSize_(shape) {
  try {
    var size = shape.getText().getTextStyle().getFontSize();
    return safeNumber_(size) || 18;
  } catch (error) {
    return 18;
  }
}

function shapeTextColorHex_(shape) {
  try {
    return colorToHex_(shape.getText().getTextStyle().getForegroundColor());
  } catch (error) {
    return "";
  }
}

function shapeTextAlign_(shape) {
  try {
    var alignment = String(shape.getText().getParagraphStyle().getParagraphAlignment() || "");
    if (/CENTER/i.test(alignment)) return "CENTER";
    if (/END|RIGHT/i.test(alignment)) return "RIGHT";
    return "LEFT";
  } catch (error) {
    return "";
  }
}

function shapeTextVerticalAlign_(shape) {
  try {
    var alignment = String(shape.getContentAlignment && shape.getContentAlignment() || "");
    if (/MIDDLE|CENTER/i.test(alignment)) return "CENTER";
    if (/BOTTOM/i.test(alignment)) return "BOTTOM";
    if (/TOP/i.test(alignment)) return "TOP";
  } catch (error) {
  }
  return "";
}

function shapeTextRuns_(shape) {
  var textRange;
  try {
    textRange = shape.getText();
  } catch (error) {
    return [];
  }

  var runs = [];
  try {
    var textRuns = textRange.getRuns ? textRange.getRuns() : [];
    for (var i = 0; i < textRuns.length; i++) {
      var runText = normalizeText_(textRuns[i].asString());
      if (!runText) continue;
      var style = textRuns[i].getTextStyle();
      runs.push({
        text: runText,
        fontSize: safeNumber_(style.getFontSize && style.getFontSize()) || shapeFontSize_(shape),
        textColor: colorToHex_(style.getForegroundColor && style.getForegroundColor()),
        bold: Boolean(style.isBold && style.isBold())
      });
    }
  } catch (error) {
    runs = [];
  }

  if (!runs.length) {
    try {
      var fullText = normalizeText_(textRange.asString());
      if (fullText) {
        runs.push({
          text: fullText,
          fontSize: shapeFontSize_(shape),
          textColor: shapeTextColorHex_(shape),
          bold: false
        });
      }
    } catch (error) {
    }
  }
  return runs;
}

function shapeRadius_(shape) {
  try {
    return /ROUND|round/i.test(String(shape.getShapeType())) ? "round" : "";
  } catch (error) {
    return "";
  }
}

function wireframeTextLines_(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u2028|\u2029/g, "\n")
    .split("\n")
    .map(function (line) {
      return line.replace(/[ \t]+/g, " ").trim();
    })
    .filter(function (line) {
      return line && !isNonHeroText_(line);
    });
}

function isExactMainPageText_(value) {
  var lines = String(value || "")
    .split(/\n+/)
    .map(function (line) {
      return line.replace(/\s+/g, "").trim();
    })
    .filter(Boolean)
    .filter(function (line) {
      return !/^slide\d+$/i.test(line);
    });
  if (!lines.length) return false;
  return lines[0] === "메인페이지" || /^mainpage$/i.test(lines[0]);
}

function isNonHeroText_(value) {
  return /디자인\s*코멘트|코멘트|comment|디자인\s*방향|기획\s*방향|비주얼\s*방향|작업\s*가이드|가이드|설명|참고|레퍼런스|확인\s*필요/i.test(String(value || ""));
}

function isDesignCommentText_(value) {
  return /디자인\s*코멘트|design\s*comment/i.test(String(value || ""));
}

function safeNumber_(value) {
  var number = Number(value);
  return isFinite(number) ? number : 0;
}

function extractPublishedSlidesText_(html) {
  var source = String(html || "");
  var parts = [];
  var seen = {};
  addUniquePart_(parts, seen, extractPublishedMetaText_(source));

  var actualLabelPattern = /\saria-label=(["'])([\s\S]*?)\1/gi;
  var actualMatch;
  while ((actualMatch = actualLabelPattern.exec(source))) {
    addUniquePart_(parts, seen, decodePublishedAttribute_(actualMatch[2]));
  }

  var escapedLabelPattern = /aria-label\\x3d\\x22([\s\S]*?)\\x22/gi;
  var escapedMatch;
  while ((escapedMatch = escapedLabelPattern.exec(source))) {
    addUniquePart_(parts, seen, decodePublishedAttribute_(escapedMatch[1]));
  }

  return normalizeText_(parts.filter(function (part) {
    return /[가-힣A-Za-z0-9]/.test(part) && !/Google Slides|자바스크립트가 브라우저/i.test(part);
  }).join("\n\n"));
}

function extractPublishedWireframeSlides_(html) {
  var svgSlides = extractPublishedSvgSlides_(html);
  if (!svgSlides.length) return [];

  var docPages = extractPublishedDocDataPages_(html);
  var pageTextById = {};
  var pageOrderById = {};
  for (var i = 0; i < docPages.length; i++) {
    pageTextById[docPages[i].pageId] = docPages[i].text;
    pageOrderById[docPages[i].pageId] = i;
  }

  if (docPages.length) {
    svgSlides.sort(function (a, b) {
      var orderA = Object.prototype.hasOwnProperty.call(pageOrderById, a.pageId) ? pageOrderById[a.pageId] : 9999;
      var orderB = Object.prototype.hasOwnProperty.call(pageOrderById, b.pageId) ? pageOrderById[b.pageId] : 9999;
      return orderA - orderB;
    });
  }

  return svgSlides.map(function (slide, index) {
    var text = pageTextById[slide.pageId] || extractTextFromSvg_(slide.svg);
    var lines = wireframeTextLines_(text);
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
        svg: slide.svg
      }]
    };
  });
}

function extractPublishedSvgSlides_(html) {
  var source = String(html || "");
  var slides = [];
  var pattern = /SK_svgData\s*=\s*'([\s\S]*?)';[\s\S]*?SK_viewerApp\.setPageData\('([^']+)'/g;
  var match;
  while ((match = pattern.exec(source))) {
    var svg = decodePublishedJsString_(match[1]);
    if (svg.indexOf("<svg") < 0) continue;
    var size = svgViewBoxSize_(svg);
    slides.push({
      pageId: match[2],
      svg: sanitizePublishedSvg_(svg),
      width: size.width,
      height: size.height
    });
  }
  return slides;
}

function extractPublishedDocDataPages_(html) {
  var source = String(html || "");
  var marker = source.indexOf("docData:");
  if (marker < 0) return [];
  var start = source.indexOf("[", marker);
  if (start < 0) return [];
  var snippet = extractBalancedJsArray_(source, start);
  if (!snippet) return [];

  try {
    var docData = Function("\"use strict\"; return (" + snippet + ");")();
    var pages = docData && Array.isArray(docData[1]) ? docData[1] : [];
    return pages.map(function (page) {
      return {
        pageId: String(page && page[0] ? page[0] : ""),
        text: normalizeText_(page && page[2] ? page[2] : "")
      };
    }).filter(function (page) {
      return page.pageId;
    });
  } catch (error) {
    return [];
  }
}

function extractBalancedJsArray_(source, start) {
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

function svgViewBoxSize_(svg) {
  var match = String(svg || "").match(/viewBox=(["'])([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\1/i);
  if (match) {
    return {
      width: safeNumber_(match[4]) || 960,
      height: safeNumber_(match[5]) || 540
    };
  }
  return { width: 960, height: 540 };
}

function extractTextFromSvg_(svg) {
  var source = String(svg || "");
  var parts = [];
  var seen = {};
  var pattern = /\saria-label=(["'])([\s\S]*?)\1/gi;
  var match;
  while ((match = pattern.exec(source))) {
    addUniquePart_(parts, seen, decodeXml_(match[2]));
  }
  return normalizeText_(parts.join("\n"));
}

function sanitizePublishedSvg_(svg) {
  return String(svg || "")
    .replace(/<script\b[\s\S]*?<\/script>/gi, "")
    .replace(/\s+on[a-z]+\s*=\s*(["'])[\s\S]*?\1/gi, "");
}

function decodePublishedJsString_(value) {
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

function extractPublishedMetaText_(html) {
  var fields = [];
  var metaPattern = /<meta[^>]+(?:property|name)=(["'])(og:title|og:description|description)\1[^>]+content=(["'])([\s\S]*?)\2[^>]*>/gi;
  var match;
  while ((match = metaPattern.exec(html))) {
    fields.push(decodePublishedAttribute_(match[3]));
  }
  return normalizeText_(fields.join("\n"));
}

function decodePublishedAttribute_(value) {
  return decodeXml_(String(value || "")
    .replace(/\\x([0-9a-f]{2})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\u([0-9a-f]{4})/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/\\\//g, "/"));
}

function decodeXml_(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#x([0-9a-f]+);/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/&#(\d+);/g, function (_, number) {
      return String.fromCharCode(parseInt(number, 10));
    });
}

function addTextPart_(parts, value) {
  var text = normalizeText_(value);
  if (text) parts.push(text);
}

function addUniquePart_(parts, seen, value) {
  var text = normalizeText_(value);
  if (!text || text.length < 3 || seen[text]) return;
  seen[text] = true;
  parts.push(text);
}

function normalizeText_(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u2028|\u2029/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
