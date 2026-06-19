figma.showUI(__html__, { width: 430, height: 720, themeColors: true });

var fontRegular = { family: "Inter", style: "Regular" };
var fontMedium = { family: "Inter", style: "Medium" };
var fontBold = { family: "Inter", style: "Bold" };
var availableFontsByFamily = {};
var DEFAULT_PINTEREST_BOARD_URL = "https://pin.it/3HCSXMMfK";

var WIRE_ERROR = {
  GOOGLE_SLIDES_ACCESS_BLOCKED: "GOOGLE_SLIDES_ACCESS_BLOCKED",
  GOOGLE_SLIDES_EXPORT_FAILED: "GOOGLE_SLIDES_EXPORT_FAILED",
  SLIDES_PROXY_TEXT_ONLY: "SLIDES_PROXY_TEXT_ONLY",
  PPTX_COORDINATE_EXTRACTION_FAILED: "PPTX_COORDINATE_EXTRACTION_FAILED",
  PPTX_UPLOAD_INVALID: "PPTX_UPLOAD_INVALID",
  WIREFRAME_COORDINATE_MISSING: "WIREFRAME_COORDINATE_MISSING"
};
// OUTPUT_POLICY_START
var WIRE_OUTPUT_POLICY = {
  "version": "2026-06-19-output-boundary-v1",
  "targetWidthPx": 1080,
  "layoutModes": {
    "exact": "원본 좌표/SVG 기반 레이아웃",
    "compat": "구버전 프록시 호환 목업"
  },
  "compat": {
    "enabled": true,
    "status": "compat-text-fallback",
    "methodSuffix": "compat-layout",
    "slidesWarning": "구버전 프록시가 좌표를 반환하지 않아 호환 목업으로 생성했습니다.",
    "pptxWarning": "PPTX 좌표 요소를 찾지 못해 호환 목업으로 생성했습니다."
  },
  "messages": {
    "coordinateMissing": "원본 슬라이드의 위치 좌표/SVG 구조를 읽지 못해 와이어프레임 생성을 중단했습니다.\nGoogle Slides 링크 권한, Apps Script 프록시, 또는 PPTX 업로드 파일을 확인해 주세요."
  },
  "styles": {
    "rootBackground": "#FFFFFF",
    "canvasBackground": "#FFFFFF",
    "placeholderFill": "#D9D9D9",
    "placeholderOpacity": 0.82,
    "imageOpacity": 0.42,
    "rootTopGapPx": 60,
    "slideGapPx": 96,
    "rootBottomGapPx": 60
  }
};
// OUTPUT_POLICY_END

var GoogleSlidesLinkProcessingModule = {
  readForAnalysis: function (slidesInfo) {
    return extractGoogleSlidesText(slidesInfo);
  },
  convertForWireframe: function (slidesInfo) {
    return convertGoogleSlidesToPptxText(slidesInfo);
  }
};

var PptxUploadProcessingModule = {
  extractDeck: async function (bytes, fileName) {
    var uploadedBuffer = messageBytesToArrayBuffer(bytes);
    assertPptxBuffer(uploadedBuffer, fileName || "업로드한 파일");
    return await extractDeckFromExportBuffer(uploadedBuffer);
  }
};

var PptxCoordinateExtractionModule = {
  extractSlides: function (slideData, slideSize, entries, buffer) {
    return extractWireframeSlides(slideData, slideSize, entries, buffer);
  },
  hasCoordinateSlides: function (slides) {
    return hasVisualWireframeSlides(slides);
  }
};

var WireframeRenderingModule = {
  normalizeDeck: function (deck) {
    return normalizeWireframeDeck(deck);
  },
  createFrame: function (deck, page) {
    return createWireframeFrame(deck, page);
  },
  createCanvas: function (slide, index) {
    return createWireframeCanvasFromElements(slide, index);
  }
};

figma.ui.onmessage = async function (message) {
  if (message.type === "resize-ui") {
    figma.ui.resize(message.width, message.height);
  }

  if (message.type === "create-analysis") {
    try {
      await prepareFonts();
      var page = await createAnalysisPage(message.analysis);
      var frame = await createAnalysisFrame(message.analysis, page);
      await activatePage(page);
      page.selection = [frame];
      figma.viewport.scrollAndZoomIntoView([frame]);
      figma.notify("새 페이지에 프로모션 분석안이 생성되었습니다.");
      figma.ui.postMessage({ type: "created", pageName: page.name });
    } catch (error) {
      figma.notify("생성 중 오류가 발생했습니다: " + error.message, { error: true });
      figma.ui.postMessage(errorPayload("error", null, error));
    }
  }

  if (message.type === "create-wireframe") {
    try {
      await prepareFonts();
      var wireframeDeck = WireframeRenderingModule.normalizeDeck(message.deck || {});
      var wireframePage = await createWireframePage(wireframeDeck);
      var wireframeFrame = await WireframeRenderingModule.createFrame(wireframeDeck, wireframePage);
      await activatePage(wireframePage);
      wireframePage.selection = [wireframeFrame];
      figma.viewport.scrollAndZoomIntoView([wireframeFrame]);
      figma.notify("새 페이지에 기획안 와이어프레임이 생성되었습니다.");
      figma.ui.postMessage({
        type: "wireframe-created",
        pageName: wireframePage.name,
        slideCount: wireframeDeck.slides.length
      });
    } catch (error) {
      figma.notify("와이어프레임 생성 중 오류가 발생했습니다: " + error.message, { error: true });
      figma.ui.postMessage(errorPayload("error", null, error));
    }
  }

  if (message.type === "fetch-slides") {
    try {
      var result = await GoogleSlidesLinkProcessingModule.readForAnalysis(message.slidesInfo);
      figma.ui.postMessage({
        type: "slides-text",
        requestId: message.requestId,
        text: result.text,
        sourceName: result.sourceName,
        referenceSlideHeroText: result.referenceSlideHeroText || "",
        wireframeSlides: result.wireframeSlides || [],
        wireframeDataStatus: result.wireframeDataStatus || "",
        proxyVersion: result.proxyVersion || "",
        viaProxy: result.viaProxy,
        method: result.method
      });
    } catch (error) {
      figma.ui.postMessage(errorPayload("slides-fetch-error", message.requestId, error));
    }
  }

  if (message.type === "convert-slides-to-pptx") {
    try {
      var converted = await GoogleSlidesLinkProcessingModule.convertForWireframe(message.slidesInfo);
      figma.ui.postMessage({
        type: "slides-converted",
        requestId: message.requestId,
        text: converted.text,
        sourceName: converted.sourceName,
        referenceSlideHeroText: converted.referenceSlideHeroText || "",
        wireframeSlides: converted.wireframeSlides || [],
        wireframeDataStatus: converted.wireframeDataStatus || "",
        proxyVersion: converted.proxyVersion || "",
        viaProxy: converted.viaProxy,
        method: converted.method
      });
    } catch (error) {
      figma.ui.postMessage(errorPayload("slides-convert-error", message.requestId, error));
    }
  }

  if (message.type === "extract-uploaded-pptx") {
    try {
      var uploadedDeck = await PptxUploadProcessingModule.extractDeck(message.bytes, message.fileName || "업로드한 파일");
      figma.ui.postMessage({
        type: "pptx-uploaded",
        requestId: message.requestId,
        text: uploadedDeck.text,
        referenceSlideImages: uploadedDeck.referenceSlideImages,
        referenceSlideHeroText: uploadedDeck.referenceSlideHeroText || "",
        wireframeSlides: uploadedDeck.wireframeSlides || [],
        wireframeDataStatus: hasVisualWireframeSlides(uploadedDeck.wireframeSlides) ? "coordinate" : "text-only",
        proxyVersion: "",
        sourceName: message.fileName || "업로드한 PPTX",
        viaProxy: false,
        method: "pptx-upload"
      });
    } catch (error) {
      figma.ui.postMessage(errorPayload("pptx-upload-error", message.requestId, error));
    }
  }

  if (message.type === "close") {
    figma.closePlugin();
  }
};

function createPluginError(code, message, details) {
  var error = new Error(message);
  error.code = code || "";
  error.details = details || "";
  return error;
}

function errorPayload(type, requestId, error) {
  return {
    type: type,
    requestId: requestId || "",
    code: error && error.code || "",
    message: error && error.message ? error.message : String(error || "오류가 발생했습니다."),
    details: error && error.details || ""
  };
}

function isGoogleSlidesAccessBlockedError(errorOrMessage) {
  var message = String(errorOrMessage && errorOrMessage.message ? errorOrMessage.message : errorOrMessage || "");
  var code = String(errorOrMessage && errorOrMessage.code || "");
  return code === WIRE_ERROR.GOOGLE_SLIDES_ACCESS_BLOCKED ||
    /failed to fetch|load failed|networkerror|cors|보안 정책|로그인|권한|access denied|sign in|request access|접근 안내 페이지|PPTX 파일이 아니거나|403|401/i.test(message);
}

function firstErrorMessage(errors) {
  for (var i = 0; i < errors.length; i++) {
    var message = errors[i] && errors[i].message ? errors[i].message : String(errors[i] || "");
    if (message) return message;
  }
  return "";
}

async function createAnalysisPage(rawAnalysis) {
  var analysis = normalizeAnalysis(rawAnalysis);
  var page = figma.createPage();
  page.name = "프로모션 분석안 - " + analysis.promotionName;
  return page;
}

async function createWireframePage(deck) {
  var page = figma.createPage();
  page.name = "기획안 와이어프레임 - " + truncatePageName(deck.sourceName || "기획안");
  return page;
}

function normalizeWireframeDeck(rawDeck) {
  var deck = rawDeck || {};
  var slides = safeArray(deck.slides, []).map(function (slide, index) {
    var lines = wireframeTextLines(slide && (slide.lines || slide.text || ""));
    var title = slide && slide.title && !isNonHeroText(slide.title) ? slide.title : "";
    var slideNumber = Number(slide && slide.number) || index + 1;
    var sourceElements = safeArray(slide && slide.normalizedSlideElements, safeArray(slide && slide.elements, []));
    var normalizedSlideElements = sourceElements.map(function (element, elementIndex) {
      return normalizeWireframeElement(element, slideNumber, elementIndex);
    }).filter(Boolean);
    return {
      number: slideNumber,
      title: String(lines[0] || title || "Slide " + (index + 1)),
      text: lines.join("\n"),
      lines: lines,
      width: Number(slide && slide.width) || 12192000,
      height: Number(slide && slide.height) || 6858000,
      normalizedSlideElements: normalizedSlideElements,
      elements: normalizedSlideElements
    };
  }).filter(function (slide) {
    return slide.normalizedSlideElements.length;
  });

  if (!slides.length) {
    throw new Error("원본 슬라이드의 위치 좌표를 읽지 못해 와이어프레임 생성을 중단했습니다. Google Slides 링크 권한, Apps Script 프록시, 또는 PPTX 업로드 파일을 확인해 주세요.");
  }

  return {
    sourceName: deck.sourceName || "기획안",
    sourceUrl: deck.sourceUrl || "",
    method: deck.method || "",
    slides: slides
  };
}

function wireframeTextLines(value) {
  var rawLines = Array.isArray(value) ? value : String(value || "").split(/\n/);
  return rawLines.map(function (line) {
    return cleanWireframeOutputLine(String(line || "")
      .replace(/\r/g, "")
      .replace(/\u2028|\u2029/g, "")
      .replace(/[ \t]+/g, " ")
      .trim());
  }).filter(function (line) {
    return line && !isNonHeroText(line);
  });
}

function cleanWireframeOutputLine(line) {
  return String(line || "")
    .replace(/(?:디자인\s*코멘트|코멘트|comment|수정\s*요청|변경\s*요청)\s*[:：]?.*$/i, "")
    .replace(/(?:메인\s*배경|타이머\s*배경|배경\s*컬러|컬러\s*반전|붉은\s*계열|긴급한\s*느낌|부탁드|부탁\s*드립니다).*$/i, "")
    .replace(/(?:디자인|레이아웃|색상|컬러|위치|크기|폰트|간격|정렬).*(?:해주세요|해\s*주세요).*$/i, "")
    .replace(/[-–—ㆍ·:：,\s]+$/g, "")
    .trim();
}

function truncatePageName(value) {
  return String(value || "기획안").replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim().slice(0, 42) || "기획안";
}

function normalizeWireframeElement(element, slideNumber, elementIndex) {
  if (!element) return null;
  var x = Number(element.x);
  var y = Number(element.y);
  var w = Number(element.w);
  var h = Number(element.h);
  if (!isFinite(x) || !isFinite(y) || !isFinite(w) || !isFinite(h) || w <= 0 || h <= 0) {
    logWireframeTextOmission("missing-bounds", element, slideNumber, elementIndex, element && element.text);
    return null;
  }

  if (element.type === "svg") {
    var svg = String(element.svg || "");
    if (svg.indexOf("<svg") < 0) return null;
    return {
      type: "svg",
      x: x,
      y: y,
      w: w,
      h: h,
      text: "",
      svg: svg,
      fill: "",
      stroke: "",
      fontSize: 1800,
      fontSizeUnit: "",
      zIndex: Number(element.zIndex) || 0,
      objectId: String(element.objectId || element.id || ""),
      groupId: element.groupId || "",
      transform: element.transform || null,
      radius: ""
    };
  }

  if (isDesignCommentText(element.text || "")) return null;
  var rawText = String(element.text || "");
  var lines = wireframeTextLines(rawText);
  if (!lines.length && element.type === "text") {
    logWireframeTextOmission("empty-after-normalize", element, slideNumber, elementIndex, rawText);
    return null;
  }
  var type = element.type || (lines.length ? "text" : "shape");
  if (!/^(text|image|shape|line|ellipse|table)$/i.test(String(type || ""))) return null;
  return {
    type: type,
    x: x,
    y: y,
    w: w,
    h: h,
    text: lines.join("\n"),
    fill: normalizeHexColor(element.fill),
    stroke: normalizeHexColor(element.stroke),
    imageDataUrl: isDataImageUrl(element.imageDataUrl || element.inlineImageDataUrl) ? String(element.imageDataUrl || element.inlineImageDataUrl) : "",
    imageUrl: String(element.imageUrl || element.contentUrl || ""),
    textColor: normalizeHexColor(element.textColor),
    fontSize: Number(element.fontSize) || 1800,
    fontSizeUnit: element.fontSizeUnit === "pt" ? "pt" : "",
    align: normalizeWireframeAlign(element.align),
    verticalAlign: normalizeWireframeVerticalAlign(element.verticalAlign),
    zIndex: Number(element.zIndex) || 0,
    objectId: String(element.objectId || element.id || ""),
    groupId: element.groupId || "",
    transform: element.transform || null,
    radius: element.radius || ""
  };
}

function logWireframeTextOmission(reason, element, slideNumber, elementIndex, rawText) {
  var text = String(rawText || "");
  if (!text.trim()) return;
  if (isDesignCommentText(text)) return;
  try {
    console.warn("[wireframe] text omitted", {
      reason: reason,
      slideNumber: slideNumber,
      elementIndex: elementIndex,
      objectId: element && (element.objectId || element.id || ""),
      groupId: element && element.groupId || "",
      text: text.slice(0, 160)
    });
  } catch (error) {
    // Console logging is best-effort in the Figma plugin runtime.
  }
}

function normalizeHexColor(value) {
  var match = String(value || "").match(/^#?[0-9A-Fa-f]{6}$/);
  return match ? ("#" + String(value).replace("#", "").toUpperCase()) : "";
}

function normalizeWireframeAlign(value) {
  var text = String(value || "").toUpperCase();
  return /^(LEFT|CENTER|RIGHT)$/.test(text) ? text : "";
}

function normalizeWireframeVerticalAlign(value) {
  var text = String(value || "").toUpperCase();
  return /^(TOP|CENTER|BOTTOM)$/.test(text) ? text : "";
}

function hasFullSlideSvgWireframeSlides(slides) {
  return safeArray(slides, []).some(function (slide) {
    return safeArray(slide && slide.elements, []).some(function (element) {
      return element && element.type === "svg" && String(element.svg || "").indexOf("<svg") >= 0;
    });
  });
}

function hasVisualWireframeSlides(slides) {
  return safeArray(slides, []).some(function (slide) {
    return safeArray(slide && slide.elements, []).some(function (element) {
      if (!element) return false;
      if (element.type === "svg" && String(element.svg || "").indexOf("<svg") >= 0) return true;
      if (element.type === "image" || element.imageDataUrl || element.imageUrl) return true;
      var hasBox = Number(element.w) > 0 && Number(element.h) > 0;
      var hasPosition = Number(element.x) !== 0 || Number(element.y) !== 0;
      return hasBox && (hasPosition || element.fill || element.stroke || element.text);
    });
  });
}

async function activatePage(page) {
  if (typeof figma.setCurrentPageAsync === "function") {
    await figma.setCurrentPageAsync(page);
    return;
  }

  try {
    figma.currentPage = page;
  } catch (error) {
    // Older plugin runtimes may not expose page switching. The page is still created.
  }
}

async function prepareFonts() {
  var availableFonts = await figma.listAvailableFontsAsync();
  var families = {};

  availableFonts.forEach(function (font) {
    var family = font.fontName.family;
    if (!families[family]) families[family] = [];
    families[family].push(font.fontName);
  });
  availableFontsByFamily = families;

  var preferredFamilies = [
    "Spoqa Han Sans Neo",
    "SUIT",
    "Wanted Sans",
    "Noto Sans KR",
    "Source Han Sans KR",
    "IBM Plex Sans KR",
    "Gmarket Sans",
    "Apple SD Gothic Neo",
    "Noto Sans CJK KR",
    "Inter",
    "Roboto"
  ];

  var selectedFamily = preferredFamilies.find(function (family) {
    return families[family] && families[family].some(function (fontName) {
      return !isTooHeavyFontStyle(fontName.style);
    });
  });

  if (!selectedFamily) {
    var readableFont = availableFonts.find(function (font) {
      return font && font.fontName && !isTooHeavyFontStyle(font.fontName.style);
    });
    var firstFont = (readableFont && readableFont.fontName) || (availableFonts[0] && availableFonts[0].fontName);
    if (firstFont) {
      fontRegular = firstFont;
      fontMedium = firstFont;
      fontBold = firstFont;
      await figma.loadFontAsync(fontRegular);
      return;
    }
  }

  var styles = families[selectedFamily] || [fontRegular];
  fontRegular = pickStyle(styles, ["Regular", "Book", "Normal", "Medium"]) || fontRegular;
  fontMedium = pickStyle(styles, ["Medium", "SemiBold", "Semibold", "Regular", "Book"]) || fontRegular;
  fontBold = pickStyle(styles, ["Bold", "SemiBold", "Semibold", "Medium"]) || fontMedium || fontRegular;

  await figma.loadFontAsync(fontRegular);
  if (fontMedium.family !== fontRegular.family || fontMedium.style !== fontRegular.style) {
    await figma.loadFontAsync(fontMedium);
  }
  if (fontBold.family !== fontRegular.family || fontBold.style !== fontRegular.style) {
    await figma.loadFontAsync(fontBold);
  }
}

function pickStyle(styles, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var style = styles.find(function (fontName) {
      return fontName.style === candidates[i] && !isTooHeavyFontStyle(fontName.style);
    });
    if (style) return style;
  }
  var readable = styles.find(function (fontName) {
    return !isTooHeavyFontStyle(fontName.style);
  });
  if (readable) return readable;
  return null;
}

function isTooHeavyFontStyle(style) {
  return /Black|Extra\s*Bold|Heavy/i.test(String(style || ""));
}

async function extractGoogleSlidesText(slidesInfo) {
  var errors = [];

  if (slidesInfo.proxyUrl) {
    try {
      var proxied = await fetchSlidesProxyText(slidesInfo);
      if (proxied.text) return proxied;
    } catch (error) {
      errors.push("Slides 프록시 요청 실패: " + error.message);
    }
  }

  if (slidesInfo.documentId) {
    var textExportUrls = [
      "https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/export/txt",
      "https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/export/txt?id=" + slidesInfo.documentId + "&exportFormat=txt",
      "https://docs.google.com/presentation/export/txt?id=" + slidesInfo.documentId + "&exportFormat=txt",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&exportFormat=txt",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&format=txt"
    ];

    for (var i = 0; i < textExportUrls.length; i++) {
      try {
        var text = await fetchText(textExportUrls[i]);
      if (text) return { text: text, sourceName: slidesInfo.sourceName, wireframeSlides: [], wireframeDataStatus: "text-only", viaProxy: false, method: "text-export" };
      } catch (error) {
        errors.push(error.message);
      }
    }

    var exportUrls = pptxExportUrlsForSlides(slidesInfo);

    for (var k = 0; k < exportUrls.length; k++) {
      try {
        var buffer = await fetchBinary(exportUrls[k]);
        var pptDeck = await extractDeckFromExportBuffer(buffer);
        if (pptDeck.text) {
          return {
            text: pptDeck.text,
            referenceSlideImages: pptDeck.referenceSlideImages,
            referenceSlideHeroText: pptDeck.referenceSlideHeroText || "",
            wireframeSlides: pptDeck.wireframeSlides || [],
            wireframeDataStatus: hasVisualWireframeSlides(pptDeck.wireframeSlides) ? "coordinate" : "text-only",
            sourceName: slidesInfo.sourceName,
            viaProxy: false,
            method: "pptx-export"
          };
        }
      } catch (error) {
        errors.push(error.message);
      }
    }

    var htmlExportUrls = [
      "https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/export/html",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&exportFormat=html",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&format=html"
    ];

    for (var h = 0; h < htmlExportUrls.length; h++) {
      try {
        var htmlText = await fetchPublishedText(htmlExportUrls[h]);
        if (htmlText) return { text: htmlText, sourceName: slidesInfo.sourceName, wireframeSlides: [], wireframeDataStatus: "text-only", viaProxy: false, method: "html-export" };
      } catch (error) {
        errors.push(error.message);
      }
    }
  }

  if (slidesInfo.publicUrl) {
    try {
      var htmlText = await fetchPublishedText(slidesInfo.publicUrl);
      if (htmlText) return { text: htmlText, sourceName: slidesInfo.sourceName, wireframeSlides: [], wireframeDataStatus: "text-only", viaProxy: false, method: "published-html" };
    } catch (error) {
      errors.push(error.message);
    }
  }

  throw new Error(summarizeNetworkErrors(errors));
}

async function fetchSlidesProxyText(slidesInfo) {
  var proxyUrl = String(slidesInfo.proxyUrl || "").replace(/\/+$/, "");
  if (!proxyUrl) throw new Error("프록시 URL이 비어 있습니다.");

  var sourceUrl = slidesInfo.originalUrl || slidesInfo.publicUrl || "";
  if (!sourceUrl) throw new Error("프록시로 전달할 Google Slides 링크가 없습니다.");

  var requestUrl = proxyUrl + "?url=" + encodeURIComponent(sourceUrl);
  var response = await fetch(requestUrl, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error("Slides 프록시 요청이 실패했습니다. (" + response.status + ")");
  }

  var contentType = "";
  if (response.headers && typeof response.headers.get === "function") {
    contentType = response.headers.get("content-type") || "";
  }
  var text = await response.text();
  text = await followAppsScriptRedirectText(text);
  var payload = null;
  if (contentType.indexOf("application/json") >= 0 || /^\s*[{[]/.test(text)) {
    try {
      payload = JSON.parse(text);
    } catch (error) {
      payload = null;
    }
  }

  if (payload && payload.error) {
    throw new Error(payload.error);
  }

  var extractedText = payload && payload.text ? payload.text : text;
  extractedText = normalizeFetchedText(extractedText);
  if (!extractedText) {
    throw new Error("Slides 프록시가 분석 가능한 텍스트를 반환하지 않았습니다.");
  }

  if (!payload && looksLikeAccessPage(extractedText)) {
    if (/<html|accounts\.google\.com|Sign in|로그인/i.test(text)) {
      throw new Error("Apps Script가 로그인 페이지를 반환했습니다. proxy=" + proxyUrl + " response=" + summarizeResponseText(text));
    }
    throw new Error("Slides 프록시가 분석 가능한 텍스트를 반환하지 않았습니다.");
  }

  var proxyWireframeSlides = payload && payload.wireframeSlides ? payload.wireframeSlides : [];
  var proxyWireframeDataStatus = payload && payload.wireframeDataStatus
    ? payload.wireframeDataStatus
    : (hasVisualWireframeSlides(proxyWireframeSlides) ? "coordinate" : "text-only");
  if (!hasVisualWireframeSlides(proxyWireframeSlides) && payload && payload.pptxBase64) {
    try {
      var proxiedPptxDeck = await extractDeckFromProxyPptxBase64(payload.pptxBase64);
      if (hasVisualWireframeSlides(proxiedPptxDeck.wireframeSlides)) {
        return {
          text: extractedText || proxiedPptxDeck.text,
          sourceName: slidesInfo.sourceName,
          referenceSlideImages: proxiedPptxDeck.referenceSlideImages || [],
          referenceSlideHeroText: proxiedPptxDeck.referenceSlideHeroText || "",
          wireframeSlides: proxiedPptxDeck.wireframeSlides,
          wireframeDataStatus: "coordinate",
          proxyVersion: payload.proxyVersion || "",
          viaProxy: true,
          method: (payload.method || "slides-proxy") + "+parsed-pptx"
        };
      }
    } catch (error) {
      proxyWireframeDataStatus = "pptx-base64-unreadable";
    }
  }
  if (!hasVisualWireframeSlides(proxyWireframeSlides)) {
    try {
      var visualDeck = await fetchPublishedSvgDeck(slidesInfo);
      if (visualDeck.wireframeSlides && visualDeck.wireframeSlides.length) {
        return {
          text: extractedText || visualDeck.text,
          sourceName: slidesInfo.sourceName,
          referenceSlideHeroText: payload && payload.referenceSlideHeroText ? payload.referenceSlideHeroText : "",
          wireframeSlides: visualDeck.wireframeSlides,
          wireframeDataStatus: "coordinate",
          proxyVersion: payload && payload.proxyVersion ? payload.proxyVersion : "",
          viaProxy: true,
          method: (payload && payload.method ? payload.method : "slides-proxy") + "+published-svg"
        };
      }
    } catch (error) {
      // Keep the stable proxy text path when the public visual layer is unavailable.
    }
  }

  return {
    text: extractedText,
    sourceName: slidesInfo.sourceName,
    referenceSlideHeroText: payload && payload.referenceSlideHeroText ? payload.referenceSlideHeroText : "",
    wireframeSlides: proxyWireframeSlides,
    wireframeDataStatus: proxyWireframeDataStatus,
    proxyVersion: payload && payload.proxyVersion ? payload.proxyVersion : "",
    viaProxy: true,
    method: payload && payload.method ? payload.method : "slides-proxy"
  };
}

async function extractDeckFromProxyPptxBase64(base64) {
  var buffer = base64ToArrayBuffer(base64);
  assertPptxBuffer(buffer, "Apps Script PPTX proxy");
  return await extractDeckFromExportBuffer(buffer);
}

function base64ToArrayBuffer(base64) {
  var clean = String(base64 || "").replace(/^data:[^,]+,/, "").replace(/\s+/g, "");
  if (!clean) throw new Error("프록시 PPTX 데이터가 비어 있습니다.");
  var binary = "";
  if (typeof atob === "function") {
    binary = atob(clean);
  } else {
    binary = decodeBase64Binary(clean);
  }
  var bytes = new Uint8Array(binary.length);
  for (var i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i) & 0xff;
  }
  return bytes.buffer;
}

function decodeBase64Binary(base64) {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var lookup = {};
  for (var i = 0; i < chars.length; i++) lookup[chars[i]] = i;
  var output = "";
  var buffer = 0;
  var bits = 0;
  for (var index = 0; index < base64.length; index++) {
    var char = base64[index];
    if (char === "=") break;
    if (!Object.prototype.hasOwnProperty.call(lookup, char)) continue;
    buffer = (buffer << 6) | lookup[char];
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      output += String.fromCharCode((buffer >> bits) & 0xff);
    }
  }
  return output;
}

function summarizeResponseText(text) {
  return normalizeFetchedText(stripHtml(String(text || ""))).slice(0, 220);
}

async function followAppsScriptRedirectText(text) {
  var redirectUrl = extractAppsScriptRedirectUrl(text);
  if (!redirectUrl) return text;

  var redirectResponse = await fetch(redirectUrl, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });

  if (!redirectResponse.ok) {
    throw new Error("Apps Script 리다이렉트 요청이 실패했습니다. (" + redirectResponse.status + ")");
  }

  var redirectedText = await redirectResponse.text();
  var nextRedirectUrl = extractAppsScriptRedirectUrl(redirectedText);
  if (nextRedirectUrl && nextRedirectUrl !== redirectUrl) {
    var nextResponse = await fetch(nextRedirectUrl, {
      credentials: "omit",
      redirect: "follow",
      cache: "no-store"
    });
    if (!nextResponse.ok) {
      throw new Error("Apps Script 리다이렉트 요청이 실패했습니다. (" + nextResponse.status + ")");
    }
    return await nextResponse.text();
  }

  return redirectedText;
}

function extractAppsScriptRedirectUrl(text) {
  var match = String(text || "").match(/https:\/\/script\.googleusercontent\.com\/macros\/echo\?[^"'<>\s]+/i);
  if (!match) return "";
  return decodeHtmlAttribute(match[0]);
}

function decodeHtmlAttribute(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

async function convertGoogleSlidesToPptxText(slidesInfo) {
  var errors = [];
  var sawReadablePptx = false;
  var sawCoordinateFailure = false;

  var exportUrls = pptxExportUrlsForSlides(slidesInfo);
  if (!exportUrls.length) {
    throw createPluginError(
      WIRE_ERROR.GOOGLE_SLIDES_EXPORT_FAILED,
      "와이어프레임으로 읽을 Google Slides 링크 형식을 찾지 못했습니다.",
      summarizeNetworkErrors(errors)
    );
  }

  for (var i = 0; i < exportUrls.length; i++) {
    try {
      var buffer = await fetchBinary(exportUrls[i]);
      var pptDeck = await extractDeckFromExportBuffer(buffer);
      sawReadablePptx = true;
      if (pptDeck.text) {
        var convertedDeck = {
          text: pptDeck.text,
          referenceSlideImages: pptDeck.referenceSlideImages,
          referenceSlideHeroText: pptDeck.referenceSlideHeroText || "",
          wireframeSlides: pptDeck.wireframeSlides || [],
          wireframeDataStatus: hasVisualWireframeSlides(pptDeck.wireframeSlides) ? "coordinate" : "text-only",
          sourceName: slidesInfo.sourceName,
          viaProxy: false,
          method: "pptx-export"
        };
        if (hasFullSlideSvgWireframeSlides(convertedDeck.wireframeSlides) || hasVisualWireframeSlides(convertedDeck.wireframeSlides)) {
          return convertedDeck;
        }
        sawCoordinateFailure = true;
        errors.push(createPluginError(
          WIRE_ERROR.PPTX_COORDINATE_EXTRACTION_FAILED,
          "PPTX export에서 텍스트는 읽었지만 와이어프레임 좌표 요소를 찾지 못했습니다."
        ));
      }
    } catch (error) {
      errors.push(error);
    }
  }

  try {
    var publishedDeck = await fetchPublishedSvgDeck(slidesInfo);
    if (publishedDeck.wireframeSlides && publishedDeck.wireframeSlides.length) {
      return publishedDeck;
    }
  } catch (error) {
    errors.push(error);
  }

  if (!sawReadablePptx && errors.some(isGoogleSlidesAccessBlockedError)) {
    throw createPluginError(
      WIRE_ERROR.GOOGLE_SLIDES_ACCESS_BLOCKED,
      "Google Slides 요청이 보안 정책 또는 권한 문제로 차단되었습니다.",
      firstErrorMessage(errors)
    );
  }

  if (sawReadablePptx || sawCoordinateFailure) {
    throw createPluginError(
      WIRE_ERROR.PPTX_COORDINATE_EXTRACTION_FAILED,
      "PPTX는 읽었지만 원본 슬라이드의 좌표 요소를 찾지 못했습니다.",
      firstErrorMessage(errors)
    );
  }

  throw createPluginError(
    WIRE_ERROR.GOOGLE_SLIDES_EXPORT_FAILED,
    "Google Slides를 PPTX로 변환하지 못했습니다.",
    firstErrorMessage(errors) || summarizeNetworkErrors(errors)
  );
}

function pptxExportUrlsForSlides(slidesInfo) {
  if (slidesInfo.documentId) {
    return [
      "https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/export/pptx",
      "https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/export/pptx?id=" + slidesInfo.documentId + "&exportFormat=pptx",
      "https://docs.google.com/presentation/export/pptx?id=" + slidesInfo.documentId + "&exportFormat=pptx",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&exportFormat=pptx",
      "https://docs.google.com/feeds/download/presentations/Export?id=" + slidesInfo.documentId + "&format=pptx"
    ];
  }

  if (slidesInfo.publishedId) {
    return [
      "https://docs.google.com/presentation/d/e/" + slidesInfo.publishedId + "/pub?output=pptx",
      "https://docs.google.com/presentation/d/e/" + slidesInfo.publishedId + "/export/pptx",
      "https://docs.google.com/presentation/export/pptx?id=" + slidesInfo.publishedId + "&exportFormat=pptx"
    ];
  }

  return [];
}

async function fetchPublishedSvgDeck(slidesInfo) {
  var errors = [];
  var urls = publishedHtmlUrlsForSlides(slidesInfo);
  for (var i = 0; i < urls.length; i++) {
    try {
      var html = await fetchRawHtml(urls[i]);
      var deck = extractPublishedSvgDeck(html, slidesInfo);
      if (deck.wireframeSlides && deck.wireframeSlides.length) {
        return deck;
      }
    } catch (error) {
      errors.push(error.message);
    }
  }

  throw new Error("공개 Google Slides에서 실제 슬라이드 SVG를 찾지 못했습니다. " + summarizeNetworkErrors(errors));
}

function publishedHtmlUrlsForSlides(slidesInfo) {
  var urls = [];
  if (slidesInfo.publicUrl) urls.push(slidesInfo.publicUrl);
  if (slidesInfo.publishedId) {
    urls.push("https://docs.google.com/presentation/d/e/" + slidesInfo.publishedId + "/pub?start=false&loop=false&delayms=3000");
    urls.push("https://docs.google.com/presentation/d/e/" + slidesInfo.publishedId + "/embed?start=false&loop=false&delayms=3000");
  }
  if (slidesInfo.documentId) {
    urls.push("https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/preview");
    urls.push("https://docs.google.com/presentation/d/" + slidesInfo.documentId + "/present");
  }
  return uniqueStrings(urls);
}

async function fetchRawHtml(url) {
  var response = await fetch(url, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });
  if (!response.ok) {
    throw new Error("공개 슬라이드 페이지 요청이 실패했습니다. (" + response.status + ")");
  }
  var html = await response.text();
  if (!html || (looksLikeAccessPage(html) && html.indexOf("SK_svgData") < 0 && html.indexOf("docData:") < 0)) {
    throw new Error("공개 슬라이드 페이지에 접근하지 못했습니다.");
  }
  return html;
}

function extractPublishedSvgDeck(html, slidesInfo) {
  var svgSlides = extractPublishedSvgSlides(html);
  if (!svgSlides.length) {
    throw new Error("공개 슬라이드 HTML 안에서 렌더링 SVG를 찾지 못했습니다.");
  }

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

  var slides = svgSlides.map(function (slide, index) {
    var text = pageTextById[slide.pageId] || extractTextFromSvg(slide.svg);
    var lines = wireframeTextLines(text);
    var width = slide.width || 960;
    var height = slide.height || 540;
    return {
      number: index + 1,
      title: lines[0] || "Slide " + (index + 1),
      text: lines.join("\n"),
      lines: lines,
      width: width,
      height: height,
      elements: [{
        type: "svg",
        x: 0,
        y: 0,
        w: width,
        h: height,
        svg: slide.svg
      }]
    };
  }).filter(function (slide) {
    return slide.elements.length;
  });

  var textParts = slides.map(function (slide) {
    return "Slide " + slide.number + "\n" + slide.text;
  });
  var extractedText = normalizeFetchedText(textParts.join("\n\n")) || extractPublishedSlidesText(html);

  return {
    text: extractedText,
    referenceSlideImages: [],
    referenceSlideHeroText: "",
    wireframeSlides: slides,
    wireframeDataStatus: "coordinate",
    sourceName: slidesInfo.sourceName || "Google Slides published link",
    viaProxy: false,
    method: "published-svg"
  };
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

function extractPublishedSvgSlides(html) {
  var source = String(html || "");
  var slides = [];
  var seen = {};
  var pattern = /SK_svgData\s*=\s*'([\s\S]*?)';[\s\S]*?SK_viewerApp\.setPageData\('([^']+)'/g;
  var match;
  while ((match = pattern.exec(source))) {
    addPublishedSvgSlide(slides, seen, match[1], match[2]);
  }

  if (!slides.length) {
    var loosePattern = /SK_svgData\s*=\s*(["'])([\s\S]*?)\1/g;
    while ((match = loosePattern.exec(source))) {
      var nearby = source.slice(match.index, match.index + 8000);
      var pageMatch = nearby.match(/SK_viewerApp\.setPageData\(\s*(["'])([^"']+)\1/);
      addPublishedSvgSlide(slides, seen, match[2], pageMatch ? pageMatch[2] : "page-" + (slides.length + 1));
    }
  }
  return slides;
}

function addPublishedSvgSlide(slides, seen, encodedSvg, pageId) {
  var svg = decodePublishedJsString(encodedSvg);
  if (svg.indexOf("<svg") < 0) return;
  var cleaned = sanitizePublishedSvg(svg);
  var key = String(pageId || "") + ":" + cleaned.slice(0, 120);
  if (seen[key]) return;
  seen[key] = true;
  var size = svgViewBoxSize(cleaned);
  slides.push({
    pageId: pageId || "page-" + (slides.length + 1),
    svg: cleaned,
    width: size.width,
    height: size.height
  });
}

function sanitizePublishedSvg(svg) {
  return String(svg || "")
    .replace(/<script\b[\s\S]*?<\/script>/gi, "")
    .replace(/\s+on[a-z]+\s*=\s*(["'])[\s\S]*?\1/gi, "");
}

function svgViewBoxSize(svg) {
  var match = String(svg || "").match(/viewBox=(["'])([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\1/i);
  if (match) {
    return {
      width: Number(match[4]) || 960,
      height: Number(match[5]) || 540
    };
  }
  var width = String(svg || "").match(/\bwidth=(["'])([-\d.]+)\1/i);
  var height = String(svg || "").match(/\bheight=(["'])([-\d.]+)\1/i);
  return {
    width: width ? Number(width[2]) || 960 : 960,
    height: height ? Number(height[2]) || 540 : 540
  };
}

function extractTextFromSvg(svg) {
  var source = String(svg || "");
  var parts = [];
  var seen = {};
  var labelPattern = /\saria-label=(["'])([\s\S]*?)\1/gi;
  var match;
  while ((match = labelPattern.exec(source))) {
    addUniquePart(parts, seen, decodeXml(match[2]).replace(/\u00a0/g, " "));
  }

  var titlePattern = /<title\b[^>]*>([\s\S]*?)<\/title>/gi;
  while ((match = titlePattern.exec(source))) {
    addUniquePart(parts, seen, decodeXml(match[1]).replace(/\u00a0/g, " "));
  }

  return normalizeFetchedText(parts.join("\n"));
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

function uniqueStrings(values) {
  var seen = {};
  var result = [];
  values.forEach(function (value) {
    var text = String(value || "").trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    result.push(text);
  });
  return result;
}

async function fetchText(url) {
  var response = await fetch(url, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });
  if (!response.ok) {
    throw new Error("Google Slides 텍스트 요청이 실패했습니다. (" + response.status + ")");
  }

  var text = normalizeFetchedText(await response.text());
  if (!text || looksLikeAccessPage(text)) {
    throw new Error("Google Slides 텍스트에 접근하지 못했습니다.");
  }
  return text;
}

async function fetchBinary(url) {
  var response = await fetch(url, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });
  if (!response.ok) {
    var code = response.status === 401 || response.status === 403
      ? WIRE_ERROR.GOOGLE_SLIDES_ACCESS_BLOCKED
      : WIRE_ERROR.GOOGLE_SLIDES_EXPORT_FAILED;
    throw createPluginError(code, "Google Slides export 요청이 실패했습니다. (" + response.status + ")");
  }

  var buffer = await response.arrayBuffer();
  assertPptxBuffer(buffer, "Google Slides export");
  return buffer;
}

function assertPptxBuffer(buffer, sourceLabel) {
  var bytes = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) {
    var label = String(sourceLabel || "");
    var code = /Google Slides|export/i.test(label)
      ? WIRE_ERROR.GOOGLE_SLIDES_ACCESS_BLOCKED
      : WIRE_ERROR.PPTX_UPLOAD_INVALID;
    throw createPluginError(code, label + "이 PPTX 파일이 아니거나 접근 안내 페이지로 내려왔습니다.");
  }
}

function messageBytesToArrayBuffer(value) {
  if (value instanceof ArrayBuffer) {
    return value;
  }

  var bytes;
  if (value instanceof Uint8Array) {
    bytes = value;
  } else if (ArrayBuffer.isView && ArrayBuffer.isView(value)) {
    bytes = new Uint8Array(value.buffer, value.byteOffset || 0, value.byteLength || value.buffer.byteLength);
  } else if (Array.isArray(value)) {
    bytes = new Uint8Array(value);
  } else if (value && value.buffer instanceof ArrayBuffer) {
    bytes = new Uint8Array(value.buffer, value.byteOffset || 0, value.byteLength || value.buffer.byteLength);
  } else if (value && typeof value === "object") {
    var keys = Object.keys(value).sort(function (a, b) {
      return Number(a) - Number(b);
    });
    bytes = new Uint8Array(keys.map(function (key) {
      return value[key];
    }));
  } else {
    throw new Error("PPTX 파일 데이터를 읽지 못했습니다.");
  }

  var copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

async function fetchPublishedText(url) {
  var response = await fetch(url, {
    credentials: "omit",
    redirect: "follow",
    cache: "no-store"
  });
  if (!response.ok) {
    throw new Error("공개 슬라이드 페이지 요청이 실패했습니다. (" + response.status + ")");
  }

  var html = await response.text();
  var text = extractPublishedSlidesText(html) || normalizeFetchedText(stripHtml(html));
  if (!text || looksLikeAccessPage(text)) {
    throw new Error("공개 슬라이드 페이지에서 분석 가능한 텍스트를 찾지 못했습니다.");
  }
  return text;
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

function extractPublishedMetaText(html) {
  var fields = [];
  var metaPattern = /<meta[^>]+(?:property|name)=(["'])(og:title|og:description|description)\1[^>]+content=(["'])([\s\S]*?)\3[^>]*>/gi;
  var match;
  while ((match = metaPattern.exec(html))) {
    fields.push(decodePublishedAttribute(match[4]));
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

function stripHtml(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, "\n")
    .replace(/<style[\s\S]*?<\/style>/gi, "\n")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, "\n")
    .replace(/<[^>]+>/g, "\n")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&amp;/g, "&");
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

function summarizeNetworkErrors(errors) {
  var filtered = [];
  errors.forEach(function (error) {
    var message = error && error.message ? error.message : String(error || "");
    if (message && filtered.indexOf(message) < 0) filtered.push(message);
  });
  if (!filtered.length) return "Google Slides 링크에서 텍스트를 읽지 못했습니다.";
  var blocked = filtered.find(function (message) {
    return /failed to fetch|load failed|networkerror|cors|보안 정책/i.test(message);
  });
  if (blocked) {
    return "Google Slides 요청이 보안 정책으로 차단되었습니다. 플러그인을 다시 import한 뒤 시도하거나, 조직 보안 정책에서 외부 export 접근이 막혀 있는지 확인해 주세요.";
  }
  return filtered[0];
}

async function extractTextFromExportBuffer(buffer) {
  var deck = await extractDeckFromExportBuffer(buffer);
  return deck.text;
}

async function extractDeckFromExportBuffer(buffer) {
  var entries = parseZipEntries(buffer);
  var slideSize = await extractPresentationSlideSize(entries, buffer);
  var slideEntries = entries
    .filter(function (entry) { return /^ppt\/slides\/slide\d+\.xml$/i.test(entry.name); })
    .sort(function (a, b) { return slideNumber(a.name) - slideNumber(b.name); });

  if (!slideEntries.length) {
    throw new Error("Google Slides export 안에서 슬라이드 데이터를 찾지 못했습니다.");
  }

  var parts = [];
  var slideData = [];
  for (var i = 0; i < slideEntries.length; i++) {
    var bytes = await readZipEntry(slideEntries[i], buffer);
    var xml = new TextDecoder("utf-8").decode(bytes);
    var slideText = extractTextFromSlideXml(xml);
    var number = slideNumber(slideEntries[i].name);
    slideData.push({ number: number, entry: slideEntries[i], xml: xml, text: slideText });
    if (slideText) parts.push("Slide " + number + "\n" + slideText);
  }
  return {
    text: normalizeFetchedText(parts.join("\n\n")),
    referenceSlideImages: await extractReferenceSlideImages(slideData, entries, buffer),
    referenceSlideHeroText: extractReferenceSlideHeroText(slideData),
    wireframeSlides: await PptxCoordinateExtractionModule.extractSlides(slideData, slideSize, entries, buffer)
  };
}

async function extractPresentationSlideSize(entries, buffer) {
  var presentationEntry = entries.find(function (entry) {
    return entry.name === "ppt/presentation.xml";
  });
  if (!presentationEntry) {
    return { width: 12192000, height: 6858000 };
  }

  try {
    var xml = new TextDecoder("utf-8").decode(await readZipEntry(presentationEntry, buffer));
    var match = xml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    if (match) {
      return {
        width: Number(match[1]) || 12192000,
        height: Number(match[2]) || 6858000
      };
    }
  } catch (error) {
    // Use the common 16:9 PowerPoint size when presentation metadata is unavailable.
  }

  return { width: 12192000, height: 6858000 };
}

async function extractReferenceSlideImages(slides, entries, buffer) {
  var targetSlide = referenceSlideAfterMainMarker(slides);
  if (!targetSlide) return [];
  var relEntryName = "ppt/slides/_rels/slide" + targetSlide.number + ".xml.rels";
  var relEntry = entries.find(function (entry) {
    return entry.name === relEntryName;
  });
  if (!relEntry) return [];

  var relXml = new TextDecoder("utf-8").decode(await readZipEntry(relEntry, buffer));
  var rels = parseSlideRelationships(relXml);
  var imageIds = embeddedImageIds(targetSlide.xml);
  var images = [];

  for (var i = 0; i < imageIds.length && images.length < 4; i++) {
    var rel = rels[imageIds[i]];
    if (!rel || !/image/i.test(rel.type || rel.target || "")) continue;
    var mediaPath = normalizeZipPath("ppt/slides/" + rel.target);
    var mediaEntry = entries.find(function (entry) {
      return entry.name === mediaPath;
    });
    if (!mediaEntry) continue;
    var bytes = await readZipEntry(mediaEntry, buffer);
    var mime = mimeForImagePath(mediaPath);
    if (!mime) continue;
    images.push({
      title: "메인페이지 다음 슬라이드 레퍼런스",
      source: "PPTX reference slide",
      imageUrl: "pptx://" + mediaPath,
      inlineImageDataUrl: "data:" + mime + ";base64," + bytesToBase64(bytes),
      pageUrl: "",
      keyword: "기획안 레퍼런스",
      whyMatched: "기획안 메인페이지 다음 슬라이드에 첨부된 원본 레퍼런스"
    });
  }
  return images;
}

function referenceSlideAfterMainMarker(slides) {
  var markerIndex = slides.findIndex(isExactMainPageMarkerSlide);
  if (markerIndex < 0 || !slides[markerIndex + 1]) return null;
  return slides[markerIndex + 1];
}

async function extractWireframeSlides(slides, slideSize, entries, buffer) {
  var markerIndex = slides.findIndex(isExactMainPageMarkerSlide);
  var markerSlides = markerIndex >= 0 ? slides.slice(markerIndex + 1) : slides;
  var targetSlides = markerSlides.length ? markerSlides : slides;
  var result = await extractWireframeSlidesFromTargets(targetSlides, slideSize, entries, buffer);
  if (!result.length && markerIndex >= 0) {
    result = await extractWireframeSlidesFromTargets(slides, slideSize, entries, buffer);
  }
  return result;
}

async function extractWireframeSlidesFromTargets(targetSlides, slideSize, entries, buffer) {
  var result = [];
  for (var index = 0; index < targetSlides.length; index++) {
    var slide = targetSlides[index];
    var lines = wireframeTextLines(slide.text);
    var imageDataByRelId = await slideImageDataByRelId(slide.number, entries, buffer);
    var elements = extractWireframeElementsFromSlideXml(slide.xml, imageDataByRelId);
    if (!lines.length && !elements.length) continue;
    result.push({
      number: slide.number || index + 1,
      title: lines[0] || "Slide " + (slide.number || index + 1),
      text: lines.join("\n"),
      lines: lines,
      width: slideSize.width,
      height: slideSize.height,
      elements: elements
    });
  }
  return result;
}

async function slideImageDataByRelId(slideNumberValue, entries, buffer) {
  var relEntryName = "ppt/slides/_rels/slide" + slideNumberValue + ".xml.rels";
  var relEntry = safeArray(entries, []).find(function (entry) {
    return entry.name === relEntryName;
  });
  if (!relEntry) return {};

  var relXml = new TextDecoder("utf-8").decode(await readZipEntry(relEntry, buffer));
  var rels = parseSlideRelationships(relXml);
  var imageData = {};
  var relIds = Object.keys(rels);
  for (var i = 0; i < relIds.length; i++) {
    var relId = relIds[i];
    var rel = rels[relId];
    if (!rel || !/image/i.test(rel.type || rel.target || "")) continue;
    var mediaPath = normalizeZipPath("ppt/slides/" + rel.target);
    var mediaEntry = safeArray(entries, []).find(function (entry) {
      return entry.name === mediaPath;
    });
    if (!mediaEntry) continue;
    var mime = mimeForImagePath(mediaPath);
    if (!mime) continue;
    var bytes = await readZipEntry(mediaEntry, buffer);
    imageData[relId] = {
      mediaPath: mediaPath,
      inlineImageDataUrl: "data:" + mime + ";base64," + bytesToBase64(bytes)
    };
  }
  return imageData;
}

function extractWireframeElementsFromSlideXml(xml, imageDataByRelId) {
  var elements = [];
  var pattern = /<p:(sp|pic|graphicFrame|cxnSp)\b[\s\S]*?<\/p:\1>/g;
  var match;
  var order = 0;
  while ((match = pattern.exec(xml))) {
    var tagName = match[1];
    var elementXml = match[0];
    var bounds = xfrmBounds(elementXml);
    if (!bounds) continue;
    var objectId = pptxElementObjectId(elementXml, order);
    var text = normalizeFetchedText(textParagraphsInXml(elementXml).join("\n"));
    if (elementXml.indexOf("<p:txBody") >= 0 && !text) {
      logPptxWireframeTextOmission("pptx-text-body-empty", objectId, elementXml);
    }
    if (isDesignCommentText(text)) continue;
    var lines = wireframeTextLines(text);

    if (tagName === "pic") {
      var imageRelId = embeddedImageIds(elementXml)[0] || "";
      var imageData = imageRelId ? imageDataByRelId[imageRelId] : null;
      elements.push({
        type: "image",
        x: bounds.x,
        y: bounds.y,
        w: bounds.w,
        h: bounds.h,
        text: "",
        fill: "#E6E6E6",
        stroke: "",
        imageDataUrl: imageData ? imageData.inlineImageDataUrl : "",
        imageUrl: imageData ? "pptx://" + imageData.mediaPath : "",
        objectId: objectId,
        groupId: "",
        shapeType: "PICTURE",
        zIndex: order++
      });
      continue;
    }

    if (tagName === "cxnSp") {
      elements.push({
        type: "line",
        x: bounds.x,
        y: bounds.y,
        w: bounds.w,
        h: bounds.h,
        text: "",
        fill: "",
        stroke: shapeLineFill(elementXml) || "#111111",
        objectId: objectId,
        groupId: "",
        shapeType: "LINE",
        zIndex: order++
      });
      continue;
    }

    var fill = tagName === "sp" ? shapeSolidFill(elementXml) : "#FFFFFF";
    var stroke = tagName === "sp" ? shapeLineFill(elementXml) : "#D9D9D9";
    if (!lines.length && text) logPptxWireframeTextOmission("pptx-text-empty-after-normalize", objectId, text);
    if (!lines.length && !fill && !stroke) {
      order++;
      continue;
    }
    elements.push({
      type: shapeElementTypeFromPptx(elementXml, lines),
      x: bounds.x,
      y: bounds.y,
      w: bounds.w,
      h: bounds.h,
      text: lines.join("\n"),
      fill: fill,
      stroke: stroke,
      textColor: shapeTextColorFromPptx(elementXml),
      align: shapeTextAlignFromPptx(elementXml),
      verticalAlign: shapeTextVerticalAlignFromPptx(elementXml),
      fontSize: largestFontSizeInXml(elementXml) || 1800,
      radius: /prst="roundRect"/i.test(elementXml) ? "round" : "",
      objectId: objectId,
      groupId: "",
      shapeType: pptxShapeType(elementXml),
      textRuns: pptxTextRuns(elementXml),
      zIndex: order++
    });
  }
  return elements;
}

function pptxElementObjectId(xml, order) {
  var match = String(xml || "").match(/<p:cNvPr\b[^>]*\bid="([^"]+)"[^>]*\bname="([^"]*)"/);
  if (match) return "pptx-" + match[1] + "-" + decodeXml(match[2] || "");
  var idOnly = String(xml || "").match(/<p:cNvPr\b[^>]*\bid="([^"]+)"/);
  if (idOnly) return "pptx-" + idOnly[1];
  return "pptx-element-" + (order + 1);
}

function pptxShapeType(xml) {
  var match = String(xml || "").match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/);
  if (match) return match[1].toUpperCase();
  if (String(xml || "").indexOf("<p:pic") >= 0) return "PICTURE";
  if (String(xml || "").indexOf("<p:graphicFrame") >= 0) return "GRAPHIC_FRAME";
  if (String(xml || "").indexOf("<p:cxnSp") >= 0) return "LINE";
  return "";
}

function pptxTextRuns(xml) {
  var runs = [];
  var runPattern = /<a:r\b[\s\S]*?<\/a:r>/g;
  var match;
  while ((match = runPattern.exec(String(xml || "")))) {
    var runXml = match[0];
    var text = normalizeFetchedText(textValuesInXml(runXml).join(""));
    if (!text) continue;
    runs.push({
      text: text,
      fontSize: largestFontSizeInXml(runXml) || largestFontSizeInXml(xml) || 1800,
      textColor: firstSrgbColor(runXml),
      bold: /\sb="1"/.test(runXml)
    });
  }
  return runs;
}

function logPptxWireframeTextOmission(reason, objectId, detail) {
  try {
    console.warn("[wireframe] text omitted", {
      reason: reason,
      objectId: objectId || "",
      detail: normalizeFetchedText(String(detail || "")).slice(0, 160)
    });
  } catch (error) {
  }
}

function shapeElementTypeFromPptx(xml, lines) {
  if (/prst="ellipse"|prst="arc"|prst="pie"|prst="donut"/i.test(String(xml || ""))) return "ellipse";
  return lines.length ? "text" : "shape";
}

function shapeTextColorFromPptx(xml) {
  var txBody = String(xml || "").match(/<p:txBody\b[\s\S]*?<\/p:txBody>/);
  return txBody ? firstSrgbColor(txBody[0]) : "";
}

function shapeTextAlignFromPptx(xml) {
  var match = String(xml || "").match(/<a:pPr\b[^>]*\balgn="([^"]+)"/);
  if (!match) return "";
  var value = match[1].toLowerCase();
  if (value === "ctr" || value === "center") return "CENTER";
  if (value === "r" || value === "right") return "RIGHT";
  return "LEFT";
}

function shapeTextVerticalAlignFromPptx(xml) {
  var match = String(xml || "").match(/<a:bodyPr\b[^>]*\banchor="([^"]+)"/);
  if (!match) return "";
  var value = match[1].toLowerCase();
  if (value === "mid" || value === "ctr" || value === "center") return "CENTER";
  if (value === "b" || value === "bottom") return "BOTTOM";
  return "TOP";
}

function xfrmBounds(xml) {
  var xfrm = String(xml || "").match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/);
  if (!xfrm) return null;
  var off = xfrm[0].match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
  var ext = xfrm[0].match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!off || !ext) return null;
  return {
    x: Number(off[1]) || 0,
    y: Number(off[2]) || 0,
    w: Number(ext[1]) || 1,
    h: Number(ext[2]) || 1
  };
}

function shapeSolidFill(xml) {
  var spPrMatch = String(xml || "").match(/<p:spPr\b[\s\S]*?<\/p:spPr>/);
  var spPr = spPrMatch ? spPrMatch[0] : "";
  if (!spPr || /<a:noFill\s*\/?>/i.test(spPr)) return "";
  return firstSrgbColor(spPr);
}

function shapeLineFill(xml) {
  var lnMatch = String(xml || "").match(/<a:ln\b[\s\S]*?<\/a:ln>/);
  return lnMatch ? firstSrgbColor(lnMatch[0]) : "";
}

function firstSrgbColor(xml) {
  var match = String(xml || "").match(/<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/);
  if (match) return "#" + match[1].toUpperCase();
  var scheme = String(xml || "").match(/<a:schemeClr\b[^>]*\bval="([^"]+)"/);
  if (scheme) return schemeColorToWireframeHex(scheme[1]);
  var preset = String(xml || "").match(/<a:prstClr\b[^>]*\bval="([^"]+)"/);
  if (preset) return presetColorToWireframeHex(preset[1]);
  return "";
}

function schemeColorToWireframeHex(value) {
  var key = String(value || "").toLowerCase();
  if (/dk|tx1|tx2/.test(key)) return "#111111";
  if (/lt|bg1|bg2/.test(key)) return "#FFFFFF";
  if (/accent/.test(key)) return "#BDBDBD";
  return "#D9D9D9";
}

function presetColorToWireframeHex(value) {
  var key = String(value || "").toLowerCase();
  if (/black|dk|dark/.test(key)) return "#111111";
  if (/white|lt|light/.test(key)) return "#FFFFFF";
  return "#BDBDBD";
}

function extractReferenceSlideHeroText(slides) {
  var targetSlide = referenceSlideAfterMainMarker(slides);
  if (!targetSlide) return "";
  var fragments = extractSizedTextFragmentsFromSlideXml(targetSlide.xml)
    .filter(function (fragment) {
      return isUsefulHeroFragment(fragment.text);
    });
  if (!fragments.length) return extractHeroTextFromReferenceBrief(targetSlide.text);

  var maxSize = fragments.reduce(function (max, fragment) {
    return Math.max(max, fragment.size || 0);
  }, 0);
  var selected = fragments.filter(function (fragment) {
    return maxSize ? fragment.size >= maxSize * 0.86 : true;
  });
  if (!selected.length) selected = fragments.slice(0, 1);

  selected.sort(function (a, b) {
    if (Math.abs(a.y - b.y) > 5000) return a.y - b.y;
    return a.x - b.x;
  });

  var lines = [];
  selected.forEach(function (fragment) {
    cleanHeroText(fragment.text).split(/\n+/).forEach(function (line) {
      var value = line.trim();
      var key = value.replace(/\s/g, "").toLowerCase();
      if (!value || lines.some(function (item) {
        return item.replace(/\s/g, "").toLowerCase() === key;
      })) return;
      lines.push(value);
    });
  });

  return cleanHeroText(lines.slice(0, 4).join("\n")) || extractHeroTextFromReferenceBrief(targetSlide.text);
}

function extractSizedTextFragmentsFromSlideXml(xml) {
  var fragments = [];
  var shapePattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  var shapeMatch;
  while ((shapeMatch = shapePattern.exec(xml))) {
    var shapeXml = shapeMatch[0];
    if (shapeXml.indexOf("<p:txBody") < 0) continue;
    var off = shapeXml.match(/<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/);
    var x = off ? Number(off[1]) : 0;
    var y = off ? Number(off[2]) : 0;
    var paragraphPattern = /<a:p\b[\s\S]*?<\/a:p>/g;
    var paragraphMatch;
    while ((paragraphMatch = paragraphPattern.exec(shapeXml))) {
      var paragraphXml = paragraphMatch[0];
      var defaultSize = largestFontSizeInXml(paragraphXml) || largestFontSizeInXml(shapeXml) || 1800;
      var pieces = [];
      var sizes = [];
      var runPattern = /<a:r\b[\s\S]*?<\/a:r>/g;
      var runMatch;
      while ((runMatch = runPattern.exec(paragraphXml))) {
        var runXml = runMatch[0];
        var runText = textValuesInXml(runXml).join("");
        if (!runText) continue;
        pieces.push(runText);
        sizes.push(largestFontSizeInXml(runXml) || defaultSize);
      }
      if (!pieces.length) {
        pieces = textValuesInXml(paragraphXml);
        sizes.push(defaultSize);
      }
      var text = normalizeFetchedText(pieces.join(""));
      if (!text) continue;
      fragments.push({
        text: text,
        size: sizes.reduce(function (max, size) {
          return Math.max(max, size || 0);
        }, 0),
        x: x,
        y: y + paragraphMatch.index
      });
    }
  }
  return fragments;
}

function textValuesInXml(xml) {
  var values = [];
  var textPattern = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
  var match;
  while ((match = textPattern.exec(xml))) {
    values.push(decodeXml(match[1]));
  }
  return values;
}

function textParagraphsInXml(xml) {
  var paragraphs = [];
  var paragraphPattern = /<a:p\b[\s\S]*?<\/a:p>/g;
  var match;
  while ((match = paragraphPattern.exec(String(xml || "")))) {
    var value = textValuesInXml(match[0]).join("");
    if (normalizeFetchedText(value)) paragraphs.push(value);
  }
  if (!paragraphs.length) return textValuesInXml(xml);
  return paragraphs;
}

function largestFontSizeInXml(xml) {
  var max = 0;
  var pattern = /<(?:a:)?(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz="(\d+)"/g;
  var match;
  while ((match = pattern.exec(xml))) {
    max = Math.max(max, Number(match[1]) || 0);
  }
  return max;
}

function isUsefulHeroFragment(value) {
  var text = normalizeFetchedText(value);
  if (!text || text.length < 2) return false;
  if (isNonHeroText(text)) return false;
  if (/https?:\/\/|pin\.it|www\./i.test(text)) return false;
  if (/^\d{1,2}[./-]\d{1,2}|^20\d{2}[./-]/.test(text)) return false;
  if (/^EVENT$/i.test(text) || /^PROMO$/i.test(text)) return false;
  return true;
}

function isNonHeroText(value) {
  var text = String(value || "");
  if (/^[※*]|쿠폰|수강권|유의사항|적용|발급|태그|확인해\s*주세요|확인해\s*주시/i.test(text)) return false;
  return /디자인\s*코멘트|코멘트|comment|디자인\s*방향|기획\s*방향|비주얼\s*방향|작업\s*가이드|가이드|설명|참고|레퍼런스|확인\s*필요|부탁드|부탁\s*드립니다|수정\s*요청|변경\s*요청|메인\s*배경|배경\s*컬러|컬러\s*반전|붉은\s*계열|긴급한\s*느낌|느낌\s*줄\s*수|(?:디자인|레이아웃|색상|컬러|위치|크기|폰트|간격|정렬).*(?:해주세요|해\s*주세요)/i.test(text);
}

function isDesignCommentText(value) {
  return /디자인\s*코멘트|design\s*comment/i.test(String(value || ""));
}

function isExactMainPageMarkerSlide(slide) {
  var lines = String(slide && slide.text || "")
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

function parseSlideRelationships(xml) {
  var rels = {};
  var pattern = /<Relationship\b([^>]+)>/g;
  var match;
  while ((match = pattern.exec(xml))) {
    var attrs = match[1];
    var id = xmlAttr(attrs, "Id");
    if (!id) continue;
    rels[id] = {
      target: xmlAttr(attrs, "Target"),
      type: xmlAttr(attrs, "Type")
    };
  }
  return rels;
}

function embeddedImageIds(xml) {
  var ids = [];
  var pattern = /<a:blip\b[^>]*r:embed="([^"]+)"/g;
  var match;
  while ((match = pattern.exec(xml))) {
    if (ids.indexOf(match[1]) < 0) ids.push(match[1]);
  }
  return ids;
}

function xmlAttr(attrs, name) {
  var match = String(attrs || "").match(new RegExp(name + "=\"([^\"]*)\""));
  return match ? decodeXml(match[1]) : "";
}

function normalizeZipPath(path) {
  var parts = [];
  String(path || "").split("/").forEach(function (part) {
    if (!part || part === ".") return;
    if (part === "..") {
      parts.pop();
    } else {
      parts.push(part);
    }
  });
  return parts.join("/");
}

function mimeForImagePath(path) {
  if (/\.png$/i.test(path)) return "image/png";
  if (/\.jpe?g$/i.test(path)) return "image/jpeg";
  if (/\.webp$/i.test(path)) return "image/webp";
  return "";
}

function bytesToBase64(bytes) {
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var output = "";
  for (var i = 0; i < bytes.length; i += 3) {
    var a = bytes[i];
    var b = i + 1 < bytes.length ? bytes[i + 1] : 0;
    var c = i + 2 < bytes.length ? bytes[i + 2] : 0;
    var triple = (a << 16) | (b << 8) | c;
    output += alphabet[(triple >> 18) & 63];
    output += alphabet[(triple >> 12) & 63];
    output += i + 1 < bytes.length ? alphabet[(triple >> 6) & 63] : "=";
    output += i + 2 < bytes.length ? alphabet[triple & 63] : "=";
  }
  return output;
}

function parseZipEntries(buffer) {
  var view = new DataView(buffer);
  var eocdOffset = findEndOfCentralDirectory(view);
  if (eocdOffset < 0) {
    throw new Error("Google Slides export 구조를 읽지 못했습니다.");
  }

  var totalEntries = view.getUint16(eocdOffset + 10, true);
  var centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  var entries = [];
  var offset = centralDirectoryOffset;

  for (var i = 0; i < totalEntries; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("Google Slides export 목록을 읽는 중 오류가 발생했습니다.");
    }

    var compressionMethod = view.getUint16(offset + 10, true);
    var compressedSize = view.getUint32(offset + 20, true);
    var uncompressedSize = view.getUint32(offset + 24, true);
    var fileNameLength = view.getUint16(offset + 28, true);
    var extraLength = view.getUint16(offset + 30, true);
    var commentLength = view.getUint16(offset + 32, true);
    var localHeaderOffset = view.getUint32(offset + 42, true);
    var nameBytes = new Uint8Array(buffer, offset + 46, fileNameLength);
    var name = new TextDecoder("utf-8").decode(nameBytes);

    entries.push({
      name: name,
      compressionMethod: compressionMethod,
      compressedSize: compressedSize,
      uncompressedSize: uncompressedSize,
      localHeaderOffset: localHeaderOffset
    });

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(view) {
  var minOffset = Math.max(0, view.byteLength - 0xffff - 22);
  for (var offset = view.byteLength - 22; offset >= minOffset; offset--) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  return -1;
}

async function readZipEntry(entry, buffer) {
  var view = new DataView(buffer);
  var offset = entry.localHeaderOffset;
  if (view.getUint32(offset, true) !== 0x04034b50) {
    throw new Error(entry.name + " 항목을 읽지 못했습니다.");
  }

  var fileNameLength = view.getUint16(offset + 26, true);
  var extraLength = view.getUint16(offset + 28, true);
  var dataOffset = offset + 30 + fileNameLength + extraLength;
  var compressed = new Uint8Array(buffer, dataOffset, entry.compressedSize);

  if (entry.compressionMethod === 0) return compressed;
  if (entry.compressionMethod !== 8) {
    throw new Error("지원하지 않는 Google Slides export 압축 방식입니다.");
  }
  if (typeof DecompressionStream === "undefined") {
    throw new Error("현재 Figma 환경에서 Google Slides export 압축 해제를 지원하지 않습니다.");
  }

  var stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function extractTextFromSlideXml(xml) {
  return normalizeFetchedText(textParagraphsInXml(xml).join("\n"));
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#x([0-9a-f]+);/gi, function (_, hex) { return String.fromCharCode(parseInt(hex, 16)); })
    .replace(/&#(\d+);/g, function (_, number) { return String.fromCharCode(parseInt(number, 10)); });
}

function slideNumber(name) {
  var match = name.match(/slide(\d+)\.xml$/i);
  return match ? Number(match[1]) : 0;
}


var STATIC_PINTEREST_THUMBNAILS = {"1063623637049116404":"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","1063623637049078082":"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","1063623637049018316":"/9j/4AAQSkZJRgABAQAASABIAAD/4QB0RXhpZgAATU0AKgAAAAgABAEaAAUAAAABAAAAPgEbAAUAAAABAAAARgEoAAMAAAABAAIAAIdpAAQAAAABAAAATgAAAAAAAABIAAAAAQAAAEgAAAABAAKgAgAEAAAAAQAAACSgAwAEAAAAAQAAAggAAAAA/+0AOFBob3Rvc2hvcCAzLjAAOEJJTQQEAAAAAAAAOEJJTQQlAAAAAAAQ1B2M2Y8AsgTpgAmY7PhCfv/iAdhJQ0NfUFJPRklMRQABAQAAAcgAAAAABDAAAG1udHJSR0IgWFlaIAfgAAEAAQAAAAAAAGFjc3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD21gABAAAAANMtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACWRlc2MAAADwAAAAJHJYWVoAAAEUAAAAFGdYWVoAAAEoAAAAFGJYWVoAAAE8AAAAFHd0cHQAAAFQAAAAFHJUUkMAAAFkAAAAKGdUUkMAAAFkAAAAKGJUUkMAAAFkAAAAKGNwcnQAAAGMAAAAPG1sdWMAAAAAAAAAAQAAAAxlblVTAAAACAAAABwAcwBSAEcAQlhZWiAAAAAAAABvogAAOPUAAAOQWFlaIAAAAAAAAGKZAAC3hQAAGNpYWVogAAAAAAAAJKAAAA+EAAC2z1hZWiAAAAAAAAD21gABAAAAANMtcGFyYQAAAAAABAAAAAJmZgAA8qcAAA1ZAAAT0AAAClsAAAAAAAAAAG1sdWMAAAAAAAAAAQAAAAxlblVTAAAAIAAAABwARwBvAG8AZwBsAGUAIABJAG4AYwAuACAAMgAwADEANv/CABEIAggAJAMBIgACEQEDEQH/xAAfAAABBQEBAQEBAQAAAAAAAAADAgQBBQAGBwgJCgv/xADDEAABAwMCBAMEBgQHBgQIBnMBAgADEQQSIQUxEyIQBkFRMhRhcSMHgSCRQhWhUjOxJGIwFsFy0UOSNIII4VNAJWMXNfCTc6JQRLKD8SZUNmSUdMJg0oSjGHDiJ0U3ZbNVdaSVw4Xy00Z2gONHVma0CQoZGigpKjg5OkhJSldYWVpnaGlqd3h5eoaHiImKkJaXmJmaoKWmp6ipqrC1tre4ubrAxMXGx8jJytDU1dbX2Nna4OTl5ufo6erz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAECAAMEBQYHCAkKC//EAMMRAAICAQMDAwIDBQIFAgQEhwEAAhEDEBIhBCAxQRMFMCIyURRABjMjYUIVcVI0gVAkkaFDsRYHYjVT8NElYMFE4XLxF4JjNnAmRVSSJ6LSCAkKGBkaKCkqNzg5OkZHSElKVVZXWFlaZGVmZ2hpanN0dXZ3eHl6gIOEhYaHiImKkJOUlZaXmJmaoKOkpaanqKmqsLKztLW2t7i5usDCw8TFxsfIycrQ09TV1tfY2drg4uPk5ebn6Onq8vP09fb3+Pn6/9sAQwACAgICAgIEAgIEBQQEBAUHBQUFBQcJBwcHBwcJCwkJCQkJCQsLCwsLCwsLDQ0NDQ0NDw8PDw8RERERERERERER/9sAQwEDAwMEBAQHBAQHEgwKDBISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhIS/9oADAMBAAIRAxEAAAH4Czj6Er5z31B5hqPLsjZHuv0D/Pb6V68fb/nY3h6jzrbc2/R2PN9P04XXHvhKaYemPo7nzFft4et1/m416nlGsXkCe1p/Ss9OX5ruqvVvPspPKn0l94/kA69hv1B8w+EWnbqdjt85zuOw5BFegUNAhSHKTTgceoMnlm9n4UNyW2BmxruxZefH23EEMtsjuvsb5O/WQj5N8V+8/lMj5R2ytf8A68fkR+1SaVvw798/DW2Pwrtkbpfvz5PuvY8P6Y+fKZ5pl85h6jmfJ9pa2yst3zW5oDKTstu04t+9675SamSRtqPbvbqqGsu++6U8SzltzP2Fg0vPotKD0nnK3a4mvs63wEMkW5g6eapK2CGWNu+4Ga9b4ClQV22TQvT9by3sLZcnU3/nGg3ZbuP2pwjC1+0+T5a2+4PlLxtO05xLn0Pd52p5Siz+V9889+PPsLyfpu/re5Z/U6/nHY+xdP8AC8PgLn3sjp4jmPNcXqWbip+m8m+eo+k/HjecN9sdCfX3x72NfUHhNdwbBjtkZz03KeudOfI1L7owfMttzuWzqPYOhPLHdtcx8323Oy3lvfVwxejqKo8UVOU2bGhaYpGHqtnzJ1Sm0Mqap2q4KI9IQRnTIdnWVbOmRqQwsmdMttVu/rbCgtnSKo8UVPXtLqu4pepqpr/WvIq//9oACAEBAAEFAslPJTyU8lPJTgj50+/fV+raILz6v7a0h8U+Gx4ZuH4a28bpvlnb3l1Bc2V9uL8abbcq2x7BfSbdue031/HIndPEibTd95u95Dit1ThO+brE/wCkF4idSbrcJlctJRJylqv7e6SlW2Rm63FdwnJAdp4cRcWUfhuJb/orG/6NRqe42fuN0m7ukpF5ch+/XT98u2uRcigCTFtCjF+h4g7yzXardnt+4Xil7R46DksvF7u7DxJLEQUnwNzHLBfyRKsrrPf0rCLn/GErWli4uC1yXKXzZO0S+WsXtoWbnbHeXFuUNMS1gxqC1RLQExqUyCk81bKlK+4kZKgsNpRPdmx3E7pYQ2hYJBF5dA+83SHJdXMwYBJtLOPli1snNEqNbgjklmj+qDxUtP8AsnPEhfirwbu/hdD2r/anJ4mvELn8S3W32/jnf7rfNne1f7UrS0ijuLqKKWP6y7UW20vav9qaDOgyST4fWsbk7Y9vhu47gfWuoM/WwVI8ZeIpPE1tJHJEravF0W3WC/G9ut/01tn/AE8ge8binc78FFBgVK5JCSgM0q6Eu9s47U97DcPc4I97t5Deqhkue0drcypVt941JUhTi5ZWk2SWJLEuUWeDs7EXKFbfBEgiyDm8J3CIJUhK7a23JcPuu+M7Zu5Mm8+L5NsuY5YprXfLi1gi8T3kZX4rv5AfEl07u6XeTp5dOjJXJaII2qDq7XdlPZHvZQwJtPdbURbjFHDdd6d0FKVQy2T51kUyy25S9l8L7hvq9y22Xblu126e7R+hLt+FRd7Rbb5sqd1uf9l5bh7Tsf6Mn3O53G6uNrvd3mtbG7vLvc0blEiwud4/RN/Fv9hMjwuhVquSe32W+V9Y3hxb95tN/wB0uti2CGS18VxxWfibetsvttf1fwQS7VPBbySOK2kmT7hM7fdNz24f0k317R4TF9ZHwmqJQ8NzVk8IyzI3fbhtl/Hf30KP0puT/Se5P9Kbk5JZZlxqoUbPsl1s36D2mOPxJHYW920Gi0+LdhWtPifwslXifd7bcI+1jvF3ag70t3FzNdSuFKFy+57XkbXay7m3s4ktAyVeeFJrTadlsIdy3HxLs1ptEjQUpV719HBNyVXF2q5777tu12m2bFaWl9uniLb9usFMGhVPUJlxBnqHYW9tO5Nv21MXu+24iKNrFFYrfU+p9To0cxSFxSZcqRmOQdoQMBi6AM1LNKwHoUlCmYY1NUaR2t64EVeDICSqmUSDhgplMlIzNDLd3Sry4hHRgC1RhLUlY7QcOUHyg5k4doAMCkPVyYkO3KaaFmSNL5sLWQpUMnLfvCH7yl5mdVzFNbTf/9oACAEDEQE/AWxdaZMkb9v1LGx9o0qQPPLcvEXawxZxImWThkDXCBKuShP+DTHLY5M5nExOljxr9oQQfGhiD5CNvpqKOmMgeXNkj7kY1pk635IZDGPTcfnbLrPk7v8ASf7EOKeQwiZxo/QlIR86VBrHqQ0WURIUWOOMTY0sNAaZs3UAS9uX5vx+Xq/fAyZ94r8qo8PuTYYus/U44TidpP8ArOPpMfTy9yPl/UM/nehhkOMz5H9D/vJl8/8AH+s/9gf95OPJiyQE4+C+zjPJiH2MX+KEADgJYyvSeSMKv1dw3bdJSquHd91fs/8A/9oACAECEQE/AXJilAAnxpiw/Yc1+EyBFy0MrHDQJuTuDg+Q+KhiEZdLcvzt/vD40+emcmTCZkwhQY45jnZ5TgzS4GL/AGB0/vasWPGIfhcH7xGErlj/ANiylciXkC61+4pBHnQSI8FO6udALSCNJ7q4cUJe3InSHSfHmAMuo5/wI6T4/wAfqf8AYOSGMTIjKwg0ylu9NIT2nxeoNOLp83UbjjF7edMWD4P2onJlnu9f8P8ArOPH8EB/GmP94/wOYYBkkMRO30/waAgNhjIxNhlklIUdJ/C9VjjKZr7RZ/OnBgzdRP28Qsv9x/I/7sllGRwylMeQAf8AB/V6OePppnLDFtPjzdhhPFtH2By/JYpdNkIkN22v8NPwvymSPUi/yPhj1maheeX/ABb/ALyYfDdXOAyCPn+o/wB5sPhvkAbhH/Yj/ebOXUwkYSkbH9UZclcSfdy/4xdxQyBGmDpMmaM5w8RR005YT1HoNMOD3IzluAr/AHjhGC8Bzbv837P/AP/aAAgBAQAGPwLi+L4vi+LRDwyIDllinVKYwnpKMSaqx9S+ZJcq0RH7KKkrX5cfJx2q5uYtYKvZpRNdPx7W9iVFGa+KdT+trup1kpWrHCSGMJSQofsk9R4tECLvJKj9LzY0rIGpBTr5UcW4bhcZyZqShOKUka9dfhXh2juYEhSx7OXk1WqLmKIJQpYRpTXhT8GmRN9DRIVliEg/iE1f8alripSwlQpqrjw7cy09scUsIEy0YjH7GgqmkxoMvj9jVKolQ/aPo8U6vmR6Ol4gKP7Y0U8sSv4F8sUQj9lOgfBouzKRmK0oz9Jw9aB/vU/4Qekp/Bm3yypTV4JkWAPKr9tX4v2z+L/er/EvKQlR9S6B5SpnSr05Zav33SP9KL4Lw8lKSU9q2MalkfssFcF1+DPMRcdXHi0i8jmKU8MniX9Fxzflw/a/0GBIlNPM1enCjX830kh05ivxesh+xVX7R7ZUB+eroUxD/hL/AEXoUV4fuP8Al54W+Cq8Ty8SP1ntVP8AC+WeLqp6Oh7a/cAYVzLkFOv7rWof009wumv7kf1MG25pQfOVGPaoeQkVX5umSg8ZlqUB6mvagf8AGUGtfPP+oOq0gfMr/uM6aV7Jih9pRoHkq8hSfTJen6nX323NPiv+445L+RMiJDQFOXEf2gO0H+7E/wALVGlMdKkOSa3jhjNOJBp/C4TdKgXy5jRUAIGqR6k9oP8Adif4WqTDzOrKJEZAh24ApWZX8A7Qf7sT/CyOWWcIzV23PBH0p9r1oO0V4iGRaUqCukHWjNNsu/1f8kvBW2XR/D/klxw29jcRFKyslQ41+QeEqSk+h0aLNUKlYedX+4kHyW68mX/cj/xdX+E1XiU45U0+T6hV14D0fTp+P9x9Qq9O6BFKmbJNTj5fD7hQiQjLinEFoubq6lEo8kxpo1SQqKgrWp0175xRqUPUBjGGT/BeKxQjsOcSE+dOL0Wv/BH910Upf2JH911iUsq+I/0e2VUCn7Ro81ctXwSqpf7tqlw6kR81SddE/N0DCrdKsT6GlXolf+E8lRn8X+ilfu6Yk6ZEehLKJhir0YgQlJCfV5IxSfhV4yEEfb/deiEfrZnWACfR9VX50fRX7X9IfwIfQRT4qHdKZ6dQyFCDp9n3OetVuT+yuuTz5lmdacTVlMSo1jjWP2dfv1UKj0eWIjPzW6LWk/CsjwjiAP7VT25O3x5EJyNTTRrgmTRUasT25kdKcHxS476xx64gFZej+lWEG4kqcBoHrcL/AMFqihPM5UgIKh8iwsRWwokDpH+i4re1hTJRA9kEl29vKlA6/KteBYtFREkJx1c8NIiSoEhaqflDzuJYkK/ZBdrMVFCCEZ/L4uPdLmvLTKaka8avWST/AAGu8t1FMC1pSVEagUFdHjFflWn7LjtRJDilCRqFV4PkWsuashpQ9pDL/pjywHbJJT9qgP4X7Uf+Gn+6zBaTKjFdQk6P/GpPxcd+Jygq8sXpdSfZG6e9y8Kfu3hJeLIP8hqswrPGmvDi+XFNIlI8gov9/L/hF/v5f8Iv9/L/AIRfMmUVK9S9dQ0GDki4VFpU65asKiTApQCEryUPI9R9riwjbkAIxFCOwJcwUogUjSitSk0TrpThXyaJZZioJOiBGdPI8Q7a3tFBSYknIhONVV4/hTvy0yYpHCiUn+FlSZlZH/YaGZpzVR+zslEqsEk6q40dPfNP91qf+OcNB9EXW3n5p9MSP4ewSzup5mFKhVOnj6tFncL5aVVqri4hZrWsLTXrTj2qoV+D5VDT0yNHmBr6gkP6WpPkVKJ/h7291ZKQoygeyvI189PJx2t9IIY1VqtWgGjg/R8yZuZHkrE1xPp2q6YIDpik/N0wR+HY+8qKflr/AAkMqRIoq8hQf8lOua604Y/6L1ZAej4vi+PYUkOnlro/X4vg9R3/AOGeg/gdCn9Yej4vU1/2/m9DR6Vr8u2jrU/7f2OlVfi+pR/F6OoI/Avin8C6ih+wtMxTlia0WKj7QzcLShBV5RpCU/gH/t/3H/t/3Hpkf9v5OtDTv5/7f2Pzfn2r8XX+p+wr8HVQUO1DT7X+V0NH5f7f2OoFHq/9v+6/9v8AusQxJzUTQCjMFzEYlp4pUCD+t//EADMQAQADAAICAgICAwEBAAACCwERACExQVFhcYGRobHB8NEQ4fEgMEBQYHCAkKCwwNDg/9oACAEBAAE/Ifc/N9z833Pzfc/N9z83eOW+JYqL1D4I5R2fsLDG4koYkJgXLvxZwmJjIHMugmOv+QS4yAjcMfmzB02xavOnRJlEjcJr+0CWMsTKSoRJR0RH3/xYPR0J78XhwAPWImS6GX+74uYs5yJLxuWKxiRffkbMd/8AJedoc/VmhI8RnhxTz0hzy4ZRYDnjPlevB3XFFHjG/wCE5BjmzbH4WCxbBwEH+6YNVNRg8djmagD7v4ZaxvR/nzR2H8v7okqBYjkmg+OQYH1e+PlX/wCi/wB3xhaTUdkv7sAJXgKCWIgOOxjWQcyPS+8y5CZ2Q9mn/CjcyM7TRgaSo81KcaC8qHPpjc/NcjCYl9d8jHj5pjP/AA1xROCWfxYNGLnb+9v6GGKkDT3/AN3lZP8AghqpD+VvNhvjzCvRx2SUXBmSldA81BPkQvUP+EeiPgfzcsfY/nixwI+Rocnj2VxzF9v6Lm7sJz/2EZZYzmggrxlKODj82LS9MBzyuL2PwzfUfX/ASwnCVfIOWVBgIma8NJxUgaEfP/IEyvRUPVxAQfabZcNIyS/FwCw4YY/Yf84FL168UzAGrV4nL9V3ke1aKHcegYToPJ/xwv8AlCxAAhD5+buTL8ke1TCg6FBNJ/1xjmXAe70gwSf+zWr/AC3hWxzO5QzaOTHNYQukDyIjr/mgN4iTpirrRR/Pu2kx852AYB4q8h5GX4a89CRIMs8RVJcEcP8AVCf8T8X/AOP/ANVFpAmy4jmkgq+Yv7dl/kKYkV8quS/tFaSYOjn/AIbAbGziK7e09l4/6s/7iPyKc08qu2fE4SP3dwWBFLnD/uvliQk30l2Vz6rhmBHE/wCArsAz9BSt8HZlLKhyr8IYYfkX/HY7Dnz4qXEB7J6LzCPzS8gEDQT6TFYcN819Hh91aIkcf4agRXlR/uojZgBE+Rp+7EHImuDbDKf5sOoUSP8Aa+kwJQDo+v8AawBhzhl9q9RXr/tohG/wixGOev8Abbs9dp/n/nNCkJ8ivk//AINIVC8EodEXmNipgD67K70gVXCYJhz/AKC8VZ5/6Jj3Mon8VYmWg2Y+qQ3fLj/FSzBhf9P/ABGEfOHDv5vGT3ZJPf8AxTiPK/5D/qrs06ZGA7xFkkkeZIVgXv5ruHPk/wC6+MHYFiBJ4phDEIe1+zVReeGByxRYkp7h5rqDChHHNgebwAYJMUy5zMsNzat1TSCIfpV5CgZYBn3f7oqQg1NAC+FBEoKsdaEHEeNEzHuxFS0Aw+T/AIczM+D0V/4Bx4/44NDNv8IvmrNnNaZ85/x80H2AYhjmaG1xEzfw3CfxH81buZEf7qLQFM1CeKQ4QCB9Df8AG/7v+d/3f8a/usEblVfy0YjJyXhtDja7Q8mchFQ7TW8MHk4unLwmXtedn/gcQEqngqecvX+xtcEpn0pSYNGyogoSQlISx/6FxJEk/Rou8nRJx1e5JEDh6IP+Q4QIGDzBrU0DrHmprJmEAyebBb/wUKnbFlDRarmEQ4aznOCRAuFXL5MyRjNf+CIk8o/ixlq//CmLDKi4lPsomDEI5+X/AA52ryuXhBvnBu7m2ICSVg5sK9h2lJQv/IuIx08UG1PB/wC0/vwbRSFkYv8Af/MYJwkoBPHIYney/i/qiYn5tTAj3tDyihmsuOf3v+M3/Gat5T80wzBHR6Uid/mqZK6LLkWV5rHJs+v7o8P1/pYEzPjNh/5X6uiEFYUQ33/SVWS+U1sB8D/bQIb/AAfzeMaimWL/AJw2Z+mtZZfmf6pCD/j1dm327eI95FI/Vj+D/PLOyXGTfxZkotnjyYl5tsfUUQfVmcefE/2vBfEYWcIk6FoSQ+3/ADTPf+dNiUn86Z9fz/8ALKAYPn/4f8nFPwJ/qlxP5f6qh3gpi8MyP+SA4ZqOqHfH+rlHH/OrBkfjSj4ijKTvqf7K9E76rTZ/FTkFSyq4AC7c/wAZifZq/wD/2gAMAwEAAhEDEQAAEMOMLJBDvA9zPp5ijT69ZCO6AOWQPVwIoQCWSYNoBLgOtQIPeFE4EIwHDjFG2yoswB+rUGt/AD3gJZgDpgFCKHGCJPJIJAHLDIFEP//EADMRAQEBAAMAAQIFBQEBAAEBCQEAESExEEFRYSBx8JGBobHRweHxMEBQYHCAkKCwwNDg/9oACAEDEQE/EFzuGQeTwgT+R3n1+0F+D6vP787+v48OCY/qf5JAD/l/XML5397aklwwMPg/iRl4/WMDRu/BsXOc/v4qU512EDsyOo26b6r8oSmLz5usl4D4uGyTQ8WWnAtd5Oj83wFSVDHk3hzfmxAFP4vy5tOoCm7jnJv2fwP4MbffHhjyu2/a/r+JzXPGTByBM3+n/b4w/mn9ph/dX+75pwx0DPOHFOHTyDmcfX4i7XameRw4Oe2dOrUU1UHg8O/B9tlDqjOX/n2tfS5hwj9U4esf1h8XsZ0BOPhNPiXOh+Rf+cQgHEsNg6ifmZ44Pmw42SGef4/981jTrnHx+c4Ge/nx4jn/APE//9oACAECEQE/EA3guUJdJ10Kfmab43Fcuk+f9v78cuWY9focf6z9v++Dg8+p8P3+zKrYfQvpZ+1xow1J1+XN+ZY9f2c/tbfBcN6N4P4LCaTpo4/lmRAk/L/pIjjGHNKbvf8Aj+sC2B9Mf6bint2DQ4fz6B8awmjz4qiw6TxekXi+BYqUgDON5/jxngg01w/JufFldAfu5/pYLhcc7N4f5tA5tu3B+RnhIp/Nv+kl13xFpFcCa6MP5zxIFkwBnDkHfG/nLh3d4HJ+2WC06e3XC8d534xqbKu5+v2vnj+/94h/ZD/XnOYAh4LkU+IUvPo5cL/w7vv54F44ct/mzrmjloHeP2iEdJ3u/wB477DGnyZn1+jORAOhyaZvzMbTn0vhpgT6D1+hcKj9qFRIjy7OH5iAUPzb/wBJ/wAyjq21wuBUfyd8z6w11D6vG9vDwRW7se939s/r4hmDuLi/b6n7XEUxDtrvz1mfz4GuSY5/+J//2gAIAQEAAT8Q/wDqr/8AVX/6q/8A1V/+qosdS5xSjJiaKcOqDTI2phVKE65WGqILkAqNpUB9wBtrkPkiV/4RhONnBQNk4QeatmJgBgjDGQ1DZ2AtIU5IMFADi2WtGR3PLjOYXj/k3Gt4SxfDtGOeaODhEFtXlUUj2rxRcWbBcxUlyYhQU2U6whHkDo9FhrqWSlDDlPfsq6FMbZ2Mu/M2SZEwxrMYD0TZrAdeOByUHXBQ/pnu9xeIuSoPhKQDSA/4I+btPttfTrDxeACIxfBq9tEjn2sV1OEBIKEt8eLCCzMoz4afVk0T4J/6sQ+RP9FGUSqVi4S8fNDU0H1zgYKPu2Ze/wA3aaKBn4Iv7rSOQqaMJUtBORAJVegu2syIzEGwb85ZgkDo6cepstzfhMEpALR/4aPslQdIju8ITUiXUcc7xcJUIgQgkmGCoCMCgPL2n27WxPUYidUPgUsOUjsH8KZixAZTspfIcPVzCGgsekD83vex3k8QzHzf8v5ur8pfwo1MgJM3ysCxJEp9rzJFklBkMX81VS6tlAIIz0fWfVYrchW2JAdBGTPdeWkIgScY5jxmUIUOi3kJj3/x6qPKvHCGrwBQgQ0ng/yq4gIMOl9CtZCQSyH8pYixIYRPySNn5P8Ah6sYmOIA/grgCfP/AFhSFAypejtqr+mAYAScMyoMscgifkeRcctCgiEGWgOIofv/AI4R0ohE7GlnjEos+Ws13BAZjJ5GWkZgHgEESYg/4AYyAJV9BX0kyWieGBJDM3jSQIGeQJD5KZJQcUdQs04kH/kvWKMpw11rzR00TxQFgwcKkzGpoQ6cB9mvk1IA0QOBATqZP+eglpsrYNgRvvZvgvbTwgQSqCPuwxJhJDOUHImPH/DJ/wCMKgNFSUlKnz5qT0D8cnhKGOYgImP+nyCgyDol3m6STFQxndQrIelKOQ6X/g5ByIGoKOoppq3es33e/wDiYQzmwdrV4KQeNbYSnyoySSAmVHBqWjyOOfNHxLAj8wc6bDR0yPxYBFGDpwModAOY8V95AP4oaGqATUX1I/ivdfIuTxEOaYiPSZfprd3dHB85P/CpQOUJs8d9FBLC+SKiocT/AKpJdSViBkkF7oiCeWICSABT9K4YpmdeoEzEdf8ASSAzDQlJCJDqwFET+V4OPHdTYyOoORHh/wCLW7jkegP5KJEpAMDpCGtk60CqTyknaJT30/PkpiOv+ImSPmiNc8bZulpyziBpU7Dijozxr/vQwGNgSe2BSkfCxQ6gAx9XlgPmCxOHGerHtqwEk2rNJSqvlaXsGAEg0oMmMZNeCGUjEkmimlNwSAmlXYB34srVWghNJDzQ/k7T8UlAumD9WZogyQhBEy9eb2HxYfu93r3Gf4rJk6BlDuYrmFDlz7hodp+Ie4hP/AUB3Yfe3DwygH07/wDgPo2EaU0Ac88TYuITBEWVEB2kx1TMuqAEoyShk5/6jBV9U8QRErCR/wBKUyUQHidH1VG+IkM1EonyzZoK8YS6UXHNKx9EIPSxp5J/44NwFhJTy0YDZjWYg8IDH5vdb4rU2ZCehp/6FJNs9m4JUEGRLm49WoWRIl8maaDNCOfid3gZ7g1lBljKCUV6eexOBbIfAUw5RE9MKmcvFclY6IwORjyVUejySiEjwvzWBextwTSdnal74QcWAbJDyxxNQGsRehgwhBhyxerXgwO0o4oQNDz/AFxW6wm5CUsJQ7rHe02I5CeIfui5GKhJIIklED91k68sRKyZ+/8AhgBAWVjmgf3QMIAxHBHCrP8AxfGiHC+Ek9hFSYH1c1Usk8FDIcOS/wD1X+qKgDGzeQzx4qRKoAiLurFLBV1TGuJaZKRpqcTzpahElB4FiJjmyeE3GZYAG3w3kf8AzBaN4M+4IJRWCnz85z7HpoJCkjQ1KSCYgjzSkSPAwXguiQ+LjJMZ4oj2O4iIKqst5pifgdox7TxGwkWIOxmhVgMAKQRb+C4eYsvrPRglCaP7/wCp4xw2zs7t3mvusCmgiX84D7pNxALIIMQzwf8ADoPOR3XxHMGteWEISGbE537yvblKF3pEiPcrSS0mY4zKuMgwnn/juoGk9tJ7BSnbZA3GN+SyfCY4wRJVICbB+d7LgNCDn/hk+clL70VXbUr2eeYzDp1Q3TvOFCiU08jibLEkT4/5CHSdo2KxQonMKjJFabRmxQkJAa90+Coy1xKIAww7/wANiamGV6TMoYgElxJ7FW0i4bM1/KS7YykBJIkexye/+HgKlWWQhBhHYvxXO6OdGSGE6BTZYjmZcS54JgmONinwHdS69eeKXEBAKRUFpCw+yo0QPtX4b+G1JWvmp+6EDCVhGRs55qCiXGJ8t382FSSEx212QEqnFeQnra2ojMHGPRauVUBv8v8ApWRHmdHGo2fLC9A7zp/81pGkwWUPnJpR2cSeDn/SuLB2Zw8828CQZmf5a2KnDK+5fxUVBCdXCBwPYf5TY0uMTGIczgWSSXMEkeU607SXFZEd7bcgJ5JL2zBzQz2RTV4UZMbzQqEWoeAs6ICART4Ukx7olQnQMxpMaOJjXXEJaQiQyuWATZuBZm+D/DaggxkYPwmGz5x2Ae5FSPXCXj5iP+CWBHTAunm0K4AmvMNYKBBN+bVuB3XT5/4IAaw1EBPIrpaUhD5s87LcKjP5pg4NQCv/ABXZJgZx2NWLCPNx8dp9VIltYDPjR0hBAMmf1v7U4QekR9QB+qlEBiDkkaCVkhuwfmhNIxzOSPHpUyJgrhElLABLVyp3824Cj31f/9k=","1063623637048993325":"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","1063623637048993315":"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","1063623637048983550":"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","1063623637048975012":"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","1063623637048974975":"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","1063623637048974912":"/9j/4AAQSkZJRgABAQAASABIAAD/4QCARXhpZgAATU0AKgAAAAgABAEaAAUAAAABAAAAPgEbAAUAAAABAAAARgEoAAMAAAABAAIAAIdpAAQAAAABAAAATgAAAAAAAABIAAAAAQAAAEgAAAABAAOgAQADAAAAAQABAACgAgAEAAAAAQAAAgigAwAEAAAAAQAAATIAAAAA/+0AOFBob3Rvc2hvcCAzLjAAOEJJTQQEAAAAAAAAOEJJTQQlAAAAAAAQ1B2M2Y8AsgTpgAmY7PhCfv/CABEIATICCAMBIgACEQEDEQH/xAAfAAABBQEBAQEBAQAAAAAAAAADAgQBBQAGBwgJCgv/xADDEAABAwMCBAMEBgQHBgQIBnMBAgADEQQSIQUxEyIQBkFRMhRhcSMHgSCRQhWhUjOxJGIwFsFy0UOSNIII4VNAJWMXNfCTc6JQRLKD8SZUNmSUdMJg0oSjGHDiJ0U3ZbNVdaSVw4Xy00Z2gONHVma0CQoZGigpKjg5OkhJSldYWVpnaGlqd3h5eoaHiImKkJaXmJmaoKWmp6ipqrC1tre4ubrAxMXGx8jJytDU1dbX2Nna4OTl5ufo6erz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAECAAMEBQYHCAkKC//EAMMRAAICAQMDAwIDBQIFAgQEhwEAAhEDEBIhBCAxQRMFMCIyURRABjMjYUIVcVI0gVAkkaFDsRYHYjVT8NElYMFE4XLxF4JjNnAmRVSSJ6LSCAkKGBkaKCkqNzg5OkZHSElKVVZXWFlaZGVmZ2hpanN0dXZ3eHl6gIOEhYaHiImKkJOUlZaXmJmaoKOkpaanqKmqsLKztLW2t7i5usDCw8TFxsfIycrQ09TV1tfY2drg4uPk5ebn6Onq8vP09fb3+Pn6/9sAQwACAgICAgIEAgIEBQQEBAUHBQUFBQcJBwcHBwcJCwkJCQkJCQsLCwsLCwsLDQ0NDQ0NDw8PDw8RERERERERERER/9sAQwEDAwMEBAQHBAQHEgwKDBISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhIS/9oADAMBAAIRAxEAAAH7F34nZr9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+J2r9sd+KThs/2l34nZdP2x34nav2x34nav2x34nav2x34nav2xX+JfpeXp/rfP5hueP6z9M2v5e+I9fzH7Yb8Ttt4/7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtX7Y78TtSts1ttRB7G22FttW21bbVttW21bbVttW21KuKXodeOjFc0ybbbLtttW21bbVP0z4Z9b+B9/wAs463l/K+t+deHmPsPxzbbTm22rbattq22rYgzbbC22rbattq22rbattq22rbattq22rbattq22rbattJo0SLbatbVJHy7fhfTvPt/NY7bl9fbattq22r3n6U5T1787/SeZ+QPrj87vc8RO2+o+R22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rvGVJ6P3+D5JtuD3ttq22rWdZ7xz6/bnQv8Al/z36r4o+eiD/Q/ldtt8ttq22rbattq22rbattq22rbattq22rbattqmNjbbC22rbattq22rbattq22rbattq22qfaPFvSOvi8/b9px3P2Qk2UhxooX3f8KfrR4nd2XwF9p/knzu32c/R+c2JdUhkbZbbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22roGH1J5iyeMb0KhXTm5sGcEWbMr3tXiXvHiO83g6sS2lxFd7+qPxJ9a/Oel8afKFox9zzrHq/e+p3X4vJe8us1B1vI1O2W0Tq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq1vUey10PENKNdJJXOCvf+h+YepDp8s89/RXmO3l+UWVD6Yp8mW4kzVLu1U/oF4D9Wfl54vfU+rch9Mev5zTqE0Kiv8JuHCv5NX7BNtq22rbattq22rbattq22rbattq22og9q22rbattqyk422wttq22rbat9AfP/Y6IGu7ylYc/ap6Tn36DuGPfdh9L9l59nJ8e8JyLwvXKt2esz9h8q9Sx16z5stWyL6v7R8ffSFzeied9d89Yuz8nsa0kWmAqk7VtsLbattq22rbattq22rbattq22rbattq22rbY2y5cIxl7K1zuRM88Hg7fFGtGiaV2fFFa9quvAund/pP2r4o+t9H93+VPrD5KCfMDj9LPHtV+Pzdp9oM359em+j32OnlnQ/Uf5282nFhdt+zmuKULXMNQuQZhulSc5ScshGcSytsVEydsh22rbattq22rbattq22rbattq22rbalrEvpU7lo49HCysqm27VsHBOhO/MsPYOr4PQ+Xqr7Ojzh8Tb6343nw8BP6Hy7LWXdUTY9x0/lHovQfv3z5rYHD5d+2vlX6cXTzL5w+muVk+ifzR+oPDrDzlp6XVtrwDV2159Wwz3WJru+7L3XHHwq99z4z3vlfJeC9X8x9rHiKPpKDy/omMEH4Xo7bZnbattq22rbattq22rbattq22rbatMYy1C2iuVs81avOeyP1rnidk/ebhVse53EZk6evqHQjmGfUn6zkbHqHvNr88/T3Vn0zPlLgc/R3begx4KnjvT/N2z99+A/s74iz2A0cteX2AtzASDGHmHC2mZH6WWlMhMM8xtm22wttq22rbattq22rbattq22rbattq22rbalpyTTthbbVpjUrJ6Hap7H133jsX5T7HrvKvUS95AYeHWPo3xHvJO1v/I2D8305yHzyzxx91b+GMl3+pPki7qcyy6Dmv115ur4d9x+uB87eJfN33u0V/wAUav8AZr4DZfmHbMmysZO2FttW21bbVttW21bbVttW21bbVtpNG2FttW21bbVtYeu6jxI/2F6V3p8Ydx9U33Vn8ueq+Q/TXm9lj1rKw8j0/JPiv9VfhHp4vEGcV/Tj6J+sHyP9Pc79IrnJxPRRzQzdR5T1p2vyS925Hzf1fmupsPoTiOn576y9Q/GX6e8r7j7B4L5Y8gx6P1R+J/klq6qTstttW21bbVttW21bbVttW21bbVttW21bbVttW1v7a5+e7X7t9z6W+Dff/o130px7zrGeqcxX+JfHuI/QDqPzs/QvlTk/A/qT5mzf6oF8+fNo6Pp75b47LPFsylPrLzj2+5w26XkvozjuTX5K436W5X0ubw83rxutOW+ZPob5z5K1+y/hsvT8/wDo58c+8fVfZ81+PG+6viLx/wBAZbYa7bVttW21bbVttW21bbVttW21bbVttWx+km5Sw9d9aunx33vqO5167TpGN/uSWMqPO7F8p/GWXJ9o/GPG7DPbYW+wfj5bL+mHyF43VsHAdGbyQUmMRuuv0s928E91i/UyWSRvgusZU6L+e/zd9F/N+DShaFlfZ3xfn5v2k8/+QfvJ+L8tPMv2S+H8vS+TctA022rbattq22rbattq22rbattqub49jdbi6r7ld77tuR7TPq7fuOK6XDp7p/8AJPy7pxfYvyDym6PL22Oe21bbVttW21aJ1Rpmo6pt9e5N7D6p8A8xpfpZvzVUW/Sbfm1mH6OVf56jZeg+cun5hBssYp21b3XwrS/rmD82vtrbzvN/jz7n8h9HyvmDdhx/m/RbbZbbbVttW21bbVttW21bbV6JZsbZfRf3jC4Te+6Xw3y0Z/QPg9Jn4IWnPlttW21bbVttSVKio21bbVrdj6928H0B676z8QcA+aOYctterbZNtssyJjCyk422wttq2mK3Qc/iv1y3+Wvb/c+TPw3e0/p8PjLf1ngPn/q6XRPn+jttW21bbVttW21bbV6dacJzydve8O0zcy0bHPbattq22rbapjattq22raVUgw+j157r7T8B+rfpPkDfnV7D88+d6DPveC+j/G+n8tP7q3F5TP0N21fCDz6w4mvk/bNZSdW21bbVttWIPQ9VvPD+4+i+d6oJ23u+Xy3H+pV3i+155nzH5329tkbbattq22rbatv0ky35t79JNX5t79JNX5t79JNX5t79JNX5t79JNX5t79JNX5t79JNX5t79JNX5t79JNX5tz+kcm/Oju/uS39Tyvmwv1Cx9nyfzt4v9Km3zH0P5tn/R3YdX5w79HtX5zo/RvV+bpP0e1fm3v0lg35t79JML829+kmr829+kmr829+kmr829+kmr829+kmN8AdV9qvfe834TH97R6fN8Ec5+je83b8yN+ljfwvU/NvfpJkP5t79JNX5t79JNX5t79JNX2Jtlttq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22rbattq22r//2gAIAQEAAQUC/pN4ef8ASbw8/wCk3h5/0m8PP+k3h5/0m8PP+k3h5/0m8PP+k3h5/wBJvDz/AKTeHn/Sbw8/6TeHn/Sbw8/6TeHn/Sbw8/6TeHn/AEm8PP8ApN4ef9JvDz/pN4ef9JvDz/pN4ef9JvDz/pL4ef8ASXw+/wCk3h5/0m8PP+k3h5/0m8PP+k3h5/0m8PP+k3h5/wBJvDz/AKTeHn/Sbw8/6TeHmPEWxLf6e2V/p3Zn+ntmf9JvD7/pN4ef9JvDz/pN4ef9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","1063623637048974910":"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","1063623637048974904":"/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAAEqADAAQAAAABAAACCAAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/8AAEQgCCAASAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAgICAgICBAICBAUEBAQFBwUFBQUHCQcHBwcHCQsJCQkJCQkLCwsLCwsLCw0NDQ0NDQ8PDw8PEREREREREREREf/bAEMBAwMDBAQEBwQEBxIMCgwSEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEv/dAAQAAv/aAAwDAQACEQMRAD8A5/8A4e3/ABZtoI5tO0TSJbj7PErRyw3KwiQAbwMXh4BHDbcnuBWTf/8ABXP49XkYguPD3h+IesP22NuRjqLgH6e9fDfwj0yCG0fXNU0Eajbru2TSJNtLqy4G5UlXA54CjPOWIGK7r4j6dpsmjSt4c0OOG7YuTcW8cjnZl921ZLaNQDuVQVwwAHcmu2OEi6XO569rP/Kxzuu+flUdO90fSN1/wVk+OMtzJKvh7w6QzsRlboHk+nnn+dQf8PYPjn/0Lvhz/vm6/wDj1fl0evFJXEdB/9D8r/BvxS8a+CUht9AudsUTM6wsiupLjDZBHPqAeh5rY1n41/EPV7VbG5uhHCquNqRIpPmcMSwGSccZzn3rzHS5lgu45Xk8oDq23dj8OtXdYvI7oR+XMJNoIICFcenXnmuhYqsocim7epm6VO93FXOdI5OOKSlPU0lc5of/0fx5TS9TbAW2mJPYI3+FSzafqMkhkW0kQHoFRscfWv6hdM/Zq+BcEUdxb+EdODoi5dYRwSoPUn0NOb9mf4BmMlvCGmlY8k4h3Be54BOK9L+z1tzo87+0NL+zZ/LU2nahuP7iX/vg/wCFJ/Z2of8APCX/AL4P+Ff0v3f7NHwMW7lVPCml4DsB+57Z+tV/+Gafgd/0Kml/9+v/AK9H9nP+YP7Qj/Kz/9L9RPG2kfEO9ih03wkbdLV4VefzjMrMxVQAGi6rjOQe9c14V0L4saVqaXN8bE2UwKXKQmdyUxgFRjaXyANx5xxmvr/T7eM2EH/XNP5Crf2aMcYruWZxUeV00zi+pSvpNnz/AHtgReSjb0kb+dVvsJ/u13OoQR/b5/8Arq/8zVTyErD67Ir6mj//0/3WktXvfDxs4sb5Lbau4sF3FeMlSGxnrg5rM8HaVq2j2E1rrM0E0pnLKbfzNqx7VAU+YzHcCDk9+OM1h+M7zxDZeGLRvD8kcTSGKOR5NwwhAzgryDjOPevKPg3rUp8YalpGhwXVxpSN5Ml5J8/+kKN5LOPlzkspwT2J61yLFUIqVGo7SbVtL3+e6+63exliKnsZ0nJaSdl669PK2va67np2oE/b5+P+Wr/zNVMn0q7qH/H/AD/9dX/map1Z0H//1P1J+Ifx4ufButjwS/gvW9agW2t5TeW0Ra2csocAEK2SpGD78Vzlp+1VqSXdvp8Hw88RKk8qoTFbsFTeQCxHlqOM19gadu+wQYz/AKpP5Crnze9ddPF4CMEp4S8ravmkr+ZzTo4lvStZdFZaHluof8f8/wD11f8Amap1c1D/AI/5/wDrq/8AM1TrzDsP/9X9vNbsdDfw2dU1l4LdbW03/ablzHDCFXO+Q7lAVTycnpXA/DzXfh5rN6LCHxFoetaqJnngTTLpGcRqAf8AVLKxbaSxJxjBGap/tCWH9rfATV9HZmRLuzjhkKMVPlkqZBkYIygNfPH7GvwA8F+HrY/E2aK4m1iC4uIbS5lupHVYZUCMPJzsB4YZOW5PTiumGDby+eMc1ZS5bfava+mm3zNvq8nD2y2R9a6h/wAf8/8A11f+ZqnVzUP+P+f/AK6v/M1TrziD/9b9Kfjz+0v8P/htDbeAr6yuNc1KeOMXNhat5bRRyRblLM6lW3g4wpJ55xXl/wAHP2xfhRpWo2/gKfQL3wxYzlpVurqUTJ5rkBVKopf5icA9Bjn1r7ll8E+DNdjh1DW9I0+8uDDGpluLaKVyAvA3OpOB25pkXw1+HMMizQ+H9JR0IZWWzgBBHIIITgg13Uq2WRpck41Lu17SVm11ty/d18zCcsdrCM48t9NH/mYmof8AH/P/ANdX/map1c1D/j/n/wCur/zNU68k6j//1/3Peyvbmwt5bK6mgMcH+rjEeJCVGNxdHIx2x685rM8Mw+JpzJda+9xbGOYoluzwSrJGFGHLRxKRuJOVzwR6V4P8XdZ+OWl6rYR/D4EaXLaQRl0WBiJ2Zg27zAWAxtxgY/Gtz4beLvG91Ja6frrG6vkjkhuxlEg80ScZZRkSbOgClTtYZFejPKeTCU67r026m0eb39E29NFstd7HLHFc9eVJQkuXd20121/I9A1D/j/n/wCur/zNU6uah/x/z/8AXV/5mqdeMdp//9D91hptlqelQwXqB18pPqOAeD1HIFV9N8K6HpVybuzhxIW37iScMRgnk9SO9SSXt1YaTbz2ts918i71RlUqoXJb5uvTGBkn0qPRtcudYck2UsESggySED5xj5dvDHIOcgY7ZzWKqVFHlUtAcIt8zWpyOof8f8//AF1f+ZqnVzUP+P8An/66v/M1TqSz/9H9stc8V+G9C0yDTPEDyKLu1xiNWJ2FQp+YdOvHOa0PBNj4as9F83wrvFtcyNOQ7OzbzhWzvJIPy8j1zXOePtJ0698FLqV5aJdPZRxygFNz7Pl3hcc9OcD0rN8CPrEerG2g0xrLTAGCYTYox0OPevCxObVKGJhg1TbjO2qvbrq9Oh6dLLfa4aWL5l7t97dLf5mnqH/H/P8A9dX/AJmqdXNQ/wCP+f8A66v/ADNU69U88//S/eXT/wDjwg/65J/IVb61gXP2/wD4R5X0x2SZIVddqK7NtXOwBiBlugJ6Vm+DrvxDf201z4gjngZZDHHFcJGjFQAfMBjJGGzjB5BBrDkbi59B8yvYwNQ/4/5/+ur/AMzVOrmof8f8/wD11f8Amap0ij//0/3AHii0sJoNGMM8kxtVmXYo2kDAIBYgZGcnt754r89V8Kf8FFW1BDF4ksnDS5IMtkI1QNkkr9j3bdvAAJPvjmv0UGki9ht5y2B5KKy5I3AYYdOmD6dc1o2djPZ+Ykb/ALuRixXrgscnk81NPkjFtyTfoDbbtY4S/wAfbps9fMb+Zqr8tWtQ/wCP+f8A66v/ADNU6xLP/9T947BlWwg3ED90vX6CrYdCcAg/jXG614bTxFolqiuI5oUV438uKQnKjKfvUcKGwASBmszwZ4avbOaTVNWhW2kV2jigCwEBQABIHjjUgsM8ZrJU04OfMvTqDm1LlsVtQ/4/5/8Arq/8zVOrmof8f8//AF1f+ZqnUFn/1f2c8feKP+EZ8G71kkhlntHjgliUMyzGI7MKSMnPIHfGK8V/Zp8ZeLb+6vvDOvwXrQxkzxXF7cmdlOERo8v+8O5t0gycLkrXv/iTwLo3jrQ7Oy1jeBAEljKHGGAHP144PaofCfw30vwjqL6lY3E8ryKVYStuBz3571yyxuJp/uKVNShL4m+nZxs1qtVqmrdDRZdgqi+tVK0lVjeyXwu9tHp5P79dkVNQ/wCP+f8A66v/ADNU6uah/wAf8/8A11f+ZqnVCP/W/dgX0WnaPDdTJLIojjGIY2kbkAZ2qCcDue1O0zWbXVgxtUnUL1MsTRjrjA3Ac+1MVbe40mOxuk3xzW4R16ZVlwR+IqvoGiaH4ctWstCtltYppTM6ISQZGAUtyTyQozWPIuRt7j1vpschqH/H/P8A9dX/AJmqdXNQ/wCP+f8A66v/ADNU6ko//9f92bZwLW3Ukj90vT6VZhbG0EnrXm2q/Fn4Z+FbuLQvEevWVlexW8Uj20sgEiq6gqSvUAjpmudl+NfwdutTtJYfF2nx7XCCPzsBmY8L1AO7OPmB9sVUcDi5RUoUpNd7P/IiWJoRdnNfejqdQ/4/5/8Arq/8zVOrmof8f8//AF1f+ZqnXKbH/9D4/wDFP7ePxf8AF+uzeIdY0rwvLPMAgaXTFkcRJxHGZC+5wgwAW54rBT9s74mRypKmi+E90biRCdIjOGU5BwXI4NZXwH+GvhTxPpM2u694u0fRnkmNmNNv3VJ5cBGWRSzqQhLEZAzwea9J+JXwU8A2Pg64l0zx/wCHZLi1gkvDaxSq8ly0KMVijPmklmbhcDBJGea7oV8WqXu1mkunM/yuc0qdFzs4K/p/wD0yL/gqx8dhEovdF0KebaPMlKTJvf8AibaHwMnnA4FSf8PWPjZ/0ANC/Kf/AOLr8vj1NJXBZHSf/9H8jfDPiAeH7rzzAkyvhZN2d2zOSF5xk9iRxXReLvHMXiNBFa2vkqcMxYndvUnBGDjGDjBz61Pomj3HiC0lvNG0j7RFbgea6DhOO+TVm78O31hbPeXejbIowSzkDAAOD39am5qqV9bnkTMNxpNwrqW1LRtxxaR9fQ/403+0tH/59I/yP+NMnlXc/9L8qdN8Pa/B4Ri8VyWizadLcyW8chukQ+ai7nHlB9/QZ3FQD6mqcv2tIHRLURtsZd4uc4wNxON3PB6V7n8Jv2b/AI9/FDwjbeJvAHhp9U06WWaGOcXltB5nlnEiBJJFf5XA5x2xXpF3+xp+1vp1jNf3ngqcJBBLI7i+tCQSCzvtEpJ4H3R+FLmXc1dOdk1FnwWetJX1av7EX7UciiRfCr4YZGbuzB59jNxS/wDDD/7Un/Qqv/4GWf8A8fo5l3F7Gp/Kz//T+fv2evjr468G/CTSPBvh7V00u2N/dI89zBC8EXnOHyNymWQjklUzwfvLivoC7/ak/aEfQL7Q9Y8SxmdLCSVnj060WKS2lUmNlBBYN5f3lLAgnI6V8pfBDRvC/iL4a6XZ3ulSX11Z3s11MloXElzF5uBHJ8wUoD1A52g8ccv1e38Ba3rl/pl5qL6XM0bSNp9ntuJVQkhIDIqFHwQC3deUPIzXgYvEyhUkoza+SaWtumvzelrbM9jCXckmrpLvvpe2v5Ls7aH3/Z+PPiBPaRTp47wrorD/AIksZ4Iz1xzVn/hN/iH/AND2f/BJH/hXz9Be/FCxgSyj+wOsKiMNIrKxCjALBogQTjkHn1qX+1/in/d039f/AI1Xmf2liu34H1P9l4H/AJ+P70f/1PyZtLDWrSIJZ6h5KlSMRyyKMN1GBjg9/Wq39gXQxi5i46fM3H6Vd8L3+nQOh1Cxjvgu8utxMYkIK/LypD8EZ4PJ61q61f6Q0Akg023t1KsCbe4MhyRhSQ5JGDzxUWkO7OeNtrqHYupOAOABPIAPpR5Gv/8AQTk/7/y1yxllyfmP50nmy/3j+dOz7iP/1flL4NfsmfF/44fB3T9W8EaJ4Z8ie/uvK1O9uZYb+Xy/3TROBujEaNymQCSPz77VP+CZ37Uemabc6pcab4c8uCKSZtl+xbaqZO3KgZGMjJ69eK+iP2NvG2veH/2Y9GjtvsVvYLqt/FPe3lp5yoSzSJGCxAOW6nICggZzxX1d4/8AjdqY+GVrZeO4NKuNUE8osVVmtbdZ4lmWKU/MCMIOAMoWO0kZGeKVSpTbbta/9dRxi5TVNdr+X3n5Kp/wTa/aXkQSbdEG4Zwb/kZ9cR07/h2v+0v6aH/4H/8A2uv0ktPiR4rmtIpn8RaCS6KxIswASRnj/SRVj/hYfin/AKGHQv8AwDX/AOSa4f7Wp/8APxfczp+pV/5D/9b5P+B3xV1PwH8I4LbTtV1K1k/tG5KQWEixJHuQEzzF1dWXKhQMcgEZG4V2PjfxBJ8QfhxrHxL1rX7C6ubK0haG0SParS/vIJDHtGw4ZfMyMZcA5I6/AR0K8kiETvGyddpfI/LGKmGj6itubRZUERGDGHO3g5+70681zTwFSTbu9WdMMXGKS5Voe92v7RevadbR6fBPqWyBFiX5B0QYH6CrH/DS/iT/AJ76l/3wK+bn0W83HMidf7//ANam/wBiXf8AfT/vv/61T/ZcPP8AD/I5OZ/zP7z/1/yh8LT6WBnVbGW/BDAAT+QA3GDuGc4GeMc568VpaodKWzJg06SFhuzIlx5uBg4+XjGCRz7dOaueEtQ8G2/g6807V9BXUtQupMQ35+07rVcKPlETiN8cthhnnByMVgyafpJik3WssX7uQq/kTHDbsr1bHC9zx6881pzNLRk2u9UcOZ5cnLn86Tz5f75/OrH9m6i3ItpiD/0zb/Cj+zNS/wCfaf8A79t/hUalH//Q9a/4J9+I/Dumfs5aQupaj5DwX9+zW2yTDEzEhiyDntwfSvr/AMZfGnwV/wAIPq06SNI6WMzSwLDiURMDHkxSFXCliAGIx9a/D39m/wCJOueH/Br+H44kFlBZalfrcSQqwWWJS+MnJcDAJ6Y6YNeoj4reMPHHgieeO7ttPSezeG5dIVQvEWY7GZVVmUMuPvDDY7VrLG1ZU/ZUmrrTVq34tJCjgqcJ+2rp8r10Tv8Acldn6qQ+IAIUCiUAKMDYOOP+ulS/8JD/ANdf++B/8cr5M07xHHFp8Eb+IdHJWJASdXsRkgDtnirn/CTQ/wDQwaP/AODew/xrD+34fyv8f8jP+yqn8/4H/9H8tPC/xO8S+EtITStHjhQL5n74qfNKy/eQtn7p7jHI610x+PPjg2clk0Nmwkj8ou8JZwuc8Etgdu3auV8N2er+ImRIZo4VZipklOEUAZyx6/kCa09U0DXdMt2uxPFPGqlt0R/LIYhufp9cUng6MveaOn2s+Xl5tC83x48dlifLs+v/ADw/+vTf+F7+O/8AnnZ/9+P/AK9eWtrt2WJ96b/bl3U/VcP2/APrNT+dn//S+A/hX+y58bfif4Ug8ceENIgvdKmllihee/gtwzxHbJ8kjq2AeM47V6Frv7DX7TGl6b/aFz4ctLNFVmEi6pa5cBSxVVMvzEgHAHJr6q/Y18Qazb/ArTrCzMZRLy8ADxo23dKT94gnkmvS7n46eMfEPh68awnhs7Cziu91vc28QmR0EiCRFZR94gtlS2F5wKwrZlg6cGkpOSdna1r/ADt+Z6ccun7ONSel1dXe/wAt/wBD86E/YQ/adkUSLoVuAwyM39qDz6jzKd/wwb+09/0A7b/wYWn/AMcr1uH9sDxFcwpcN4y0eAyKGMT2kpZMjO0kWTAkdCcn6mpf+GuvEH/Q76L/AOAc3/yDXV7Wj/z7n+ByclL+Y//T+Sf2edIjufDGm6trfii00fTba9d2juj/AKtkff8AKN4J8wgDCoevUGvoDx/qPwj8S694h1nVLi3n05dOMkc9vKFQzhQmFDjc/wA5bnHI55HFfKfwc+C/wV8X+Ch4o8e+N9O0u+kkkQaXM6xTRrG2ASzOufMHI4AHqa9C1X9nT9m5tEubnTviRpFvdRRu8QeVJtxQEgBVk/i6evsa58RkFXFR5+a19bqST6vv5nt0MZWhBKMFa3deXfW+h65D4m+EIhQDwDo8nyj5xkBuOv8Aq+/Wpf8AhJ/hF/0T7SPzP/xuvyfMUROSo/Kk8qL+6v5V5P8AYkv+f7+5HmfWq3835f5H/9T8gV06PaPmPSnHTYx1LV9L/Ab4b+CPHMtyvjeTUYf9X9j+wCNt4ztkLKwLEKSvIHfua9B+Lvwq+D/h3wbNqHhSbVBexSKqz3/lrAw3sjACMFix2MFGByK9OGWVZ4f6wrWOOWMgqnsup8Otp8e4/MetJ/Z8f941oN94/Wkrzjquf//V/KDTtSv9MJksJpIS6hX8t2TcvXDYIyPY1oan4l1zV4hb393PLGMEJJK7Lx04JI4qnpMumwTCfUVZwuCqBQQ31yR/I1q63eeHryFW01HilU4x5aIrLnvtY8++DWqnLl5b6EWOPb7x+tNpzfeP1ptSB//W+EfhmdKGhRC8j0J2Z3UnUid456tjn6e1drrv9gi1kjhi8K5Kud1qWJ4HT5v0r5is9JW7sRcI2G6YOMcfr0/Wkk0S7iRpGCkKM8ZOfpxX1eH4jVGhCj7BOyte/wCOx41XKuepKftHq77f8E589TSUrfeP1pK+XPYuf//X9dsf+CQ+l/ZI3i+IN6gZQ+F05AASM/8APxTJP+CRNtKwil8eXrIUZmJsYzhgRtXHn85GTntj3r9l9P8A+PCD/rkn8hVundisfhdc/wDBJbRYriSL/hOrr5XK/wDIOj7H/rvUP/DprRf+h6uv/Bcn/wAfr9h9Q/4/5/8Arq/8zVOi4z//2Q==","1063623637048974900":"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","1063623637048974899":"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","1063623637048974848":"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bLb9D9+eGr9D94d7ntgnKzKnK1JytScrUnK1JqLnUjLDBea5g6zVcT5roOsM4KMvAoy9SMvU2C5DTpSVVttW21aJ8kU+sT+aGx6v0v35oav0v+SfBko8bbLfbTX3T6VX2Hf5e2xG21bbVttW21bajZbENNPTk8S3hhYkr8pXLfMtqahhW6qef6Dn122RttqCEwadKSqttq22rcj1wAfzRkweD1Ntq22rbat1fKezFfs7bd/mbbVttW21bYtC21bl+o5nfKt18voy57dDo89uh1c9uh1c9uh1U3Ycl1uD7bc+221BCYNOlJVW21bbVttXzn8t/phx+HR+fe9T2HV5X2Psv07pl+ePN/T3mqaeV/UPmP02+fo226uLbas6avwXzN7kdAz6qHFFpnq2yxCVbA7bVttW21bbVV2mzDbZTttQQmDTpSVVttW21bbVttW21bbVttW2ip21bS/Bh9OR9to7bUxr7ukZNtmXbattq22rbattq22rbattqCEwadKSqttq22rYa6nbVttW21bbVuT6yag7tyrQlk0E+f0Tk1rkqV9h1xDyrjMm2xGyKarzZAl7juE5fS9so/nSg873PezeSd5H2TRPufIbbVttQQmDTpSVVttW21UVtQV/m6dtuOkHsNx+N2G48Yu03E9Tsr1aN2o/ZJwttjbbUZywwpjY221bbVAB+b4dfqnAtrPHq+bd13I/Mfoe2tVe6+iKnpfqPzzbbu8fbattqCEwadKSqttq22rh7jm/QfA6F7b3ufbasE2F552dM78Tovtt7vPttW21bbVttW21bbVttXFfN30Z85/N/ebbeX9Euzqc2fvHpfm/pf1f5vtt1+bttW21bbUEJg06UlVbbVtmq3Eeg8Z2fl7bbetjttW21VNT1HFeNt3O29nGqtGzqts7ppnmpnjHplnc0zxRVttTD5S+vOP833fmnX9B8397tsuntQ/Gd3+N9izy/UfTfn+2z5bbVttQQmDTpSVVttWp7iv52r+gbuK226F22rbatxHb0/BpcbbvzypvaqbR1q22qmszattqpg3+rl9aVdbZiD5d4x635R8t+ij23D7G21e/ekcR2/1/5htt08G21bbUEJg06UlVbbVttW21bbVttW21bbVttTjoOdvqJmGp/mGp/mGp/mGp/mC6VQ3lHWQvC8w573DcfrfILf6+rPO9z5b9M9uL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","1063623637048887426":"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"};

function attachPinterestOutputReferences(analysis) {
  var candidates = staticPinterestCandidates();
  var briefRefs = normalizeReferenceList(analysis.referenceSlideImages || []);
  var concepts = safeArray(analysis.concepts, []).map(function (concept, index) {
    var references = buildPinterestReferencesForCanvas(concept, candidates, index);
    if (briefRefs.length) {
      references.mainVisual = [briefRefs[index % briefRefs.length]].concat(references.mainVisual).slice(0, 4);
      references.texture = [briefRefs[(index + 1) % briefRefs.length]].concat(references.texture).slice(0, 2);
      references.angle = [briefRefs[(index + 2) % briefRefs.length]].concat(references.angle).slice(0, 2);
      references.background = [briefRefs[(index + 3) % briefRefs.length]].concat(references.background).slice(0, 2);
      references.material = references.texture;
      references.composition = references.angle;
      references.landingLayout = references.background;
      references.layout = references.background;
    }
    return Object.assign({}, concept, {
      references: references,
      referenceSearchStatus: briefRefs.length ? "pptx-reference-slide + pinterest-board" : "pinterest-board",
      referenceSearchError: "",
      pinterestBoardSource: DEFAULT_PINTEREST_BOARD_URL
    });
  });

  return Object.assign({}, analysis, {
    concepts: concepts,
    pinterestBoard: {
      requestedUrl: DEFAULT_PINTEREST_BOARD_URL,
      resolvedUrl: DEFAULT_PINTEREST_BOARD_URL,
      title: "Pinterest board",
      count: candidates.length,
      error: ""
    }
  });
}

function staticPinterestCandidates() {
  var values = [
    ["Black Friday campaign reference", "https://i.pinimg.com/originals/3b/7a/df/3b7adf1359daa294e4bd0fe89aea5d13.png", "1063623637049116404", "black friday neon dark sale editorial"],
    ["Birthday campaign reference", "https://i.pinimg.com/originals/87/5b/7a/875b7a308319bda9e4e37392a321f83e.png", "1063623637049078082", "celebration warm badge"],
    ["Long landing page reference", "https://i.pinimg.com/originals/bc/aa/a3/bcaaa370cc73603aea271c3bea34bb41.jpg", "1063623637049018316", "landing commerce layout editorial"],
    ["Kream member event reference", "https://i.pinimg.com/originals/6f/24/07/6f24076c31a2510a2bf33944cb6bbfc3.png", "1063623637048993325", "premium dark event"],
    ["Product page visual reference", "https://i.pinimg.com/originals/a3/eb/cb/a3ebcb796780759599dfd9ef68d48cae.jpg", "1063623637048993315", "product card commerce editorial"],
    ["Promotion page reference", "https://i.pinimg.com/originals/43/a3/13/43a313e17cb57c4c927dad6b5a118df7.jpg", "1063623637048983550", "promotion page module layout"],
    ["Coupon event visual reference", "https://i.pinimg.com/736x/97/3c/31/973c316982dcefa987d35d617c9ca07c.jpg", "1063623637048975012", "coupon ticket benefit event"],
    ["Sale app visual reference", "https://i.pinimg.com/736x/31/cb/56/31cb56ad7393e3be0bc5b36ec81a67f1.jpg", "1063623637048974975", "sale app benefit module"],
    ["Black Friday sale reference", "https://i.pinimg.com/736x/68/51/bf/6851bf657a0ea215e93766ec3ffbde53.jpg", "1063623637048974912", "black friday sale dark neon"],
    ["Rocket event visual reference", "https://i.pinimg.com/736x/f5/dd/16/f5dd16fde46606f399b3622e39c68082.jpg", "1063623637048974910", "dynamic object event"],
    ["Interface line visual reference", "https://i.pinimg.com/originals/d3/c9/d5/d3c9d5e76bbf7e47a9246a5a39ffb2c2.png", "1063623637048974904", "interface line dark premium"],
    ["Mobile game promotion reference", "https://i.pinimg.com/736x/f3/39/f6/f339f6d08654c96ce9a2a03b7c97f2db.jpg", "1063623637048974900", "game colorful event"],
    ["Confetti game event reference", "https://i.pinimg.com/736x/88/69/17/8869171b53b124a78413942d6eea04e4.jpg", "1063623637048974899", "confetti bright playful"],
    ["Editorial home poster reference", "https://i.pinimg.com/736x/d7/75/b4/d775b4a7da92070fc2413098ca0b0c7d.jpg", "1063623637048974848", "home editorial poster commerce"],
    ["Orange dessert campaign reference", "https://i.pinimg.com/originals/66/9b/a3/669ba39e80ccaa3758ed95724ffcbb5e.png", "1063623637048887426", "orange warm campaign"]
  ];

  return values.map(function (item, index) {
    return {
      title: item[0],
      description: item[0],
      imageUrl: item[1],
      pinId: item[2],
      tags: String(item[3] || ""),
      inlineImageDataUrl: staticPinterestThumbnailDataUrl(item[2]),
      fallbackImageUrls: [item[1]],
      pageUrl: "https://www.pinterest.com/pin/" + item[2] + "/",
      order: index
    };
  });
}

function staticPinterestThumbnailDataUrl(pinId) {
  var data = STATIC_PINTEREST_THUMBNAILS[String(pinId || "")];
  return data ? "data:image/jpeg;base64," + data : "";
}

function buildPinterestReferencesForCanvas(concept, candidates, conceptIndex) {
  var used = {};
  var baseOffset = conceptIndex * 4;
  var mainVisual = pickPinterestCandidateReferences(candidates, concept, "mainVisual", used, 1, baseOffset);
  var texture = pickPinterestCandidateReferences(candidates, concept, "texture", used, 1, baseOffset + 1);
  var angle = pickPinterestCandidateReferences(candidates, concept, "angle", used, 1, baseOffset + 2);
  var background = pickPinterestCandidateReferences(candidates, concept, "background", used, 1, baseOffset + 3);

  return {
    mainVisual: mainVisual,
    texture: texture,
    angle: angle,
    background: background,
    material: texture,
    composition: angle,
    landingLayout: background,
    typography: [],
    layout: background
  };
}

function pickPinterestCandidateReferences(candidates, concept, category, used, limit, fallbackOffset) {
  var scored = scorePinterestCandidates(candidates, concept, category, used, fallbackOffset);
  var selected = scored.slice(0, limit).map(function (item) {
    used[item.candidate.imageUrl] = true;
    return pinterestCandidateToCanvasReference(item.candidate, category);
  });

  if (selected.length < limit) {
    var relaxed = scorePinterestCandidates(candidates, concept, category, {}, fallbackOffset);
    for (var i = 0; i < relaxed.length && selected.length < limit; i++) {
      selected.push(pinterestCandidateToCanvasReference(relaxed[i].candidate, category));
    }
  }

  return selected;
}

function scorePinterestCandidates(candidates, concept, category, used, fallbackOffset) {
  return candidates
    .filter(function (candidate) {
      return candidate && candidate.imageUrl && !used[candidate.imageUrl];
    })
    .map(function (candidate, index) {
      return {
        candidate: candidate,
        score: scorePinterestCandidateForConcept(candidate, concept, category) + scorePinterestCandidateOrder(index, fallbackOffset) + (candidate.inlineImageDataUrl ? 50 : 0)
      };
    })
    .sort(function (a, b) {
      return b.score - a.score;
    });
}

function scorePinterestCandidateOrder(index, fallbackOffset) {
  var distance = Math.abs(index - (fallbackOffset || 0));
  return 0.5 / (distance + 1);
}

function scorePinterestCandidateForConcept(candidate, concept, category) {
  var haystack = [candidate.title, candidate.description, candidate.tags, candidate.imageUrl].join(" ").toLowerCase();
  var tokens = pinterestConceptTokens(concept, category);
  var tokenScore = tokens.reduce(function (score, token) {
    var normalized = String(token || "").toLowerCase();
    if (normalized && haystack.indexOf(normalized) >= 0) return score + Math.min(8, Math.max(2, normalized.length / 2));
    return score;
  }, 0);
  return tokenScore + curatedPinterestPriority(candidate, concept, category);
}

function curatedPinterestPriority(candidate, concept, category) {
  var id = String(candidate.pinId || "");
  var title = String(concept.title || "");
  var copy = concept.campaignCopy || {};
  var text = [title, concept.koreanTitle, concept.mainAsset, concept.description, copy.headline, copy.benefit, copy.action].join(" ");
  var sets;

  if (/Typographic|Typography|Poster|Campaign/i.test(title)) {
    sets = {
      mainVisual: ["1063623637048974912", "1063623637049116404", "1063623637048974904"],
      texture: ["1063623637048974904", "1063623637048974912", "1063623637048993325"],
      angle: ["1063623637049018316", "1063623637048974848", "1063623637048983550"],
      background: ["1063623637048974904", "1063623637049018316", "1063623637048993325"]
    };
  } else if (/Offer|Commerce|Benefit|Module|Lucky|Draw|럭키|혜택|쿠폰|ticket/i.test(title)) {
    sets = {
      mainVisual: ["1063623637048975012", "1063623637048974975", "1063623637048983550"],
      texture: ["1063623637048975012", "1063623637048974975", "1063623637048974912"],
      angle: ["1063623637048983550", "1063623637049018316", "1063623637048974848"],
      background: ["1063623637049018316", "1063623637048983550", "1063623637048993315"]
    };
  } else if (/Curated|Curation|Board|29CM|큐레이션|카테고리/i.test(title)) {
    sets = {
      mainVisual: ["1063623637048974848", "1063623637048993315", "1063623637049018316"],
      texture: ["1063623637048993315", "1063623637048974848", "1063623637048887426"],
      angle: ["1063623637049018316", "1063623637048983550", "1063623637048974848"],
      background: ["1063623637049018316", "1063623637048983550", "1063623637048993315"]
    };
  } else if (/Editorial Sale|Summer|Black|블프|black friday|세일/i.test(text)) {
    sets = {
      mainVisual: ["1063623637049116404", "1063623637048974912", "1063623637048974848"],
      texture: ["1063623637048974912", "1063623637049116404", "1063623637048974904"],
      angle: ["1063623637048974848", "1063623637049018316", "1063623637048983550"],
      background: ["1063623637049018316", "1063623637048974904", "1063623637048993325"]
    };
  } else {
    sets = {
      mainVisual: ["1063623637048993325", "1063623637049018316", "1063623637048974848"],
      texture: ["1063623637048974904", "1063623637048993325", "1063623637048993315"],
      angle: ["1063623637049018316", "1063623637048983550", "1063623637048974848"],
      background: ["1063623637049018316", "1063623637048983550", "1063623637048993315"]
    };
  }

  var priority = (sets && sets[category]) || [];
  var index = priority.indexOf(id);
  if (index >= 0) return 120 - index * 18;
  if (/playful|game|confetti/i.test(candidate.tags || "") && /Editorial|Curated|29CM|블프/i.test(text)) return -30;
  return 0;
}

function pinterestConceptTokens(concept, category) {
  var keywords = concept.referenceKeywords || {};
  var asset = concept.asset || {};
  var values = [];

  if (category === "mainVisual") {
    values = safeArray(keywords.mainVisual, []).concat([concept.mainAsset, concept.visualDirection, concept.description, concept.hook]);
  } else if (category === "texture") {
    values = safeArray(keywords.material, []).concat([asset.texture, "texture", "material"]);
  } else if (category === "angle") {
    values = safeArray(keywords.composition, []).concat([asset.angle, "composition", "layout", "angle"]);
  } else {
    values = safeArray(keywords.landingLayout, []).concat([asset.background, "background", "stage"]);
  }

  var seen = {};
  return values.join(" ").split(/[^0-9a-zA-Z가-힣]+/).filter(function (token) {
    if (!token || token.length < 2 || /^(and|the|for|with|page|event|mobile|visual|hero)$/i.test(token)) return false;
    var key = token.toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).slice(0, 36);
}

function pinterestCandidateToCanvasReference(candidate, category) {
  var label = {
    mainVisual: "대표 이미지",
    texture: "질감",
    angle: "구도",
    background: "배경"
  }[category] || "레퍼런스";

  return {
    title: candidate.title || "Pinterest board pin",
    source: "Pinterest board",
    imageUrl: candidate.imageUrl,
    pinId: candidate.pinId || "",
    inlineImageDataUrl: candidate.inlineImageDataUrl || "",
    fallbackImageUrls: safeArray(candidate.fallbackImageUrls, []),
    pageUrl: candidate.pageUrl,
    keyword: label,
    whyMatched: referenceMatchReason(candidate, category)
  };
}

function referenceMatchReason(candidate, category) {
  var categoryText = {
    mainVisual: "대표 이미지",
    texture: "질감",
    angle: "구도",
    background: "배경"
  }[category] || "레퍼런스";
  var tags = String(candidate.tags || "").split(/\s+/).filter(Boolean).slice(0, 3).join(", ");
  return categoryText + " 기준: " + (tags || "보드 톤") + " 일치";
}

async function createAnalysisFrame(rawAnalysis, targetPage) {
  var analysis = attachPinterestOutputReferences(normalizeAnalysis(rawAnalysis));
  var root = figma.createFrame();
  root.name = "프로모션 비주얼 분석안 - " + analysis.promotionName;
  root.resize(3080, 1720);
  root.fills = [paint("#FFFFFF")];
  root.clipsContent = false;
  root.x = 0;
  root.y = 0;
  (targetPage || figma.currentPage).appendChild(root);

  var infoTitle = createText("🧐프로모션 정보 분석", 22, true, "#000000", {
    width: 520,
    weight: "bold"
  });
  root.appendChild(infoTitle);
  infoTitle.x = 42;
  infoTitle.y = 63;

  var info = createPromotionInfo(analysis);
  root.appendChild(info);
  info.x = 65;
  info.y = 131;

  var divider = createRectangle("섹션 구분선", 3080, 4, "#D9D9D9", 0);
  root.appendChild(divider);
  divider.x = 0;
  divider.y = Math.ceil(info.y + info.height + 42);

  var conceptTitle = createText("🧐디자인 컨셉 제안", 22, true, "#000000", {
    width: 520,
    weight: "bold"
  });
  root.appendChild(conceptTitle);
  conceptTitle.x = 42;
  conceptTitle.y = divider.y + 74;

  var conceptRow = autoFrame("컨셉 카드 리스트", "HORIZONTAL", {
    width: 3015,
    gap: 30,
    fill: null,
    padding: 0
  });

  for (var i = 0; i < analysis.concepts.length; i++) {
    conceptRow.appendChild(await createConceptCard(analysis.concepts[i], i));
  }

  root.appendChild(conceptRow);
  conceptRow.x = 65;
  conceptRow.y = conceptTitle.y + 57;

  var rootHeight = Math.max(1500, Math.ceil(conceptRow.y + conceptRow.height + 80));
  root.resize(3080, rootHeight);
  return root;
}

async function createWireframeFrame(rawDeck, targetPage) {
  var deck = normalizeWireframeDeck(rawDeck);
  var targetWidth = wireframeOutputTargetWidth();
  var root = figma.createFrame();
  root.name = "기획안 페이지 목업 - " + truncatePageName(deck.sourceName);
  root.resize(targetWidth, 1200);
  root.fills = [paint(wireframeOutputStyle("rootBackground", "#FFFFFF"))];
  root.clipsContent = false;
  root.x = 0;
  root.y = 0;
  (targetPage || figma.currentPage).appendChild(root);

  var currentY = wireframeOutputStyle("rootTopGapPx", 60);
  for (var i = 0; i < deck.slides.length; i++) {
    var pageMockup = createWireframeCanvas(deck.slides[i], i);
    root.appendChild(pageMockup);
    pageMockup.x = 0;
    pageMockup.y = currentY;
    currentY += pageMockup.height + (i === deck.slides.length - 1
      ? wireframeOutputStyle("rootBottomGapPx", 60)
      : wireframeOutputStyle("slideGapPx", 96));
  }

  var rootHeight = Math.max(900, Math.ceil(currentY));
  root.resize(targetWidth, rootHeight);
  return root;
}

function createWireframeSlideCard(slide, index, total) {
  return createWireframeCanvas(slide, index);
}

function wireframeOutputTargetWidth() {
  var value = WIRE_OUTPUT_POLICY && Number(WIRE_OUTPUT_POLICY.targetWidthPx);
  return isFinite(value) && value > 0 ? value : 1080;
}

function wireframeOutputStyle(key, fallback) {
  var styles = WIRE_OUTPUT_POLICY && WIRE_OUTPUT_POLICY.styles || {};
  var value = styles[key];
  return value === undefined || value === null || value === "" ? fallback : value;
}

function createWireframeCanvasFromElements(slide, index) {
  var sourceWidth = slide.width || 12192000;
  var sourceHeight = slide.height || 6858000;
  var width = wireframeOutputTargetWidth();
  var isFallbackLayout = false;
  var elements = safeArray(slide.normalizedSlideElements, slide.elements || []).map(function (element, order) {
    element._order = order;
    return element;
  }).sort(function (a, b) {
    var zA = isFinite(Number(a.zIndex)) ? Number(a.zIndex) : a._order;
    var zB = isFinite(Number(b.zIndex)) ? Number(b.zIndex) : b._order;
    if (zA !== zB) return zA - zB;
    return a._order - b._order;
  });

  var renderElements = cleanWireframeRenderElements(elements, sourceWidth, sourceHeight);
  var scale = width / Math.max(1, sourceWidth);
  var height = Math.max(140, Math.round(sourceHeight * scale));
  var offsetX = 0;
  var offsetY = 0;

  var canvas = figma.createFrame();
  canvas.name = "페이지 목업 - Slide " + slide.number;
  canvas.resize(width, height);
  canvas.fills = [paint(wireframeOutputStyle("canvasBackground", "#FFFFFF"))];
  canvas.clipsContent = true;

  renderElements.forEach(function (element, elementIndex) {
    var x = Math.round(element.x * scale + offsetX);
    var y = Math.round(element.y * scale + offsetY);
    var w = Math.max(2, Math.round(element.w * scale));
    var h = Math.max(2, Math.round(element.h * scale));
    if (x > width || y > height || x + w < 0 || y + h < 0) return;

    if (element.type === "svg" && element.svg) {
      var svgMockup = createGrayscaleSvgMockup(element.svg, w, h) || createWireframePlaceholderNode("원본 슬라이드 영역", w, h);
      svgMockup.x = x;
      svgMockup.y = y;
      canvas.appendChild(svgMockup);
      return;
    }

    if (element.type === "ellipse") {
      var ellipseFill = wireframeElementFill(element) || "#E8E8E8";
      var ellipse = figma.createEllipse();
      ellipse.name = "와이어 원형 요소";
      ellipse.resize(w, h);
      ellipse.x = x;
      ellipse.y = y;
      ellipse.fills = [paint(ellipseFill)];
      if (element.stroke) {
        ellipse.strokes = [paint(element.stroke)];
        ellipse.strokeWeight = 1;
      }
      canvas.appendChild(ellipse);
      if (element.text) {
        var ellipseTextWidth = Math.max(20, w - 12);
        var ellipseTextSize = wireframeFittedFontSize(element, scale, ellipseTextWidth, Math.max(12, h - 10), isFallbackLayout);
        var ellipseText = createText(element.text, ellipseTextSize, ellipseTextSize >= 24, textColorForElement(element, ellipseFill), {
          width: ellipseTextWidth,
          align: textAlignForElement(element) || "CENTER",
          minLines: element.text.split(/\n/).length,
          weight: ellipseTextSize >= 24 ? "bold" : "medium"
        });
        ellipseText.x = x + 6;
        ellipseText.y = wireframeTextY(y, h, ellipseText.height, true, element);
        canvas.appendChild(ellipseText);
      }
      return;
    }

    if (element.type === "line") {
      var lineBox = createRectangle("와이어 라인", Math.max(1, w), Math.max(1, h), element.stroke ? grayscaleColorForHex(element.stroke, false) : "#555555", 1);
      lineBox.x = x;
      lineBox.y = y;
      canvas.appendChild(lineBox);
      return;
    }

    if (element.type === "image") {
      var imageNode = createWireframeImageNode(element, w, h);
      if (imageNode) {
        imageNode.x = x;
        imageNode.y = y;
        canvas.appendChild(imageNode);
        return;
      }
      var pictogram = createWireframePlaceholderNode("이미지 영역", w, h);
      pictogram.x = x;
      pictogram.y = y;
      canvas.appendChild(pictogram);
      return;
    }

    var hasText = Boolean(element.text);
    var hasBox = element.fill || element.stroke || !hasText;
    var fill = wireframeElementFill(element);
    var textColor = readableTextColor(fill || "#FFFFFF");

    if (hasBox) {
      var rect = createRectangle("와이어 박스", w, h, fill || "#F2F2F2", element.fill ? 1 : 0.92);
      rect.x = x;
      rect.y = y;
      rect.cornerRadius = element.radius === "round" ? Math.min(20, Math.round(Math.min(w, h) * 0.28)) : Math.min(10, Math.round(Math.min(w, h) * 0.04));
      if (element.stroke) {
        rect.strokes = [paint("#C9C9C9")];
        rect.strokeWeight = 1;
      }
      canvas.appendChild(rect);
    }

    if (hasText) {
      var textInset = hasBox ? 6 : 0;
      var textWidth = Math.max(20, w - textInset * 2);
      var textSize = wireframeFontSize(element, scale, isFallbackLayout);
      if (hasBox) {
        var estimatedHeight = textEstimatedHeight(element.text || "", textSize, textWidth, element.text.split(/\n/).length);
        if (estimatedHeight > Math.max(12, h - 8) * 1.18) {
          textSize = wireframeFittedFontSize(element, scale, textWidth, Math.max(12, h - 8), isFallbackLayout);
        }
      }
      var text = createText(element.text, textSize, textSize >= 24, textColorForElement(element, fill), {
        width: textWidth,
        align: textAlignForElement(element),
        minLines: element.text.split(/\n/).length,
        weight: textSize >= 24 ? "bold" : "medium"
      });
      text.x = x + textInset;
      text.y = wireframeTextY(y, h, text.height, hasBox, element);
      canvas.appendChild(text);
    }
  });

  return canvas;
}

function createWireframePlaceholderNode(name, width, height) {
  var rect = createRectangle(
    name || "와이어 영역",
    Math.max(1, width),
    Math.max(1, height),
    wireframeOutputStyle("placeholderFill", "#D9D9D9"),
    wireframeOutputStyle("placeholderOpacity", 0.82)
  );
  rect.cornerRadius = Math.min(10, Math.round(Math.min(width, height) * 0.04));
  return rect;
}

function createWireframeImageNode(element, width, height) {
  var source = element && (element.imageDataUrl || element.inlineImageDataUrl || "");
  if (!isDataImageUrl(source)) return null;
  try {
    var image = figma.createImage(imageBytesFromDataUrl(source));
    var rect = createImageRectangle(image, Math.max(1, width), Math.max(1, height), 0, "원본 이미지 영역");
    rect.opacity = wireframeOutputStyle("imageOpacity", 0.42);
    return rect;
  } catch (error) {
    return null;
  }
}

function cleanWireframeRenderElements(elements, sourceWidth, sourceHeight) {
  var renderElements = safeArray(elements, []).filter(function (element) {
    return !isFullSlideBackgroundElement(element, sourceWidth, sourceHeight);
  });
  if (renderElements.length) return renderElements;
  return safeArray(elements, []);
}

function isFullSlideBackgroundElement(element, sourceWidth, sourceHeight) {
  if (!element || element.type === "svg" || element.type === "image" || element.text) return false;
  return element.x <= sourceWidth * 0.02 &&
    element.y <= sourceHeight * 0.02 &&
    element.w >= sourceWidth * 0.94 &&
    element.h >= sourceHeight * 0.94 &&
    (!element.fill || /^#?F{6}$/i.test(element.fill) || /^#?(FFFFFF|F5F5F5|F7F7F7|FAFAFA)$/i.test(element.fill));
}

function wireframeContentBounds(elements, sourceWidth, sourceHeight) {
  var minX = sourceWidth;
  var minY = sourceHeight;
  var maxX = 0;
  var maxY = 0;
  safeArray(elements, []).forEach(function (element) {
    if (!element) return;
    minX = Math.min(minX, element.x);
    minY = Math.min(minY, element.y);
    maxX = Math.max(maxX, element.x + element.w);
    maxY = Math.max(maxY, element.y + element.h);
  });
  if (minX > maxX || minY > maxY) {
    return { x: 0, y: 0, w: sourceWidth, h: sourceHeight };
  }
  var padX = sourceWidth * 0.018;
  var padY = sourceHeight * 0.018;
  minX = Math.max(0, minX - padX);
  minY = Math.max(0, minY - padY);
  maxX = Math.min(sourceWidth, maxX + padX);
  maxY = Math.min(sourceHeight, maxY + padY);
  return {
    x: minX,
    y: minY,
    w: Math.max(1, maxX - minX),
    h: Math.max(1, maxY - minY)
  };
}

function createGrayscaleSvgMockup(svg, width, height) {
  if (!figma.createNodeFromSvg) return null;
  try {
    var node = figma.createNodeFromSvg(grayscaleSvgMarkup(svg));
    node.name = "원본 슬라이드 SVG 목업";
    if (typeof node.resize === "function") {
      node.resize(Math.max(1, width), Math.max(1, height));
    }
    return node;
  } catch (error) {
    return null;
  }
}

function createSlidePictogramMockup(element, slide, width, height, elementIndex) {
  var frame = figma.createFrame();
  frame.name = "슬라이드 이미지 픽토그램 목업";
  frame.resize(width, height);
  frame.fills = [paint("#F3F3F3")];
  frame.clipsContent = true;
  frame.cornerRadius = Math.min(14, Math.round(Math.min(width, height) * 0.045));

  var context = wireframePictogramContext(slide);
  var ratio = width / Math.max(1, height);
  if (/국가|회원|돌파|해외|국내|한국|북미|일본|대만|태국|스페인|만\s*\+|만명|70만/i.test(context)) {
    drawMetricPictogram(frame, width, height);
  } else if (/쿠폰|혜택|할인|특가|지원금|PICK|OFF|%|응모|럭키|드로우/i.test(context)) {
    drawBenefitPictogram(frame, width, height);
  } else if (/카드|상품|강의|클래스|패스|큐레이션|리스트|탭|카테고리/i.test(context)) {
    drawCardPictogram(frame, width, height);
  } else if (ratio > 2.1) {
    drawBannerPictogram(frame, width, height);
  } else if (ratio < 0.78) {
    drawMobilePictogram(frame, width, height);
  } else {
    drawEditorialPictogram(frame, width, height, elementIndex);
  }
  return frame;
}

function createWholeSlideWireframeMockup(slide, width, height) {
  var frame = figma.createFrame();
  frame.name = "슬라이드 구조 목업";
  frame.resize(width, height);
  frame.fills = [paint("#FFFFFF")];
  frame.clipsContent = true;

  var lines = safeArray(slide && slide.lines, []).filter(Boolean);
  if (!lines.length && slide && slide.text) lines = wireframeTextLines(slide.text);
  var context = wireframePictogramContext(slide);
  var isMetric = /국가|회원|돌파|해외|국내|한국|북미|일본|대만|태국|스페인|만\s*\+|만명|70만/i.test(context);
  var isBenefit = /쿠폰|혜택|할인|특가|지원금|PICK|OFF|%|응모|럭키|드로우|구매하기/i.test(context);
  var isGrid = /카드|상품|강의|클래스|패스|큐레이션|리스트|탭|카테고리/i.test(context);

  if (isMetric) {
    drawWholeMetricSlide(frame, lines, width, height);
  } else if (isBenefit) {
    drawWholeBenefitSlide(frame, lines, width, height);
  } else if (isGrid) {
    drawWholeGridSlide(frame, lines, width, height);
  } else {
    drawWholeEditorialSlide(frame, lines, width, height);
  }
  return frame;
}

function drawWholeMetricSlide(frame, lines, width, height) {
  var topH = height * 0.46;
  appendMockRect(frame, "상단 시각 영역", width * 0.12, height * 0.04, width * 0.76, topH, "#FAFAFA", 0);
  var metricVisual = createSlidePictogramMockup({}, { title: "국가 회원 돌파", text: lines.join("\n"), lines: lines }, width * 0.56, topH * 0.70, 0);
  metricVisual.x = width * 0.22;
  metricVisual.y = height * 0.12;
  frame.appendChild(metricVisual);

  appendMockRect(frame, "섹션 구분선", width * 0.12, height * 0.52, width * 0.76, 1, "#DADADA", 0);
  addWireframeTextLines(frame, lines.slice(0, 3), width * 0.22, height * 0.55, width * 0.56, 28, "CENTER", true);
  addWireframeTextLines(frame, lines.slice(3, 9), width * 0.18, height * 0.72, width * 0.64, 15, "CENTER", false);
}

function drawWholeBenefitSlide(frame, lines, width, height) {
  appendMockRect(frame, "혜택 히어로", width * 0.08, height * 0.08, width * 0.84, height * 0.34, "#111111", Math.min(16, height * 0.025));
  addWireframeTextLines(frame, lines.slice(0, 3), width * 0.18, height * 0.16, width * 0.64, 30, "CENTER", true, "#FFFFFF");
  appendMockRect(frame, "CTA", width * 0.32, height * 0.47, width * 0.36, height * 0.055, "#111111", 999);
  addWireframeTextLines(frame, lines.slice(3, 5), width * 0.24, height * 0.56, width * 0.52, 20, "CENTER", true);
  addWireframeCardColumns(frame, lines.slice(5), width * 0.09, height * 0.70, width * 0.82, height * 0.20, 4);
}

function drawWholeGridSlide(frame, lines, width, height) {
  addWireframeTextLines(frame, lines.slice(0, 2), width * 0.20, height * 0.08, width * 0.60, 28, "CENTER", true);
  addWireframeCardColumns(frame, lines.slice(2), width * 0.08, height * 0.23, width * 0.84, height * 0.56, 3);
  appendMockRect(frame, "하단 버튼", width * 0.34, height * 0.84, width * 0.32, height * 0.06, "#111111", 999);
}

function drawWholeEditorialSlide(frame, lines, width, height) {
  appendMockRect(frame, "좌측 비주얼", width * 0.10, height * 0.12, width * 0.36, height * 0.48, "#E4E4E4", Math.min(16, height * 0.025));
  var editorialVisual = createSlidePictogramMockup({}, { title: "에디토리얼", text: lines.join("\n"), lines: lines }, width * 0.36, height * 0.48, 0);
  editorialVisual.x = width * 0.10;
  editorialVisual.y = height * 0.12;
  frame.appendChild(editorialVisual);
  addWireframeTextLines(frame, lines.slice(0, 4), width * 0.54, height * 0.18, width * 0.34, 27, "LEFT", true);
  addWireframeTextLines(frame, lines.slice(4, 10), width * 0.54, height * 0.47, width * 0.34, 15, "LEFT", false);
  addWireframeCardColumns(frame, lines.slice(10), width * 0.10, height * 0.72, width * 0.80, height * 0.16, 3);
}

function addWireframeTextLines(parent, lines, x, y, width, size, align, isBold, color) {
  var values = safeArray(lines, []).filter(Boolean).slice(0, 8);
  if (!values.length) return;
  var text = createText(values.join("\n"), size, Boolean(isBold), color || "#111111", {
    width: Math.max(20, width),
    align: align || "LEFT",
    minLines: values.length,
    weight: isBold ? "bold" : "medium"
  });
  text.x = Math.round(x);
  text.y = Math.round(y);
  parent.appendChild(text);
}

function addWireframeCardColumns(parent, lines, x, y, width, height, columns) {
  var count = Math.max(2, Math.min(columns || 3, 4));
  var gap = Math.max(8, width * 0.03);
  var cardW = (width - gap * (count - 1)) / count;
  var values = safeArray(lines, []).filter(Boolean);
  for (var i = 0; i < count; i++) {
    var cx = x + i * (cardW + gap);
    appendMockRect(parent, "정보 카드 " + i, cx, y, cardW, height, "#F3F3F3", Math.min(12, height * 0.08));
    appendMockRect(parent, "카드 이미지 " + i, cx + cardW * 0.22, y + height * 0.15, cardW * 0.56, height * 0.24, "#DADADA", Math.min(8, height * 0.06));
    if (values[i]) {
      addWireframeTextLines(parent, [values[i]], cx + cardW * 0.10, y + height * 0.58, cardW * 0.80, 13, "LEFT", false);
    }
  }
}

function wireframePictogramContext(slide) {
  return [
    slide && slide.title,
    slide && slide.text,
    safeArray(slide && slide.lines, []).join(" ")
  ].join(" ");
}

function appendMockRect(parent, name, x, y, width, height, fill, radius, opacity) {
  var rect = createRectangle(name, Math.max(1, width), Math.max(1, height), fill || "#D9D9D9", opacity == null ? 1 : opacity);
  rect.x = Math.round(x);
  rect.y = Math.round(y);
  rect.cornerRadius = radius || 0;
  parent.appendChild(rect);
  return rect;
}

function appendMockEllipse(parent, name, x, y, width, height, fill, opacity) {
  var ellipse = figma.createEllipse();
  ellipse.name = name;
  ellipse.resize(Math.max(1, width), Math.max(1, height));
  ellipse.x = Math.round(x);
  ellipse.y = Math.round(y);
  ellipse.fills = [paint(fill || "#D9D9D9", opacity == null ? 1 : opacity)];
  parent.appendChild(ellipse);
  return ellipse;
}

function drawMetricPictogram(frame, width, height) {
  var pad = Math.max(8, Math.round(Math.min(width, height) * 0.07));
  appendMockRect(frame, "상단 강조 바", width * 0.24, pad, width * 0.52, Math.max(6, height * 0.08), "#111111", 1);
  appendMockEllipse(frame, "지도 실루엣 1", width * 0.12, height * 0.50, width * 0.30, height * 0.16, "#D0D0D0", 1);
  appendMockEllipse(frame, "지도 실루엣 2", width * 0.36, height * 0.43, width * 0.28, height * 0.14, "#D8D8D8", 1);
  appendMockEllipse(frame, "지도 실루엣 3", width * 0.58, height * 0.51, width * 0.31, height * 0.15, "#D0D0D0", 1);
  appendMockRect(frame, "지도 기준선", width * 0.18, height * 0.70, width * 0.64, Math.max(2, height * 0.02), "#E4E4E4", 999);
  var bubbles = [
    [0.47, 0.17, 0.24, "#8C8C8C"],
    [0.19, 0.33, 0.19, "#B6B6B6"],
    [0.66, 0.36, 0.17, "#BDBDBD"],
    [0.40, 0.56, 0.13, "#C5C5C5"],
    [0.56, 0.60, 0.12, "#C5C5C5"]
  ];
  bubbles.forEach(function (item, index) {
    var size = Math.min(width, height) * item[2];
    appendMockEllipse(frame, "회원 수 원형 배지 " + index, width * item[0] - size / 2, height * item[1], size, size, item[3], 1);
    appendMockRect(frame, "배지 텍스트 라인 " + index, width * item[0] - size * 0.22, height * item[1] + size * 0.38, size * 0.44, Math.max(2, size * 0.035), "#F5F5F5", 999);
  });
}

function drawBenefitPictogram(frame, width, height) {
  var cardW = width * 0.58;
  var cardH = height * 0.46;
  appendMockRect(frame, "혜택 카드", width * 0.21, height * 0.20, cardW, cardH, "#FFFFFF", Math.min(16, height * 0.07));
  appendMockRect(frame, "쿠폰 절취선", width * 0.27, height * 0.33, cardW * 0.46, Math.max(3, height * 0.035), "#111111", 999);
  appendMockEllipse(frame, "쿠폰 심볼", width * 0.58, height * 0.26, height * 0.18, height * 0.18, "#B8B8B8", 1);
  appendMockRect(frame, "CTA 바", width * 0.17, height * 0.74, width * 0.66, Math.max(6, height * 0.07), "#111111", 999);
  appendMockEllipse(frame, "좌측 펀칭", width * 0.17 - height * 0.035, height * 0.42, height * 0.07, height * 0.07, "#F3F3F3", 1);
  appendMockEllipse(frame, "우측 펀칭", width * 0.83 - height * 0.035, height * 0.42, height * 0.07, height * 0.07, "#F3F3F3", 1);
}

function drawCardPictogram(frame, width, height) {
  var pad = Math.max(8, Math.round(width * 0.06));
  var gap = Math.max(6, Math.round(width * 0.04));
  var cols = width > height ? 3 : 2;
  var cardW = (width - pad * 2 - gap * (cols - 1)) / cols;
  var cardH = Math.max(height * 0.32, (height - pad * 2 - gap) / 2);
  for (var i = 0; i < Math.min(6, cols * 2); i++) {
    var col = i % cols;
    var row = Math.floor(i / cols);
    var x = pad + col * (cardW + gap);
    var y = pad + row * (cardH + gap);
    if (y + cardH > height - pad) break;
    appendMockRect(frame, "콘텐츠 카드 " + i, x, y, cardW, cardH, "#FFFFFF", Math.min(12, cardH * 0.08));
    appendMockRect(frame, "카드 썸네일 " + i, x + cardW * 0.14, y + cardH * 0.14, cardW * 0.72, cardH * 0.34, "#DADADA", Math.min(8, cardH * 0.06));
    appendMockRect(frame, "카드 텍스트 1 " + i, x + cardW * 0.14, y + cardH * 0.63, cardW * 0.62, Math.max(2, cardH * 0.035), "#111111", 999);
    appendMockRect(frame, "카드 텍스트 2 " + i, x + cardW * 0.14, y + cardH * 0.75, cardW * 0.44, Math.max(2, cardH * 0.03), "#BDBDBD", 999);
  }
}

function drawBannerPictogram(frame, width, height) {
  appendMockRect(frame, "배너 배경", width * 0.05, height * 0.16, width * 0.90, height * 0.68, "#111111", Math.min(14, height * 0.07));
  appendMockRect(frame, "배너 헤드라인 1", width * 0.12, height * 0.30, width * 0.34, Math.max(5, height * 0.09), "#FFFFFF", 999);
  appendMockRect(frame, "배너 헤드라인 2", width * 0.12, height * 0.45, width * 0.48, Math.max(5, height * 0.09), "#FFFFFF", 999);
  appendMockRect(frame, "배너 CTA", width * 0.12, height * 0.68, width * 0.30, Math.max(4, height * 0.055), "#D8D8D8", 999);
  appendMockEllipse(frame, "배너 오브젝트", width * 0.68, height * 0.25, height * 0.36, height * 0.36, "#7A7A7A", 1);
}

function drawMobilePictogram(frame, width, height) {
  appendMockRect(frame, "모바일 화면", width * 0.19, height * 0.06, width * 0.62, height * 0.88, "#FFFFFF", Math.min(20, width * 0.07));
  appendMockRect(frame, "모바일 상단", width * 0.28, height * 0.14, width * 0.44, height * 0.16, "#111111", Math.min(10, width * 0.04));
  appendMockRect(frame, "모바일 버튼", width * 0.30, height * 0.35, width * 0.40, height * 0.08, "#D9D9D9", 999);
  appendMockRect(frame, "모바일 리스트 1", width * 0.28, height * 0.51, width * 0.44, height * 0.08, "#EFEFEF", Math.min(8, width * 0.04));
  appendMockRect(frame, "모바일 리스트 2", width * 0.28, height * 0.63, width * 0.44, height * 0.08, "#EFEFEF", Math.min(8, width * 0.04));
}

function drawEditorialPictogram(frame, width, height, elementIndex) {
  var flip = elementIndex % 2;
  appendMockRect(frame, "에디토리얼 이미지 블록", width * (flip ? 0.48 : 0.08), height * 0.16, width * 0.42, height * 0.48, "#DADADA", Math.min(12, height * 0.05));
  appendMockRect(frame, "에디토리얼 제목 1", width * (flip ? 0.08 : 0.55), height * 0.22, width * 0.30, Math.max(4, height * 0.055), "#111111", 999);
  appendMockRect(frame, "에디토리얼 제목 2", width * (flip ? 0.08 : 0.55), height * 0.34, width * 0.24, Math.max(4, height * 0.05), "#111111", 999);
  appendMockRect(frame, "에디토리얼 본문 1", width * (flip ? 0.08 : 0.55), height * 0.51, width * 0.34, Math.max(3, height * 0.025), "#BDBDBD", 999);
  appendMockRect(frame, "에디토리얼 본문 2", width * (flip ? 0.08 : 0.55), height * 0.59, width * 0.26, Math.max(3, height * 0.025), "#CFCFCF", 999);
  appendMockEllipse(frame, "에디토리얼 포인트", width * (flip ? 0.66 : 0.19), height * 0.28, height * 0.20, height * 0.20, "#B8B8B8", 1);
}

function wireframeElementFill(element) {
  if (!element.fill) return "";
  return grayscaleColorForHex(element.fill, false);
}

function textColorForElement(element, fill) {
  if (element.textColor) return grayscaleColorForHex(element.textColor, true);
  if (!fill) return "#111111";
  return readableTextColor(fill);
}

function textAlignForElement(element) {
  if (element.align) return element.align;
  if (element.radius === "round" || element.h < 700000) return "CENTER";
  if (element.w > 5000000 && element.h < 900000) return "CENTER";
  return "LEFT";
}

function grayscaleColorForHex(value, forText) {
  var rgb = hexToRgb(value);
  var brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
  if (forText) {
    if (brightness < 95) return "#111111";
    if (brightness > 218) return "#FFFFFF";
    return "#555555";
  }
  if (brightness < 42) return "#050505";
  if (brightness > 248) return "#FFFFFF";
  if (brightness < 90) return "#333333";
  if (brightness < 140) return "#777777";
  if (brightness < 190) return "#B8B8B8";
  if (brightness < 225) return "#D8D8D8";
  return "#F1F1F1";
}

function grayscaleSvgMarkup(svg) {
  return String(svg || "")
    .replace(/#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?\b/g, function (match) {
      return grayscaleColorForHex(match, false);
    })
    .replace(/rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/gi, function (_, r, g, b) {
      var brightness = (Number(r) * 299 + Number(g) * 587 + Number(b) * 114) / 1000;
      var value = Math.max(0, Math.min(255, Math.round(brightness)));
      return "rgb(" + value + "," + value + "," + value + ")";
    });
}

function wireframeTextY(y, boxHeight, textHeight, hasBox, element) {
  var inset = hasBox ? Math.max(4, Math.min(10, Math.round(boxHeight * 0.06))) : 0;
  if (textHeight >= boxHeight - inset * 2) return y + inset;
  var vertical = element && element.verticalAlign ? element.verticalAlign : "";
  if (!vertical && element && (element.type === "ellipse" || element.radius === "round" || textAlignForElement(element) === "CENTER" && element.h < 700000)) {
    vertical = "CENTER";
  }
  if (vertical === "BOTTOM") return y + boxHeight - textHeight - inset;
  if (vertical === "CENTER") return y + Math.max(inset, Math.round((boxHeight - textHeight) / 2));
  return y + inset;
}

function wireframeFontSize(element, scaleY, isFallbackLayout) {
  var pointSize = wireframeFontPointSize(element);
  if (isFallbackLayout) {
    return Math.max(10, Math.min(72, Math.round(pointSize)));
  }
  var size;
  if (element.fontSizeUnit === "pt") {
    size = Math.round(pointSize * scaleY);
  } else {
    size = Math.round(pointSize * 12700 * scaleY);
  }
  if (pointSize >= 10.5 && pointSize <= 11.5) size = 14;
  if (pointSize >= 14.5 && pointSize <= 15.5) size = 18;
  if (pointSize >= 24.5 && pointSize <= 25.5) size = 28;
  if (pointSize > 0 && pointSize <= 16) {
    size = Math.max(size, Math.round(pointSize + 3));
  }
  return Math.max(10, Math.min(72, size));
}

function wireframeFittedFontSize(element, scaleY, width, maxHeight, isFallbackLayout) {
  var size = wireframeFontSize(element, scaleY, isFallbackLayout);
  var minLines = String(element.text || "").split(/\n/).length;
  while (size > 10 && textEstimatedHeight(element.text || "", size, width, minLines) > maxHeight) {
    size -= 1;
  }
  return size;
}

function wireframeFontPointSize(element) {
  if (element.fontSizeUnit === "pt") return Number(element.fontSize) || 18;
  return (Number(element.fontSize) || 1800) / 100;
}

function createWireframeCanvas(slide, index) {
  return createWireframeCanvasFromElements(slide, index);
}

function drawWireframeChrome(canvas, rect, width, height, index) {
  rect("safe area", 34, 34, width - 68, height - 68, "#FFFFFF", 0, 0.6);
  rect("top guide", 34, 34, width - 68, 1, "#D8D8D8", 0, 1);
  rect("bottom guide", 34, height - 35, width - 68, 1, "#D8D8D8", 0, 1);
  rect("left guide", 34, 34, 1, height - 68, "#D8D8D8", 0, 1);
  rect("right guide", width - 35, 34, 1, height - 68, "#D8D8D8", 0, 1);
  for (var i = 1; i < 4; i++) {
    rect("wireframe column guide " + i, Math.round(width * i / 4), 34, 1, height - 68, "#E6E6E6", 0, index % 2 ? 0.75 : 0.5);
  }
}

function drawWireframeRows(canvas, rect, lineText, lines, x, y, width, height, columns) {
  if (!lines.length) {
    rect("empty copy block", x, y, width, 52, "#EFEFEF", 10, 1);
    return;
  }

  var visibleLines = lines.slice(0, Math.min(lines.length, columns > 1 ? 14 : 10));
  var colCount = Math.max(1, columns || 1);
  var gap = 18;
  var colWidth = Math.floor((width - gap * (colCount - 1)) / colCount);
  var fontSize = visibleLines.length > 10 ? 11 : visibleLines.length > 6 ? 12 : 14;
  var col = 0;
  var rowY = y;

  visibleLines.forEach(function (line, index) {
    var rowHeight = Math.max(34, Math.min(74, textEstimatedHeight(line, fontSize, colWidth - 24) + 14));
    if (rowY + rowHeight > y + height && col < colCount - 1) {
      col += 1;
      rowY = y;
    }
    var rowX = x + col * (colWidth + gap);
    rect("wire copy row " + index, rowX, rowY, colWidth, rowHeight, "#FFFFFF", 10, 1);
    rect("wire copy bullet " + index, rowX + 12, rowY + 14, 8, 8, "#111111", 999, 1);
    lineText(line, rowX + 28, rowY + 9, colWidth - 42, fontSize, "medium", "#111111", "LEFT");
    rowY += rowHeight + 8;
  });
}

function detectWireframeLayout(slide) {
  var lines = slide.lines || [];
  var text = lines.join(" ");
  if (/^메인\s*페이지$|^main\s*page$/i.test(String(lines[0] || ""))) return "marker";
  if (lines.length <= 2) return "hero";
  if (/이미지|레퍼런스|비주얼|디자인|컨셉|main\s*visual|reference/i.test(text)) return "visual";
  if (/프로모션명|진행기간|목적|주요\s*스킴|명분|개요|기간|혜택/i.test(text)) return "info";
  if (lines.length >= 9) return "dense";
  return "content";
}

function fitWireframeTitleSize(value, maxSize) {
  var text = String(value || "");
  var longest = text.split(/\n+/).reduce(function (max, line) {
    return Math.max(max, line.length);
  }, 1);
  if (longest > 42) return Math.max(20, Math.min(maxSize, 26));
  if (longest > 28) return Math.max(22, Math.min(maxSize, 30));
  return maxSize;
}

function createPromotionInfo(analysis) {
  var group = autoFrame("프로모션 정보", "VERTICAL", {
    width: 3015,
    gap: 24,
    fill: null,
    padding: 0
  });

  group.appendChild(createInfoRow("프로모션명", createText(analysis.promotionName, 16, false, "#000000", {
    width: 1800,
    weight: "medium"
  })));
  group.appendChild(createLine(3015));
  group.appendChild(createInfoRow("진행기간", createText(analysis.period, 16, false, "#000000", {
    width: 2200,
    weight: "medium"
  })));
  group.appendChild(createLine(3015));
  group.appendChild(createOverviewRow(analysis));
  group.appendChild(createLine(3015));
  group.appendChild(createInfoRow("프로모션\n인사이트", createText(analysis.insight, 16, false, "#000000", {
    width: 2600,
    weight: "medium"
  })));

  return group;
}

function createAccuracyRow(accuracy) {
  var row = autoFrame("정확도 체크", "HORIZONTAL", {
    gap: 50,
    fill: null,
    padding: 0
  });

  row.appendChild(createText("정확도\n체크", 16, true, "#000000", {
    width: 118,
    weight: "bold"
  }));

  var detail = autoFrame("정확도 체크 상세", "VERTICAL", {
    width: 2600,
    gap: 7,
    fill: null,
    padding: 0
  });

  detail.appendChild(createText("확정값: " + formatAccuracyList(accuracy.confirmedFromBrief), 16, false, "#444444", { width: 2500, weight: "medium" }));
  detail.appendChild(createText("추론값: " + formatAccuracyList(accuracy.inferred), 16, false, "#444444", { width: 2500, weight: "medium" }));
  detail.appendChild(createText("확인 필요: " + formatAccuracyList(accuracy.needsCheck), 16, false, "#444444", { width: 2500, weight: "medium" }));

  row.appendChild(detail);
  return row;
}

function formatAccuracyList(values) {
  return safeArray(values, ["-"]).join(", ");
}

function createInfoRow(label, valueNode, rowHeight) {
  var options = {
    gap: 50,
    fill: null,
    padding: 0
  };
  if (rowHeight) options.height = rowHeight;
  var row = autoFrame("정보 행 - " + label.replace(/\n/g, " "), "HORIZONTAL", options);
  row.counterAxisAlignItems = "CENTER";

  var labelNode = createText(label, 16, true, "#000000", {
    width: 118,
    weight: "bold"
  });
  row.appendChild(labelNode);
  row.appendChild(valueNode);
  return row;
}

function createOverviewRow(analysis) {
  var row = autoFrame("개요", "HORIZONTAL", {
    gap: 50,
    fill: null,
    padding: 0
  });

  row.appendChild(createText("개요", 16, true, "#000000", { width: 118, weight: "bold" }));

  var overview = autoFrame("개요 상세", "VERTICAL", {
    gap: 11,
    fill: null,
    padding: 0
  });

  overview.appendChild(createLabeledChipRow("목적", [analysis.purpose], 100));
  overview.appendChild(createLabeledChipRow("주요 스킴", analysis.schemes, 100));
  overview.appendChild(createLabeledChipRow("명분", analysis.grounds, 100));
  row.appendChild(overview);
  return row;
}

function createLabeledChipRow(label, values, labelWidth) {
  var row = autoFrame("칩 행 - " + label, "HORIZONTAL", {
    gap: 28,
    fill: null,
    padding: 0
  });
  row.counterAxisAlignItems = "CENTER";

  row.appendChild(createText(label, 16, true, "#000000", {
    width: labelWidth,
    weight: "bold"
  }));

  var chips = autoFrame("칩 리스트", "HORIZONTAL", {
    gap: 4,
    fill: null,
    padding: 0
  });

  values.slice(0, 7).forEach(function (value) {
    chips.appendChild(createTag(String(value)));
  });

  row.appendChild(chips);
  return row;
}

async function createConceptCard(concept, index) {
  var card = autoFrame("Concept " + (index + 1) + " - " + concept.title, "VERTICAL", {
    width: 971,
    gap: 20,
    padding: 40,
    fill: "#FFFFFF",
    stroke: "#CDCDCD",
    radius: 10
  });

  card.appendChild(createText("Concept " + (index + 1), 22, true, "#000000", {
    width: 891,
    height: textBoxHeight(22)
  }));
  card.appendChild(createConceptSummary(concept));
  card.appendChild(createColorSection(concept));
  card.appendChild(await createMainAssetSection(concept, index));
  card.appendChild(await createTypographySection(concept, index));
  return card;
}

function createConceptSummary(concept) {
  var box = autoFrame("컨셉 요약", "VERTICAL", {
    width: 891,
    gap: 9,
    padding: { top: 22, right: 22, bottom: 22, left: 22 },
    fill: "#F3F3F3",
    radius: 10
  });
  box.counterAxisAlignItems = "CENTER";

  var copy = visualConceptCopy(concept);
  box.appendChild(createText(copy.title, 18, false, "#000000", { width: 847, align: "CENTER", weight: "medium" }));
  box.appendChild(createText(copy.hook, 17, true, "#000000", { width: 847, align: "CENTER", weight: "bold" }));
  box.appendChild(createText(copy.detail, 15, false, "#333333", { width: 847, align: "CENTER", weight: "regular" }));
  box.appendChild(createMoodTagRow(concept));
  return box;
}

function visualConceptCopy(concept) {
  var title = String(concept.title || "");
  var koreanTitle = String(concept.koreanTitle || "비주얼 컨셉");
  var asset = concept.asset || {};
  var conceptName = title ? title + " (" + koreanTitle + ")" : koreanTitle;
  var hook = "대표 워딩을 중심으로 한 첫 화면 키비주얼";
  var detail = "타이포 위계, 오브젝트 밀도, 배경 스테이지를 분리해 랜딩 첫 화면의 인상을 설계합니다.";

  if (/Offer|Commerce|Benefit|Deal/i.test(title + " " + koreanTitle)) {
    hook = "혜택 구조가 한눈에 읽히는 커머스 히어로";
    detail = "상품 카드, 가격 라벨, CTA 레일을 분리해 구매 흐름이 보이는 비주얼 시스템입니다.";
  } else if (/Typographic|Typography|Poster|Campaign/i.test(title + " " + koreanTitle)) {
    hook = "대표 문구를 강한 타이포 리듬으로 압축한 포스터형 히어로";
    detail = "큰 글자, 얇은 룰, 여백 대비를 사용해 캠페인 인지도를 빠르게 만듭니다.";
  } else if (/Draw|Lucky|Pop|Coupon|쿠폰|드로우|응모/i.test(title + " " + koreanTitle + " " + concept.mainAsset)) {
    hook = "참여 오브젝트가 튀어나오는 액션형 이벤트 히어로";
    detail = "티켓, 리본, 카드 레이어가 전면으로 떠오르며 즉시 참여감을 만듭니다.";
  } else if (/Editorial|Curation|Grid|큐레이션|에디토리얼/i.test(title + " " + koreanTitle + " " + concept.mainAsset)) {
    hook = "편집숍처럼 정돈된 큐레이션형 메인 비주얼";
    detail = "카드 그리드와 라벨 체계를 사용해 프로모션을 감도 있게 정리합니다.";
  } else if (/Chrome|Premium|Signal|프리미엄|기념/i.test(title + " " + koreanTitle + " " + concept.mainAsset)) {
    hook = "상징 오브젝트와 조명감으로 만든 프리미엄 히어로";
    detail = "금속감, 유리감, 중심 구도를 이용해 이벤트의 규모감과 신뢰감을 강화합니다.";
  } else if (/Reference|Key Visual|Landing/i.test(title + " " + koreanTitle)) {
    hook = "레퍼런스 무드를 직접 반영한 키비주얼";
    detail = "대표 문구, 메인 오브젝트, 배경 무드를 한 화면에 정렬하는 방향입니다.";
  }

  if (asset.texture || asset.angle || asset.background) {
    detail = truncateText(detail + " 질감은 " + localizeAssetLabel(asset.texture, "texture") + ", 구도는 " + localizeAssetLabel(asset.angle, "angle") + " 기준입니다.", 110);
  }

  return {
    title: truncateText(conceptName, 54),
    hook: truncateText(hook, 64),
    detail: truncateText(detail, 118)
  };
}

function shortConceptStatement(concept) {
  return truncateText(concept.hook || concept.description || concept.strategicReason, 54);
}

function createMoodTagRow(concept) {
  var row = autoFrame("mood tags", "HORIZONTAL", {
    gap: 6,
    fill: null,
    padding: 0
  });
  row.primaryAxisAlignItems = "CENTER";
  conceptMoodTags(concept).forEach(function (tag) {
    row.appendChild(createMoodTag(tag));
  });
  return row;
}

function conceptMoodTags(concept) {
  var notes = safeArray(concept.designNotes, []);
  if (notes.length) return notes.slice(0, 4);
  if (/Chrome|Premium|Signal|Luxury/i.test(concept.title)) return ["Luxury", "Metal", "Glass", "Light"];
  if (/Benefit|Carnival|Sale|Event/i.test(concept.title)) return ["Event", "Coupon", "Burst", "Contrast"];
  return ["Warm", "Growth", "Paper", "Community"];
}

function createMoodTag(value) {
  var tag = autoFrame("mood - " + value, "HORIZONTAL", {
    height: 28,
    gap: 8,
    padding: { top: 5, right: 12, bottom: 5, left: 12 },
    fill: "#FFFFFF",
    radius: 4
  });
  tag.counterAxisAlignItems = "CENTER";
  tag.appendChild(createText(value, 14, false, "#000000", { weight: "medium" }));
  return tag;
}

function createColorSection(concept) {
  var box = autoFrame("컬러", "VERTICAL", {
    width: 891,
    height: 297,
    gap: 16,
    padding: 20,
    fill: "#F3F3F3",
    radius: 7
  });

  box.appendChild(createText("color", 22, false, "#000000", {
    width: 851,
    height: textBoxHeight(22),
    weight: "medium"
  }));

  var stack = autoFrame("컬러 상세", "VERTICAL", {
    width: 851,
    gap: 7,
    fill: null,
    padding: 0
  });

  var hexRow = autoFrame("hex code", "HORIZONTAL", {
    width: 851,
    height: 68,
    gap: 12,
    padding: 10,
    fill: "#FFFFFF",
    radius: 6
  });
  hexRow.primaryAxisAlignItems = "SPACE_BETWEEN";
  hexRow.counterAxisAlignItems = "MIN";
  hexRow.appendChild(createText("hex code", 16, false, "#000000", {
    width: 90,
    height: textBoxHeight(16),
    weight: "medium"
  }));

  var swatches = autoFrame("hex chips", "HORIZONTAL", {
    gap: 5,
    fill: null,
    padding: 0
  });
  concept.colors.forEach(function (color) {
    swatches.appendChild(createColorChip(color.hex));
  });
  hexRow.appendChild(swatches);

  var ratioRow = autoFrame("ratio", "HORIZONTAL", {
    width: 851,
    height: 139,
    gap: 12,
    padding: 10,
    fill: "#FFFFFF",
    radius: 6
  });
  ratioRow.primaryAxisAlignItems = "SPACE_BETWEEN";
  ratioRow.counterAxisAlignItems = "MIN";
  ratioRow.appendChild(createText("ratio", 16, false, "#000000", {
    width: 90,
    height: textBoxHeight(16),
    weight: "medium"
  }));
  ratioRow.appendChild(createRatioBar(concept.colors));

  stack.appendChild(hexRow);
  stack.appendChild(ratioRow);
  box.appendChild(stack);
  return box;
}

function createColorMoodStrip(colors) {
  var row = autoFrame("color mood strip", "HORIZONTAL", {
    width: 851,
    height: 120,
    gap: 0,
    fill: null,
    padding: 0,
    clip: true,
    radius: 6
  });

  colors.forEach(function (color, index) {
    var width = index === 0 ? 596 : index === 1 ? 213 : 42;
    var swatch = autoFrame("mood color " + color.hex, "VERTICAL", {
      width: width,
      height: 120,
      gap: 4,
      padding: 14,
      fill: color.hex
    });
    swatch.primaryAxisAlignItems = "MAX";
    swatch.appendChild(createText(color.hex.toUpperCase(), 18, true, readableTextColor(color.hex), {
      width: Math.max(1, width - 28),
      height: textBoxHeight(18)
    }));
    row.appendChild(swatch);
  });

  return row;
}

function createColorLabelRow(colors) {
  var row = autoFrame("color labels", "HORIZONTAL", {
    width: 851,
    height: 26,
    gap: 8,
    fill: null,
    padding: 0
  });

  colors.forEach(function (color, index) {
    var label = index === 0 ? "Base" : index === 1 ? "Mood" : "Accent";
    row.appendChild(createText(label + " " + color.ratio + "%", 16, true, "#000000", {
      width: index === 0 ? 548 : index === 1 ? 187 : 100,
      height: textBoxHeight(16),
      align: index === 2 ? "RIGHT" : "LEFT"
    }));
  });

  return row;
}

function createColorChip(hex) {
  var chip = autoFrame("컬러칩 " + hex, "HORIZONTAL", {
    height: 48,
    gap: 9,
    padding: { top: 13, right: 20, bottom: 13, left: 20 },
    fill: "#FFFFFF",
    stroke: "#C3C3C3",
    radius: 4
  });
  chip.counterAxisAlignItems = "CENTER";

  var swatch = figma.createEllipse();
  swatch.name = "swatch " + hex;
  swatch.resize(20, 20);
  swatch.fills = [paint(hex)];
  swatch.strokes = [paint("#D8D8D8")];
  swatch.strokeWeight = 1;

  chip.appendChild(swatch);
  chip.appendChild(createText(hex.toUpperCase(), 18, false, "#000000", {
    width: 82,
    height: textBoxHeight(18),
    weight: "medium"
  }));
  return chip;
}

function createRatioBar(colors) {
  var totalWidth = 434;
  var height = 119;
  var ratioTotal = colors.reduce(function (sum, color) {
    return sum + Number(color.ratio || 0);
  }, 0) || 100;
  var usedWidth = 0;

  var bar = autoFrame("컬러 비율", "HORIZONTAL", {
    width: totalWidth,
    height: height,
    gap: 0,
    fill: null,
    radius: 4,
    padding: 0,
    clip: true
  });

  colors.forEach(function (color, index) {
    var segmentWidth = index === colors.length - 1
      ? totalWidth - usedWidth
      : Math.round(totalWidth * Number(color.ratio || 0) / ratioTotal);
    usedWidth += segmentWidth;
    var segment = autoFrame("ratio " + color.ratio, "HORIZONTAL", {
      width: Math.max(1, segmentWidth),
      height: height,
      gap: 0,
      fill: color.hex,
      padding: 0
    });
    segment.primaryAxisAlignItems = "CENTER";
    segment.counterAxisAlignItems = "CENTER";
    segment.appendChild(createText(String(color.ratio), 15, false, readableTextColor(color.hex), { weight: "medium" }));
    bar.appendChild(segment);
  });

  return bar;
}

async function createMainAssetSection(concept, index) {
  var box = autoFrame("메인 에셋", "VERTICAL", {
    width: 891,
    gap: 12,
    padding: 20,
    fill: "#F3F3F3",
    radius: 8
  });

  var header = autoFrame("메인 에셋 헤더", "HORIZONTAL", {
    width: 851,
    gap: 24,
    fill: null,
    padding: 0
  });
  header.primaryAxisAlignItems = "SPACE_BETWEEN";
  header.appendChild(createText("main asset", 22, true, "#000000", {
    width: 180,
    weight: "bold"
  }));
  header.appendChild(createText(concept.mainAsset, 16, false, "#222222", {
    width: 620,
    align: "RIGHT",
    weight: "medium"
  }));
  box.appendChild(header);

  box.appendChild(createVisualPreview("main", 851, 340, concept, index));

  var minis = autoFrame("에셋 목업 카드", "HORIZONTAL", {
    width: 851,
    height: 204,
    gap: 10,
    fill: null,
    padding: 0
  });
  minis.appendChild(createAssetMiniCard("texture", concept.asset.texture, concept, index + 1));
  minis.appendChild(createAssetMiniCard("angle", concept.asset.angle, concept, index + 2));
  minis.appendChild(createAssetMiniCard("background", concept.asset.background, concept, index + 3));
  box.appendChild(minis);

  var refHeader = autoFrame("레퍼런스 헤더", "HORIZONTAL", {
    width: 851,
    height: 24,
    gap: 12,
    fill: null,
    padding: 0
  });
  refHeader.primaryAxisAlignItems = "SPACE_BETWEEN";
  refHeader.counterAxisAlignItems = "CENTER";
  refHeader.appendChild(createText("Pinterest references", 16, true, "#000000", {
    width: 220,
    height: textBoxHeight(16)
  }));
  refHeader.appendChild(createText("대표 이미지 · 질감 · 구도 · 배경 참고용", 14, false, "#666666", {
    width: 400,
    height: textBoxHeight(14),
    align: "RIGHT"
  }));
  box.appendChild(refHeader);
  box.appendChild(await createReferenceStrip(concept, 851, 170));

  return box;
}

function createAssetMiniCard(label, value, concept, index) {
  var card = autoFrame(label, "VERTICAL", {
    width: 277,
    height: 204,
    gap: 8,
    padding: 10,
    fill: "#FFFFFF",
    radius: 7
  });
  card.appendChild(createText(assetSectionLabel(label), 12, false, "#000000", {
    width: 257,
    height: textBoxHeight(12),
    weight: "medium"
  }));
  card.appendChild(createVisualPreview(label, 257, 126, concept, index));
  card.appendChild(createText(localizeAssetLabel(value, label), 15, false, "#222222", {
    width: 257,
    height: 42,
    weight: "regular"
  }));
  return card;
}

function assetSectionLabel(label) {
  if (label === "texture") return "질감";
  if (label === "angle") return "구도";
  if (label === "background") return "배경";
  return label;
}

function localizeAssetLabel(value, label) {
  var text = String(value || "").trim();
  var source = text.toLowerCase();
  var maps = {
    texture: [
      [/matte|editorial poster|high contrast sticker/i, "무광 포스터 질감"],
      [/soft commerce|paper|gradient/i, "소프트 페이퍼 질감"],
      [/sharp digital|digital type/i, "선명한 디지털 질감"],
      [/glossy|ticket/i, "글로시 티켓 질감"],
      [/chrome|iridescent|metal|silver/i, "크롬 금속 질감"],
      [/sticker/i, "고대비 스티커 질감"]
    ],
    angle: [
      [/module|card|grid|상품/i, "모듈형 카드 배열"],
      [/isometric|3\/4|popup|pop/i, "아이소메트릭 팝업 구도"],
      [/typographic|poster|타이포/i, "대형 타이포 중심 구도"],
      [/low angle|로우/i, "중앙 집중 로우 앵글"],
      [/symmetry|대칭/i, "정면 대칭형 히어로 구도"],
      [/headline|정면/i, "정면형 헤드라인 구도"]
    ],
    background: [
      [/dark|black|studio|다크/i, "다크 스튜디오 배경"],
      [/lime|neon|라임|네온/i, "라임 포인트 스테이지"],
      [/offer|commerce|module|event/i, "커머스 오퍼 스테이지"],
      [/grid|그리드/i, "정돈된 그리드 배경"],
      [/paper|community|축하|밝은/i, "밝은 페이퍼 스테이지"],
      [/type|poster|타이포/i, "타이포 포스터 배경"]
    ]
  };
  var list = maps[label] || [];
  for (var i = 0; i < list.length; i++) {
    if (list[i][0].test(text) || list[i][0].test(source)) return list[i][1];
  }
  if (text && /[가-힣]/.test(text)) return text;
  if (label === "texture") return "비주얼 질감 샘플";
  if (label === "angle") return "메인 구도 샘플";
  if (label === "background") return "배경 스테이지 샘플";
  return text || "비주얼 샘플";
}

async function createReferenceStrip(concept, width, height) {
  var refs = [];
  refs.push(firstReference(concept.references && concept.references.mainVisual));
  refs.push(firstReference(concept.references && (concept.references.texture || concept.references.material)));
  refs.push(firstReference(concept.references && (concept.references.angle || concept.references.composition)));
  refs.push(firstReference(concept.references && (concept.references.background || concept.references.landingLayout || concept.references.layout)));
  refs = refs.filter(Boolean);

  if (!refs.length) {
    return createReferenceConnectionState(width, height, concept);
  }

  var row = autoFrame("pinterest reference strip", "HORIZONTAL", {
    width: width,
    height: height,
    gap: 9,
    fill: null,
    padding: 0
  });

  var labels = ["hero", "texture", "angle", "background"];
  for (var i = 0; i < refs.length && i < 4; i++) {
    row.appendChild(await createReferenceTile(refs[i], Math.floor((width - 27) / 4), height, labels[i]));
  }
  return row;
}

async function createMainReferencePreview(concept, width, height) {
  var refs = safeArray(concept.references && concept.references.mainVisual, []);
  if (refs[0]) {
    try {
      return await createReferenceHeroTile(refs[0], width, height);
    } catch (error) {
      return createReferenceLoadFailed(width, height);
    }
  }
  return createReferenceConnectionState(width, height, concept);
}

async function createReferenceMoodboard(concept, width, height) {
  var refs = safeArray(concept.references && concept.references.mainVisual, []).slice(0, 6);
  if (!refs.length) {
    return createReferenceConnectionState(width, height, concept);
  }

  var board = autoFrame("main asset moodboard", "VERTICAL", {
    width: width,
    height: height,
    gap: 8,
    fill: null,
    padding: 0
  });

  board.appendChild(await createReferenceHeroTile(refs[0], width, 324));

  var grid = autoFrame("supporting reference grid", "HORIZONTAL", {
    width: width,
    height: 264,
    gap: 8,
    fill: null,
    padding: 0
  });

  for (var i = 1; i < 5; i++) {
    if (refs[i]) {
      grid.appendChild(await createReferenceTile(refs[i], 206, 264, referenceRoleLabel(i)));
    }
  }

  board.appendChild(grid);
  return board;
}

async function createReferenceHeroTile(ref, width, height) {
  var tile = autoFrame("hero visual reference", "VERTICAL", {
    width: width,
    height: height,
    gap: 8,
    padding: 0,
    fill: "#FFFFFF",
    radius: 7,
    clip: true
  });

  var imageHeight = Math.max(1, height - 38);
  try {
    tile.appendChild(await createReferenceImageFromUrl(ref, width, imageHeight));
  } catch (error) {
    tile.appendChild(createReferenceLoadFailed(width, imageHeight));
  }

  tile.appendChild(createReferenceSourceLine(ref, "hero visual", width, 30));
  return tile;
}

function createMoodboardCaption(concept) {
  var row = autoFrame("moodboard keywords", "HORIZONTAL", {
    width: 851,
    height: 44,
    gap: 8,
    fill: null,
    padding: 0
  });
  row.primaryAxisAlignItems = "SPACE_BETWEEN";
  row.counterAxisAlignItems = "CENTER";
  var keywords = [
    concept.asset.texture,
    concept.asset.angle,
    concept.asset.background
  ].concat(safeArray(concept.referenceKeywords && concept.referenceKeywords.mainVisual, []).slice(0, 2));

  keywords.slice(0, 5).forEach(function (keyword) {
    row.appendChild(createMoodTag(truncateText(keyword, 24)));
  });
  return row;
}

async function createImagePreviewFromUrl(url, width, height) {
  var image = await loadFigmaImageFromSource(url);
  return createImageRectangle(image, width, height, 8, "image preview");
}

async function createReferenceImageFromUrl(ref, width, height) {
  var urls = [ref.inlineImageDataUrl, ref.imageDataUrl, ref.imageUrl].concat(safeArray(ref.fallbackImageUrls, []));
  var lastError = null;

  for (var i = 0; i < urls.length; i++) {
    var url = urls[i];
    if (!url) continue;
    try {
      var image = await loadFigmaImageFromSource(url);
      return createImageRectangle(image, width, height, 6, "reference - " + (ref.source || ref.title || "external"));
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("reference image load failed");
}

async function loadFigmaImageFromSource(source) {
  var url = String(source || "");
  if (!url) throw new Error("reference image source missing");

  if (isDataImageUrl(url)) {
    return figma.createImage(imageBytesFromDataUrl(url));
  }

  if (/^https?:\/\//i.test(url) && typeof figma.createImageAsync === "function") {
    try {
      return await figma.createImageAsync(url);
    } catch (error) {
      // Fall back to fetch below for plugin hosts where createImageAsync is unavailable for the URL.
    }
  }

  var response = await fetch(url, {
    method: "GET",
    headers: {
      "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"
    }
  });
  if (!response.ok) {
    throw new Error("reference image load failed: " + response.status);
  }
  return figma.createImage(new Uint8Array(await response.arrayBuffer()));
}

function isDataImageUrl(value) {
  return /^data:image\/(?:png|jpe?g|webp);base64,/i.test(String(value || ""));
}

function imageBytesFromDataUrl(dataUrl) {
  var match = String(dataUrl || "").match(/^data:image\/(?:png|jpe?g|webp);base64,([A-Za-z0-9+/=]+)$/i);
  if (!match) throw new Error("invalid inline image data");
  if (typeof atob !== "function") throw new Error("inline image decoder unavailable");

  var binary = atob(match[1]);
  var bytes = new Uint8Array(binary.length);
  for (var i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function createImageRectangle(image, width, height, radius, name) {
  var rect = figma.createRectangle();
  rect.resize(width, height);
  rect.fills = [{
    type: "IMAGE",
    scaleMode: "FILL",
    imageHash: image.hash
  }];
  rect.cornerRadius = radius;
  rect.name = name;
  return rect;
}

async function createReferenceImageGrid(concept, width, height) {
  var refs = safeArray(concept.references && concept.references.mainVisual, []).slice(0, 4);
  if (!refs.length) {
    return createReferenceConnectionState(width, height, concept);
  }

  var grid = autoFrame("reference image grid", "VERTICAL", {
    width: width,
    height: height,
    gap: 6,
    padding: 0,
    fill: null
  });

  for (var rowIndex = 0; rowIndex < 2; rowIndex++) {
    var row = autoFrame("reference row " + (rowIndex + 1), "HORIZONTAL", {
      width: width,
      height: 149,
      gap: 7,
      padding: 0,
      fill: null
    });

    for (var columnIndex = 0; columnIndex < 2; columnIndex++) {
      var ref = refs[rowIndex * 2 + columnIndex];
      if (ref) {
        row.appendChild(await createReferenceTile(ref, 166, 149));
      }
    }
    grid.appendChild(row);
  }

  return grid;
}

async function createReferenceTile(ref, width, height, roleLabel) {
  var tile = autoFrame("reference tile - " + (ref.source || ref.title || "external"), "VERTICAL", {
    width: width,
    height: height,
    gap: 5,
    padding: 0,
    fill: "#FFFFFF",
    radius: 6,
    clip: true
  });

  try {
    var imageNode = await createReferenceImageFromUrl(ref, width, height - 42);
    applySourceLink(imageNode, ref.pageUrl);
    tile.appendChild(imageNode);
  } catch (error) {
    tile.appendChild(createReferenceLoadFailed(width, height - 42));
  }

  tile.appendChild(createReferenceSourceLine(ref, roleLabel || ref.keyword || "reference", width, 37));

  return tile;
}

function createReferenceSourceLine(ref, label, width, height) {
  var reason = String(ref.whyMatched || "").replace(/^.*기준:\s*/, "");
  var sourceText = [label, reason || ref.source || "Pinterest"].filter(Boolean).join(" · ");
  var text = createText(truncateText(sourceText, 44), 10, false, "#333333", {
    width: width - 12,
    height: height,
    align: "CENTER",
    verticalAlign: "CENTER",
    weight: "medium"
  });
  applySourceLink(text, ref.pageUrl);
  return text;
}

function applySourceLink(node, url) {
  if (!node || !url) return node;
  try {
    if ("hyperlink" in node) node.hyperlink = { type: "URL", value: url };
  } catch (error) {
    // Some node types only support text range links.
  }
  try {
    if (node.type === "TEXT" && typeof node.setRangeHyperlink === "function") {
      node.setRangeHyperlink(0, node.characters.length, { type: "URL", value: url });
    }
  } catch (error) {
    // Hyperlink support varies by Figma runtime version.
  }
  try {
    if (typeof node.setPluginData === "function") node.setPluginData("sourceUrl", url);
  } catch (error) {
    // Plugin data is optional metadata for inspection.
  }
  return node;
}

function referenceRoleLabel(index) {
  return ["material", "composition", "lighting", "object"][index - 1] || "reference";
}

function createReferenceConnectionState(width, height, concept) {
  var isPinterest = concept && /pinterest/i.test(String(concept.referenceSearchStatus || concept.pinterestBoardSource || ""));
  var title = isPinterest ? "Pinterest board\nimage required" : "Reference API\nconnection required";
  var detail = isPinterest
    ? (concept.referenceSearchError || "보드 이미지를 불러오지 못했습니다.")
    : truncateText(safeArray(concept.referenceKeywords && concept.referenceKeywords.mainVisual, concept.asset.searchKeywords).join(" / "), 96);
  var frame = autoFrame("reference api connection required", "VERTICAL", {
    width: width,
    height: height,
    gap: 10,
    padding: 18,
    fill: "#FFFFFF",
    radius: 6
  });
  frame.primaryAxisAlignItems = "CENTER";
  frame.counterAxisAlignItems = "CENTER";

  frame.appendChild(createText(title, 20, true, "#000000", {
    width: width - 36,
    height: textBoxHeight(20, 2),
    align: "CENTER"
  }));
  frame.appendChild(createText(truncateText(detail, 96), 12, false, "#666666", {
    width: width - 36,
    height: textBoxHeight(12, 3),
    align: "CENTER"
  }));
  return frame;
}

function createReferenceLoadFailed(width, height) {
  var frame = autoFrame("reference image load failed", "VERTICAL", {
    width: width,
    height: height,
    gap: 6,
    padding: 10,
    fill: "#F3F3F3",
    radius: 6
  });
  frame.primaryAxisAlignItems = "CENTER";
  frame.counterAxisAlignItems = "CENTER";
  frame.appendChild(createText("image load failed", 12, true, "#666666", {
    width: width - 20,
    height: textBoxHeight(12),
    align: "CENTER"
  }));
  return frame;
}

function createImagePreviewFromBytes(rawBytes, width, height) {
  var bytes = rawBytes instanceof Uint8Array ? rawBytes : new Uint8Array(rawBytes);
  var image = figma.createImage(bytes);
  return createImageRectangle(image, width, height, 8, "image preview");
}

function referenceForAssetLabel(concept, label, fallbackIndex) {
  var refs = concept.references || {};
  if (label === "texture") return firstReference(refs.texture) || firstReference(refs.material) || firstReference(refs.mainVisual, fallbackIndex);
  if (label === "angle") return firstReference(refs.angle) || firstReference(refs.composition) || firstReference(refs.mainVisual, fallbackIndex);
  if (label === "background") return firstReference(refs.background) || firstReference(refs.layout) || firstReference(refs.landingLayout) || firstReference(refs.mainVisual, fallbackIndex);
  return firstReference(refs.mainVisual, fallbackIndex);
}

function firstReference(values, index) {
  var list = safeArray(values, []);
  return list[index || 0] || null;
}

function miniCardKeywordText(label, concept) {
  var keywords = concept.referenceKeywords || {};
  if (label === "texture") return safeArray(keywords.material, []).slice(0, 3).map(prefixHash).join("\n");
  if (label === "angle") return safeArray(keywords.composition, []).slice(0, 3).map(prefixHash).join("\n");
  if (label === "background") return safeArray(keywords.landingLayout, []).slice(0, 3).map(prefixHash).join("\n");
  return "";
}

function prefixHash(value) {
  return "#" + value;
}

async function createTypographySection(concept, index) {
  var box = autoFrame("타이포그래피", "VERTICAL", {
    width: 891,
    height: 322,
    gap: 7,
    padding: 20,
    fill: "#F3F3F3",
    radius: 7
  });

  var header = autoFrame("타이포그래피 헤더", "HORIZONTAL", {
    width: 851,
    height: 33,
    gap: 24,
    fill: null,
    padding: 0
  });
  header.primaryAxisAlignItems = "SPACE_BETWEEN";
  header.appendChild(createText("TYPOGRAPHY", 22, true, "#000000", {
    width: 180,
    height: textBoxHeight(22)
  }));
  header.appendChild(createText(concept.typography.style, 22, false, "#000000", {
    width: 460,
    height: textBoxHeight(22),
    align: "RIGHT"
  }));
  box.appendChild(header);

  var body = autoFrame("타이포그래피 상세", "HORIZONTAL", {
    width: 851,
    height: 245,
    gap: 25,
    fill: null,
    padding: 0
  });

  body.appendChild(await createTypePreview(405, 245, concept, index));

  var side = autoFrame("타이포그래피 속성", "VERTICAL", {
    width: 421,
    height: 245,
    gap: 7,
    fill: null,
    padding: 0
  });
  side.appendChild(createPairRow("특징", safeArray(concept.typography.features, []).slice(0, 3).map(prefixBullet), 421, 15, true, 77));
  side.appendChild(createPairRow("폰트", concept.typography.fontCandidates.map(formatFontCandidate), 421, 15, true, 77));
  side.appendChild(createPairRow("search keyword", concept.typography.searchKeywords.slice(0, 3).map(function (keyword) {
    return "#" + keyword;
  }), 421, 15, true, 77));

  body.appendChild(side);
  box.appendChild(body);
  return box;
}

function createFontRoleCard(label, font, width) {
  font = font || fallbackSecondaryFont();
  var card = autoFrame(label + " font", "HORIZONTAL", {
    width: width,
    height: 75,
    gap: 16,
    padding: 14,
    fill: "#FFFFFF",
    radius: 5
  });
  card.counterAxisAlignItems = "CENTER";
  card.appendChild(createText(label, 12, true, "#666666", {
    width: 80,
    height: textBoxHeight(12)
  }));
  card.appendChild(createText(font.name + "\n" + font.source, 20, true, "#000000", {
    width: width - 124,
    height: textBoxHeight(20, 2)
  }));
  return card;
}

function fallbackSecondaryFont() {
  return { name: "Spoqa Han Sans Neo", source: "Free", usage: "Body", reason: "" };
}

function prefixBullet(value) {
  return "- " + value;
}

function formatFontCandidate(font) {
  return "- " + font.name + " (" + font.source + ")";
}

function truncateText(value, maxLength) {
  var text = String(value || "");
  return text.length > maxLength ? text.slice(0, maxLength - 1) + "…" : text;
}

function stableHash(value) {
  var text = String(value || "");
  var hash = 2166136261;
  for (var i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return Math.abs(hash >>> 0);
}

function visualVariant(concept, index, count, salt) {
  var colors = safeArray(concept && concept.colors, []).map(function (color) {
    return color && color.hex;
  }).join(",");
  var copy = concept && concept.campaignCopy || {};
  var seed = [
    copy.headline,
    concept && concept.title,
    concept && concept.koreanTitle,
    concept && concept.mainAsset,
    colors,
    salt || ""
  ].join("|");
  return (stableHash(seed) + (index || 0) * 7) % count;
}

function createPairRow(label, values, width, fontSize, alignRight, fixedHeight) {
  var height = fixedHeight || Math.max(40, textBoxHeight(fontSize, values.length) + 20);
  var row = autoFrame("정보 박스 - " + label, "HORIZONTAL", {
    width: width,
    height: height,
    gap: 18,
    padding: 10,
    fill: "#FFFFFF",
    radius: 4
  });
  row.primaryAxisAlignItems = "SPACE_BETWEEN";
  row.counterAxisAlignItems = "MIN";
  row.appendChild(createText(label, 12, true, "#000000", {
    width: 110,
    height: textBoxHeight(12)
  }));
  row.appendChild(createText(values.join("\n"), fontSize, false, "#000000", {
    width: width - 150,
    height: height - 20,
    align: alignRight ? "RIGHT" : "LEFT"
  }));
  return row;
}

function createVisualPreview(label, width, height, concept, index) {
  var colors = concept.colors;
  var base = colors[0] ? colors[0].hex : "#111111";
  var ink = colors[1] ? colors[1].hex : "#111111";
  var point = colors[2] ? colors[2].hex : "#C7FF00";
  var copy = concept.campaignCopy || {};
  var headline = copy.headline || concept.visualNumber || concept.title || "프로모션";
  var benefit = copy.benefit || concept.visualNumber || "핵심 혜택";
  var action = copy.action || concept.koreanTitle || "참여 혜택";
  var support = copy.support || concept.purpose || concept.description || "";
  var isMain = label === "main";
  var preview = figma.createFrame();
  preview.name = "coded mockup - " + label;
  preview.resize(width, height);
  preview.fills = [paint(base)];
  preview.cornerRadius = isMain ? 12 : 8;
  preview.clipsContent = true;

  function rect(name, x, y, w, h, hex, radius, opacity) {
    var node = createRectangle(name, Math.max(1, w), Math.max(1, h), hex, opacity == null ? 1 : opacity);
    node.x = x;
    node.y = y;
    if (radius) node.cornerRadius = radius;
    preview.appendChild(node);
    return node;
  }

  function ellipse(name, x, y, w, h, hex, opacity) {
    var node = figma.createEllipse();
    node.name = name;
    node.resize(Math.max(1, w), Math.max(1, h));
    node.x = x;
    node.y = y;
    node.fills = [paint(hex, opacity == null ? 1 : opacity)];
    preview.appendChild(node);
    return node;
  }

  function rotate(node, degrees) {
    node.rotation = degrees;
    return node;
  }

  function heroTextNode(value, x, y, w, h, maxSize, hex, align) {
    var size = heroFontSize(value, w, h, maxSize);
    var node = createText(value, size, true, hex, {
      width: w,
      align: align || "LEFT",
      weight: "bold"
    });
    node.x = x;
    node.y = y + Math.max(0, Math.round((h - textEstimatedHeight(value, size, w)) / 2));
    preview.appendChild(node);
    return node;
  }

  function heroFontSize(value, targetWidth, targetHeight, maxSize) {
    var lines = cleanHeroText(value).split(/\n+/).filter(Boolean);
    var longest = lines.reduce(function (max, line) {
      return Math.max(max, line.length);
    }, 1);
    var byWidth = Math.floor(targetWidth / Math.max(0.62 * longest, 1));
    var byHeight = Math.floor((targetHeight || 120) / Math.max(lines.length * 1.2, 1));
    return Math.max(22, Math.min(maxSize || 76, byWidth, byHeight));
  }

  function drawEditorialGrid(lineColor, opacity) {
    for (var gx = 1; gx < 6; gx++) {
      rect("grid column " + gx, Math.round(width * gx / 6), 0, 1, height, lineColor, 0, opacity);
    }
    for (var gy = 1; gy < 4; gy++) {
      rect("grid row " + gy, 0, Math.round(height * gy / 4), width, 1, lineColor, 0, opacity);
    }
  }

  function drawTextureStudy(mode) {
    if (mode === 1) {
      rect("gloss base", 12, 12, width - 24, height - 24, point, 10, 0.96);
      for (var r = 0; r < 8; r++) {
        rotate(rect("burst strip " + r, Math.round(width * 0.48), Math.round(height * 0.16), Math.round(width * 0.08), Math.round(height * 0.62), r % 2 ? base : ink, 999, 0.9), -45 + r * 16);
      }
      for (var gi = 0; gi < 5; gi++) {
        rect("gloss perforation " + gi, Math.round(width * (0.18 + gi * 0.14)), Math.round(height * 0.35), 5, 5, "#FFFFFF", 999, 0.74);
        rect("gloss confetti " + gi, Math.round(width * (0.22 + gi * 0.12)), Math.round(height * (0.18 + (gi % 2) * 0.52)), Math.round(width * 0.04), 4, gi % 2 ? ink : "#FFFFFF", 999, 0.86);
      }
      rect("coupon layer back", Math.round(width * 0.18), Math.round(height * 0.38), Math.round(width * 0.56), Math.round(height * 0.34), "#FFFFFF", 12, 0.92);
      rotate(rect("coupon layer front", Math.round(width * 0.24), Math.round(height * 0.42), Math.round(width * 0.58), Math.round(height * 0.28), ink, 12, 0.95), -7);
      ellipse("specular light", Math.round(width * 0.62), Math.round(height * 0.22), Math.round(width * 0.22), Math.round(height * 0.22), "#FFFFFF", 0.48);
      return;
    }

    if (mode === 2) {
      rect("chrome base", 12, 12, width - 24, height - 24, darkBase ? "#090B17" : "#111111", 10, 1);
      ellipse("glass orb", Math.round(width * 0.24), Math.round(height * 0.06), Math.round(width * 0.52), Math.round(height * 0.78), base, 0.35);
      ellipse("glass highlight", Math.round(width * 0.34), Math.round(height * 0.12), Math.round(width * 0.28), Math.round(height * 0.36), "#FFFFFF", 0.28);
      ellipse("glass inner ring", Math.round(width * 0.3), Math.round(height * 0.18), Math.round(width * 0.4), Math.round(height * 0.52), "#FFFFFF", 0.12);
      for (var cg = 0; cg < 4; cg++) {
        rotate(rect("chrome light slit " + cg, Math.round(width * 0.18), Math.round(height * (0.23 + cg * 0.12)), Math.round(width * 0.58), 3, cg % 2 ? point : "#FFFFFF", 999, cg % 2 ? 0.62 : 0.42), -10 + cg * 5);
      }
      rotate(rect("chrome beam one", Math.round(width * 0.1), Math.round(height * 0.42), Math.round(width * 0.8), 14, "#FFFFFF", 999, 0.58), 12);
      rotate(rect("chrome beam two", Math.round(width * 0.08), Math.round(height * 0.57), Math.round(width * 0.76), 12, point, 999, 0.82), -8);
      rect("chrome floor", Math.round(width * 0.16), Math.round(height * 0.78), Math.round(width * 0.68), 9, ink, 999, 0.75);
      return;
    }

    if (mode === 3) {
      rect("paper base", 12, 12, width - 24, height - 24, "#FFFFFF", 10, 1);
      for (var pg = 0; pg < 6; pg++) {
        rect("paper grain line " + pg, 22, 22 + pg * Math.round(height * 0.11), width - 44, 1, ink, 0, 0.05);
      }
      for (var p = 0; p < 4; p++) {
        rotate(rect("paper card " + p, 28 + p * Math.round(width * 0.17), 26 + (p % 2) * 14, Math.round(width * 0.18), Math.round(height * 0.48), p % 2 ? point : quiet, 12, 0.95), p % 2 ? 6 : -5);
        rect("paper chip " + p, 38 + p * Math.round(width * 0.17), 42 + (p % 2) * 14, Math.round(width * 0.1), 5, ink, 999, 0.35);
      }
      rotate(rect("paper tear shadow", Math.round(width * 0.13), Math.round(height * 0.62), Math.round(width * 0.68), 6, point, 999, 0.35), -3);
      ellipse("paper soft shadow", Math.round(width * 0.18), Math.round(height * 0.68), Math.round(width * 0.62), Math.round(height * 0.18), ink, 0.1);
      return;
    }

    rect("texture base", 12, 12, width - 24, height - 24, darkBase ? "#151515" : "#FFFFFF", 10, 1);
    ellipse("material glow large", Math.round(width * 0.5), -24, Math.round(width * 0.52), Math.round(height * 0.86), base, 0.2);
    ellipse("material glow small", Math.round(width * 0.62), 16, Math.round(width * 0.28), Math.round(height * 0.58), point, 0.28);
    for (var mt = 0; mt < 4; mt++) {
      rect("material sample rail " + mt, 22, 24 + mt * Math.round(height * 0.16), width - 44, 1, ink, 0, 0.07);
    }
    for (var i = 0; i < 5; i++) {
      rotate(rect("material ribbon " + i, 30 + i * Math.round(width * 0.13), 30 + (i % 2) * 12, Math.round(width * 0.08), Math.round(height * 0.55), i % 2 ? ink : point, 999, 0.96), i % 2 ? 7 : -7);
      rect("material highlight " + i, 34 + i * Math.round(width * 0.13), 38 + (i % 2) * 12, Math.round(width * 0.055), 3, "#FFFFFF", 999, 0.55);
    }
    rect("material shadow", 24, height - 18, width - 48, 5, ink, 999, 0.22);
  }

  function drawCompositionStudy(mode) {
    if (mode === 1) {
      rect("composition full bleed", 12, 12, width - 24, height - 24, darkBase ? "#111111" : "#FFFFFF", 10, 1);
      rect("composition left rail", 28, 28, Math.round(width * 0.16), height - 56, base, 6, 1);
      rect("composition crop guide top", 28, 28, width - 56, 1, ink, 0, 0.15);
      rect("composition crop guide bottom", 28, height - 29, width - 56, 1, ink, 0, 0.15);
      for (var t = 0; t < 4; t++) {
        rect("composition type line " + t, Math.round(width * 0.27), 36 + t * 18, Math.round(width * (0.5 - t * 0.06)), 5, t % 2 ? ink : point, 0, 0.92);
      }
      rect("composition image block", Math.round(width * 0.55), Math.round(height * 0.34), Math.round(width * 0.3), Math.round(height * 0.34), quiet, 12, 1);
      ellipse("composition image object", Math.round(width * 0.63), Math.round(height * 0.4), Math.round(width * 0.18), Math.round(height * 0.18), point, 0.9);
      rotate(rect("composition direction axis", Math.round(width * 0.2), Math.round(height * 0.64), Math.round(width * 0.64), 4, point, 999, 0.92), -8);
      rect("composition bottom cta", Math.round(width * 0.26), Math.round(height * 0.78), Math.round(width * 0.58), 7, ink, 999, 1);
      return;
    }

    if (mode === 2) {
      rect("composition stage", 12, 12, width - 24, height - 24, quiet, 10, 1);
      ellipse("radial focal field", Math.round(width * 0.28), Math.round(height * 0.18), Math.round(width * 0.44), Math.round(height * 0.54), "#FFFFFF", 0.28);
      for (var c = 0; c < 5; c++) {
        rotate(rect("radial card " + c, Math.round(width * 0.42), Math.round(height * 0.12), Math.round(width * 0.16), Math.round(height * 0.5), c % 2 ? base : point, 12, 0.9), -36 + c * 18);
      }
      for (var rs = 0; rs < 6; rs++) {
        rotate(rect("radial small signal " + rs, Math.round(width * 0.48), Math.round(height * 0.18), Math.round(width * 0.035), 10, rs % 2 ? ink : "#FFFFFF", 999, 0.88), -70 + rs * 28);
      }
      rect("radial hero plate", Math.round(width * 0.22), Math.round(height * 0.46), Math.round(width * 0.56), Math.round(height * 0.28), "#FFFFFF", 12, 0.98);
      rect("radial cta", Math.round(width * 0.3), Math.round(height * 0.66), Math.round(width * 0.4), 7, ink, 999, 1);
      return;
    }

    if (mode === 3) {
      rect("grid composition board", 12, 12, width - 24, height - 24, "#FFFFFF", 10, 1);
      rect("grid hero block", 28, 28, Math.round(width * 0.38), Math.round(height * 0.58), darkBase ? "#111111" : ink, 12, 1);
      rect("grid side one", Math.round(width * 0.52), 28, Math.round(width * 0.34), Math.round(height * 0.26), point, 12, 1);
      rect("grid side two", Math.round(width * 0.52), Math.round(height * 0.48), Math.round(width * 0.34), Math.round(height * 0.2), quiet, 12, 1);
      for (var gl = 0; gl < 3; gl++) {
        rect("grid editorial line " + gl, Math.round(width * 0.12), 46 + gl * 18, Math.round(width * (0.26 - gl * 0.04)), 4, "#FFFFFF", 0, 0.7);
        rect("grid product line " + gl, Math.round(width * 0.56), Math.round(height * 0.55) + gl * 11, Math.round(width * (0.2 - gl * 0.03)), 3, ink, 0, 0.26);
      }
      rect("grid rule", 28, Math.round(height * 0.78), Math.round(width * 0.76), 8, base, 999, 1);
      return;
    }

    rect("composition board", 14, 12, width - 28, height - 24, paper, 10, 0.96);
    rect("composition hero", Math.round(width * 0.17), Math.round(height * 0.22), Math.round(width * 0.58), Math.round(height * 0.42), darkBase ? "#151515" : "#111111", 10, 1);
    rotate(rect("composition slash", Math.round(width * 0.18), Math.round(height * 0.46), Math.round(width * 0.62), 8, point, 999, 1), -9);
    rotate(rect("composition card one", Math.round(width * 0.58), Math.round(height * 0.12), Math.round(width * 0.22), Math.round(height * 0.28), base, 10, 0.9), 4);
    rotate(rect("composition card two", Math.round(width * 0.08), Math.round(height * 0.58), Math.round(width * 0.3), Math.round(height * 0.24), "#FFFFFF", 10, 0.96), -5);
    for (var cd = 0; cd < 3; cd++) {
      rect("composition deck line " + cd, Math.round(width * 0.2), Math.round(height * (0.28 + cd * 0.12)), Math.round(width * 0.42), 3, "#FFFFFF", 0, 0.28);
    }
    ellipse("composition focal ring", Math.round(width * 0.49), Math.round(height * 0.24), Math.round(width * 0.32), Math.round(height * 0.32), "#FFFFFF", 0.22);
    rect("composition bottom weight", Math.round(width * 0.16), Math.round(height * 0.76), Math.round(width * 0.68), 7, ink, 999, 1);
  }

  function drawBackgroundStudy(mode) {
    if (mode === 1) {
      rect("dark studio", 10, 10, width - 20, height - 20, "#101010", 10, 1);
      for (var dl = 0; dl < 5; dl++) {
        rect("studio depth line " + dl, 20 + dl * Math.round(width * 0.14), 20, 1, height - 40, "#FFFFFF", 0, 0.05);
      }
      ellipse("studio light left", -20, Math.round(height * 0.06), Math.round(width * 0.5), Math.round(height * 0.72), base, 0.18);
      ellipse("studio light right", Math.round(width * 0.58), Math.round(height * 0.08), Math.round(width * 0.42), Math.round(height * 0.58), point, 0.18);
      rect("studio plinth", Math.round(width * 0.18), Math.round(height * 0.62), Math.round(width * 0.64), Math.round(height * 0.18), "#FFFFFF", 12, 0.12);
      ellipse("studio object shadow", Math.round(width * 0.28), Math.round(height * 0.62), Math.round(width * 0.44), Math.round(height * 0.16), "#000000", 0.42);
      rect("studio floor line", Math.round(width * 0.12), Math.round(height * 0.78), Math.round(width * 0.76), 7, point, 999, 0.86);
      return;
    }

    if (mode === 2) {
      rect("commerce canvas", 10, 10, width - 20, height - 20, "#FFFFFF", 10, 1);
      for (var s = 0; s < 6; s++) {
        rect("commerce shelf " + s, 22, 24 + s * Math.round(height * 0.12), width - 44, 1, ink, 0, 0.08);
      }
      rect("commerce hero shelf", Math.round(width * 0.12), Math.round(height * 0.3), Math.round(width * 0.76), Math.round(height * 0.32), quiet, 14, 1);
      for (var m = 0; m < 3; m++) {
        rect("commerce item " + m, Math.round(width * (0.2 + m * 0.2)), Math.round(height * 0.4), Math.round(width * 0.12), Math.round(height * 0.14), m === 1 ? point : "#FFFFFF", 8, 0.9);
        ellipse("commerce item shine " + m, Math.round(width * (0.215 + m * 0.2)), Math.round(height * 0.405), Math.round(width * 0.05), Math.round(height * 0.05), "#FFFFFF", 0.55);
      }
      rect("commerce action rail", Math.round(width * 0.18), Math.round(height * 0.72), Math.round(width * 0.64), 7, base, 999, 1);
      return;
    }

    if (mode === 3) {
      rect("poster background", 10, 10, width - 20, height - 20, darkBase ? "#111111" : base, 10, 1);
      for (var pl = 0; pl < 4; pl++) {
        rect("poster baseline " + pl, 24, 26 + pl * Math.round(height * 0.15), width - 48, 1, "#FFFFFF", 0, darkBase ? 0.08 : 0.12);
      }
      for (var d = 0; d < 5; d++) {
        rotate(rect("poster layer " + d, Math.round(width * (0.14 + d * 0.12)), Math.round(height * (0.18 + d * 0.06)), Math.round(width * 0.36), Math.round(height * 0.12), d % 2 ? "#FFFFFF" : point, 999, d % 2 ? 0.82 : 0.95), -12 + d * 5);
      }
      ellipse("poster depth", Math.round(width * 0.54), Math.round(height * 0.18), Math.round(width * 0.38), Math.round(height * 0.5), ink, 0.18);
      rect("poster ground", Math.round(width * 0.14), Math.round(height * 0.78), Math.round(width * 0.72), 8, "#111111", 999, 0.92);
      return;
    }

    rect("background stage base", 10, 10, width - 20, height - 20, darkBase ? "#121212" : "#FFFFFF", 10, 1);
    for (var gx = 1; gx < 5; gx++) {
      rect("background column " + gx, Math.round(width * gx / 5), 10, 1, height - 20, ink, 0, 0.08);
    }
    for (var gy = 1; gy < 4; gy++) {
      rect("background row " + gy, 10, Math.round(height * gy / 4), width - 20, 1, ink, 0, 0.08);
    }
    ellipse("background spotlight one", Math.round(width * 0.06), Math.round(height * 0.06), Math.round(width * 0.42), Math.round(height * 0.62), base, 0.2);
    ellipse("background spotlight two", Math.round(width * 0.58), Math.round(height * 0.02), Math.round(width * 0.36), Math.round(height * 0.5), point, 0.18);
    rotate(rect("background light sweep", Math.round(width * 0.2), Math.round(height * 0.2), Math.round(width * 0.55), 8, "#FFFFFF", 999, 0.22), -11);
    rect("background hero zone", Math.round(width * 0.18), Math.round(height * 0.26), Math.round(width * 0.64), Math.round(height * 0.34), quiet, 12, 0.72);
    rect("background floor", Math.round(width * 0.12), Math.round(height * 0.72), Math.round(width * 0.76), 8, point, 999, 0.96);
    for (var b = 0; b < 4; b++) {
      rect("background module " + b, Math.round(width * (0.18 + b * 0.16)), Math.round(height * (0.34 + (b % 2) * 0.12)), Math.round(width * 0.09), Math.round(height * 0.14), "#FFFFFF", 7, 0.65);
    }
  }

  var title = String(concept.title || "");
  var darkBase = readableTextColor(base) === "#FFFFFF";
  var textColor = darkBase ? "#FFFFFF" : "#111111";
  var paletteHexes = colors.map(function (color) {
    return String(color && color.hex || "").toUpperCase();
  });
  var hasWarmPaper = paletteHexes.indexOf("#F6F0E4") >= 0 ||
    paletteHexes.indexOf("#F5EBDD") >= 0 ||
    paletteHexes.indexOf("#F8F3E8") >= 0;
  var quiet = darkBase ? "#2A2A2A" : (hasWarmPaper ? "#F6F0E4" : "#F4F4F4");
  var paper = darkBase ? (hasWarmPaper ? "#F6F0E4" : "#FFFFFF") : "#FFFFFF";

  drawEditorialGrid(textColor, darkBase ? 0.08 : 0.12);

  if (isMain) {
    var heroText = cleanHeroText(headline) || cleanHeroText(concept.title) || "PROMOTION";
    var heroColor = darkBase ? "#FFFFFF" : "#111111";
    var variant = visualVariant(concept, index, 6, "main");

    if (variant === 1) {
      rect("commerce hero field", 38, 34, width - 76, height - 68, quiet, 22, 1);
      rect("commerce stage crop", 70, 64, width - 140, height - 128, "#FFFFFF", 20, 1);
      rect("commerce top rule", 96, 86, width - 192, 2, ink, 0, 0.12);
      rect("commerce campaign chip", 102, 104, 96, 22, base, 999, 1);
      rect("commerce campaign chip shine", 116, 111, 56, 4, "#FFFFFF", 999, 0.68);
      for (var cp = 0; cp < 3; cp++) {
        var cx = 108 + cp * Math.round(width * 0.25);
        rect("commerce product card " + cp, cx, 90 + (cp === 1 ? -12 : 10), Math.round(width * 0.19), Math.round(height * 0.42), cp === 1 ? base : "#F7F7F7", 18, 1);
        ellipse("commerce object " + cp, cx + 34, 126 + (cp === 1 ? -10 : 16), 84, 84, cp === 1 ? "#FFFFFF" : ink, cp === 1 ? 0.66 : 0.88);
        rotate(rect("commerce ticket " + cp, cx + 18, 202 + (cp === 1 ? -12 : 8), Math.round(width * 0.14), 9, cp === 1 ? point : "#111111", 999, 1), cp === 1 ? -4 : 3);
        rect("commerce fine rule " + cp, cx + 22, 236 + (cp === 1 ? -12 : 8), Math.round(width * 0.11), 3, ink, 0, 0.18);
      }
      rotate(rect("commerce campaign band", 96, 222, width - 192, 18, "#111111", 999, 1), -1);
      ellipse("commerce glow", Math.round(width * 0.42), 78, 190, 190, point, 0.22);
      ellipse("commerce soft pool", Math.round(width * 0.18), 230, Math.round(width * 0.62), 44, ink, 0.08);
      heroTextNode(heroText, Math.round(width * 0.28), 106, Math.round(width * 0.44), 118, 58, heroColor, "CENTER");
      return preview;
    }

    if (variant === 2) {
      rect("poster field", 38, 34, width - 76, height - 68, "#111111", 22, 1);
      rect("poster left plate", 76, 66, Math.round(width * 0.13), height - 132, "#FFFFFF", 5, 1);
      rect("poster left accent", 94, 86, Math.round(width * 0.04), height - 172, base, 999, 1);
      for (var pr = 0; pr < 5; pr++) {
        rect("poster scanline " + pr, Math.round(width * 0.23), 82 + pr * 42, Math.round(width * (0.54 - pr * 0.04)), pr % 2 ? 3 : 5, pr % 2 ? "#FFFFFF" : point, 0, pr % 2 ? 0.85 : 1);
      }
      for (var ps = 0; ps < 4; ps++) {
        rect("poster micro paragraph " + ps, Math.round(width * 0.23), 246 + ps * 12, Math.round(width * (0.36 - ps * 0.045)), 3, "#FFFFFF", 0, 0.52);
      }
      ellipse("poster radial one", Math.round(width * 0.64), 56, 120, 120, base, 0.34);
      ellipse("poster radial two", Math.round(width * 0.72), 152, 160, 160, point, 0.16);
      rotate(rect("poster diagonal", Math.round(width * 0.24), height - 102, Math.round(width * 0.58), 8, base, 999, 1), -2);
      heroTextNode(heroText, Math.round(width * 0.22), 82, Math.round(width * 0.58), 150, 62, "#FFFFFF", "LEFT");
      return preview;
    }

    if (variant === 3) {
      rect("burst field", 38, 34, width - 76, height - 68, point, 22, 1);
      for (var br = 0; br < 10; br++) {
        rotate(rect("burst ray " + br, Math.round(width * 0.48), 48, Math.round(width * 0.08), Math.round(height * 0.66), br % 2 ? base : ink, 999, br % 2 ? 0.85 : 0.95), -58 + br * 13);
      }
      for (var bc = 0; bc < 14; bc++) {
        rotate(rect("burst confetti " + bc, Math.round(width * (0.12 + (bc % 7) * 0.11)), 72 + Math.floor(bc / 7) * 174, Math.round(width * 0.035), 5, bc % 3 === 0 ? "#FFFFFF" : (bc % 2 ? ink : base), 999, 0.88), -30 + bc * 9);
      }
      rect("burst hero slab", Math.round(width * 0.16), Math.round(height * 0.36), Math.round(width * 0.68), Math.round(height * 0.34), "#FFFFFF", 18, 0.96);
      rotate(rect("burst dark band", Math.round(width * 0.18), Math.round(height * 0.48), Math.round(width * 0.64), 28, "#111111", 999, 1), -4);
      ellipse("burst orb", Math.round(width * 0.62), 82, 142, 142, "#FFFFFF", 0.38);
      heroTextNode(heroText, Math.round(width * 0.24), 112, Math.round(width * 0.52), 126, 54, "#111111", "CENTER");
      return preview;
    }

    if (variant === 4) {
      rect("editorial field", 38, 34, width - 76, height - 68, "#FFFFFF", 22, 1);
      rect("editorial masthead", 68, 62, width - 136, 52, "#111111", 14, 1);
      rect("editorial mast accent", 92, 82, Math.round(width * 0.18), 8, base, 999, 1);
      for (var em = 0; em < 5; em++) {
        rect("editorial nav dot " + em, Math.round(width * (0.58 + em * 0.045)), 82, 20, 8, em % 2 ? "#FFFFFF" : point, 999, em % 2 ? 0.35 : 0.85);
      }
      rect("editorial color page", 68, 130, Math.round(width * 0.42), height - 210, base, 18, 1);
      rect("editorial product page", Math.round(width * 0.55), 130, Math.round(width * 0.32), height - 210, quiet, 18, 1);
      for (var er = 0; er < 5; er++) {
        rect("editorial rule " + er, Math.round(width * 0.58), 156 + er * 22, Math.round(width * (0.22 - er * 0.018)), 4, er % 2 ? ink : point, 0, 0.9);
      }
      ellipse("editorial object", Math.round(width * 0.67), 168, 110, 110, "#FFFFFF", 0.46);
      heroTextNode(heroText, 96, 146, Math.round(width * 0.34), 132, 48, readableTextColor(base), "LEFT");
      return preview;
    }

    if (variant === 5) {
      rect("premium field", 38, 34, width - 76, height - 68, "#090B18", 22, 1);
      for (var pg = 0; pg < 5; pg++) {
        rect("premium vertical grid " + pg, Math.round(width * (0.16 + pg * 0.14)), 54, 1, height - 108, "#FFFFFF", 0, 0.06);
      }
      ellipse("premium glass", Math.round(width * 0.18), 60, Math.round(width * 0.46), Math.round(height * 0.72), base, 0.34);
      ellipse("premium highlight", Math.round(width * 0.31), 86, Math.round(width * 0.22), Math.round(height * 0.28), "#FFFFFF", 0.24);
      rotate(rect("premium beam one", 96, 146, width - 200, 18, "#FFFFFF", 999, 0.52), 12);
      rotate(rect("premium beam two", 112, 210, width - 220, 15, point, 999, 0.86), -8);
      rect("premium stage", 92, height - 82, width - 184, 16, ink, 999, 0.8);
      heroTextNode(heroText, Math.round(width * 0.5), 94, Math.round(width * 0.34), 142, 46, "#FFFFFF", "LEFT");
      return preview;
    }

    rect("landing canvas", 38, 34, width - 76, height - 68, paper, 22, 0.98);
    rect("landing top nav", 72, 58, width - 144, 34, "#FFFFFF", 999, 0.82);
    rect("landing nav mark", 96, 72, 70, 5, base, 999, 1);
    rect("landing nav date", width - 210, 72, 96, 5, ink, 999, 0.34);
    rect("hero black panel", 78, 72, Math.round(width * 0.5), height - 144, "#111111", 18, 1);
    rotate(rect("key ribbon one", 95, 112, Math.round(width * 0.36), 30, base, 999, 1), 7);
    rotate(rect("key ribbon two", Math.round(width * 0.35), 170, Math.round(width * 0.44), 28, point, 999, 0.9), -7);
    rect("hero accent panel", Math.round(width * 0.62), 86, Math.round(width * 0.23), Math.round(height * 0.3), "#111111", 18, 1);
    rect("hero secondary panel", Math.round(width * 0.58), Math.round(height * 0.55), Math.round(width * 0.3), Math.round(height * 0.2), point, 18, 0.88);
    ellipse("key object one", Math.round(width * 0.68), 116, 96, 96, "#FFFFFF", 0.64);
    ellipse("key object two", Math.round(width * 0.74), 162, 120, 120, ink, 0.16);
    for (var lm = 0; lm < 4; lm++) {
      rect("landing module line " + lm, Math.round(width * 0.62), 222 + lm * 10, Math.round(width * (0.18 - lm * 0.025)), 3, "#111111", 0, 0.3);
    }
    rect("bottom line", 104, height - 82, Math.round(width * 0.72), 12, base, 999, 1);
    heroTextNode(heroText, 112, 98, Math.round(width * 0.36), 150, 52, "#FFFFFF", "LEFT");
    return preview;
  }

  if (label === "texture") {
    drawTextureStudy(visualVariant(concept, index, 4, "texture"));
    return preview;
  }

  if (label === "angle") {
    drawCompositionStudy(visualVariant(concept, index, 4, "angle"));
    return preview;
  }

  if (label === "background") {
    drawBackgroundStudy(visualVariant(concept, index, 4, "background"));
    return preview;
  }

  if (/Reference|Key Visual|Landing/i.test(title)) {
    rect("landing canvas", 28, 24, width - 56, height - 48, paper, 18, 0.98);
    rect("nav line", 58, 52, width - 116, 1, "#111111", 0, 0.18);
    rect("hero black panel", 58, 80, Math.round(width * 0.46), height - 132, "#111111", 16, 1);
    rect("hero accent panel", Math.round(width * 0.56), 82, Math.round(width * 0.31), Math.round(height * 0.42), base, 18, 1);
    rect("sub accent panel", Math.round(width * 0.56), Math.round(height * 0.58), Math.round(width * 0.31), Math.round(height * 0.22), "#FFFFFF", 18, 1);
    ellipse("product object one", Math.round(width * 0.61), 118, 88, 88, "#FFFFFF", 0.72);
    ellipse("product object two", Math.round(width * 0.72), 148, 120, 120, ink, 0.16);
    rect("thin cta", 84, height - 82, Math.round(width * 0.38), 10, base, 999, 1);
    return preview;
  }

  if (/Offer|Commerce|Module|Benefit/i.test(title)) {
    rect("commerce bg", 28, 24, width - 56, height - 48, quiet, 18, 1);
    for (var p = 0; p < 3; p++) {
      var px = 58 + p * Math.round(width * 0.27);
      rect("product card " + p, px, 102, Math.round(width * 0.23), Math.round(height * 0.5), "#FFFFFF", 16, 1);
      rect("product image " + p, px + 16, 120, Math.round(width * 0.23) - 32, Math.round(height * 0.27), p === 1 ? base : quiet, 12, 1);
      ellipse("product shine " + p, px + 38, 142, 54, 54, p === 1 ? "#FFFFFF" : point, 0.7);
      rect("price line " + p, px + 20, Math.round(height * 0.48), Math.round(width * 0.12), 5, "#111111", 999, 1);
      rect("caption line " + p, px + 20, Math.round(height * 0.54), Math.round(width * 0.16), 4, ink, 999, 0.22);
    }
    rect("coupon rail", 58, height - 76, width - 116, 38, "#111111", 999, 1);
    return preview;
  }

  if (/Typographic|Typography|Poster|Campaign/i.test(title)) {
    rect("poster black", 28, 28, width - 56, height - 56, "#111111", 16, 1);
    rect("vertical neon", 54, 54, Math.round(width * 0.13), height - 108, point, 4, 1);
    for (var t = 0; t < 5; t++) {
      rect("type rule " + t, Math.round(width * 0.22), Math.round(height * (0.72 + t * 0.04)), Math.round(width * (0.48 - t * 0.055)), 3, t % 2 ? "#FFFFFF" : point, 0, 1);
    }
    return preview;
  }

  if (/Curated|Curation|Board/i.test(title)) {
    rect("dark field", 28, 28, width - 56, height - 56, darkBase ? "#151515" : "#111111", 16, 1);
    for (var c = 0; c < 4; c++) {
      var cx = Math.round(width * (0.52 + (c % 2) * 0.2));
      var cy = Math.round(height * (0.2 + Math.floor(c / 2) * 0.28));
      rect("curation card " + c, cx, cy, Math.round(width * 0.16), Math.round(height * 0.2), c % 2 ? paper : point, 10, 1);
      rect("curation line " + c, cx + 14, cy + Math.round(height * 0.14), Math.round(width * 0.1), 3, "#111111", 0, 0.8);
    }
    rect("primary cta", 54, height - 66, Math.round(width * 0.42), 12, point, 999, 1);
    return preview;
  }

  rect("poster field", 28, 28, width - 56, height - 56, base, 16, 1);
  rect("black block", 48, 54, Math.round(width * 0.42), Math.round(height * 0.64), "#111111", 12, 1);
  rect("neon block", Math.round(width * 0.46), 54, Math.round(width * 0.38), Math.round(height * 0.3), point, 12, 1);
  rect("paper block", Math.round(width * 0.46), Math.round(height * 0.48), Math.round(width * 0.38), Math.round(height * 0.22), paper, 12, 1);
  rect("bottom cta", 68, height - 62, Math.round(width * 0.74), 10, ink, 999, 1);
  return preview;
}

async function createTypePreview(width, height, concept, index) {
  var colors = concept.colors;
  var background = colors[0] ? colors[0].hex : "#111111";
  var accent = colors[2] ? colors[2].hex : "#FFE734";
  var primaryFont = await loadConceptFont(concept.typography.fontCandidates[0], true);
  var secondaryFont = await loadConceptFont(concept.typography.fontCandidates[1], true);
  var preview = figma.createFrame();
  preview.name = "typography preview";
  preview.resize(width, height);
  preview.fills = [paint("#FFFFFF")];
  preview.cornerRadius = 6;
  preview.clipsContent = true;

  var field = createRectangle("type field", width, height, background, 1);
  preview.appendChild(field);

  var headline = createTextWithFont(concept.visualNumber, 96, primaryFont, accent, {
    width: width - 40,
    height: textBoxHeight(96),
    align: "CENTER"
  });
  headline.x = 20;
  headline.y = 38;
  preview.appendChild(headline);

  var sub = createTextWithFont(concept.title.toUpperCase(), 24, secondaryFont, readableTextColor(background), {
    width: width - 40,
    height: textBoxHeight(24),
    align: "CENTER"
  });
  sub.x = 20;
  sub.y = 145;
  preview.appendChild(sub);

  var bar = createRectangle("type accent", Math.round(width * 0.46), 8, colors[1] ? colors[1].hex : "#B2C6E7", 1);
  bar.cornerRadius = 4;
  bar.x = Math.round(width * 0.27);
  bar.y = 196 + index * 2;
  preview.appendChild(bar);
  return preview;
}

async function loadConceptFont(font, bold) {
  var family = font && font.name;
  var styles = family && availableFontsByFamily[family];
  if (!styles || !styles.length) return bold ? fontBold : fontRegular;
  var fontName = pickStyle(styles, bold ? ["Bold", "SemiBold", "Semibold", "Medium", "Regular"] : ["Regular", "Medium", "Book", "Normal"]) || styles[0];
  await figma.loadFontAsync(fontName);
  return fontName;
}

function createTag(value) {
  var width = chipWidthForText(value, 14);
  var tag = autoFrame("tag - " + value, "HORIZONTAL", {
    width: width,
    height: 28,
    gap: 10,
    padding: { top: 5, right: 10, bottom: 5, left: 10 },
    fill: "#EAEAEA",
    radius: 5
  });
  tag.counterAxisAlignItems = "CENTER";
  tag.appendChild(createText(value, 14, false, "#000000", {
    width: width - 20,
    height: textBoxHeight(14),
    weight: "medium",
    verticalAlign: "CENTER"
  }));
  return tag;
}

function chipWidthForText(value, size) {
  var text = sanitizeText(value).replace(/\s+/g, " ").trim();
  var korean = (text.match(/[가-힣]/g) || []).length;
  var other = Math.max(0, text.length - korean);
  return Math.min(360, Math.max(54, Math.ceil(korean * size * 0.98 + other * size * 0.56 + 26)));
}

function createLine(width) {
  return createRectangle("divider", width, 1, "#DCDCDC", 1);
}

function createRectangle(name, width, height, hex, opacity) {
  var rect = figma.createRectangle();
  rect.name = name;
  rect.resize(width, height);
  rect.fills = [paint(hex, opacity == null ? 1 : opacity)];
  return rect;
}

function autoFrame(name, direction, options) {
  var frame = figma.createFrame();
  frame.name = name;
  frame.fills = options.fill ? [paint(options.fill)] : [];
  frame.clipsContent = Boolean(options.clip);
  frame.layoutMode = direction;
  frame.itemSpacing = options.gap || 0;
  frame.primaryAxisAlignItems = "MIN";
  frame.counterAxisAlignItems = "MIN";

  var padding = normalizePadding(options.padding);
  frame.paddingTop = padding.top;
  frame.paddingRight = padding.right;
  frame.paddingBottom = padding.bottom;
  frame.paddingLeft = padding.left;

  if (options.radius) frame.cornerRadius = options.radius;
  if (options.stroke) {
    frame.strokes = [paint(options.stroke)];
    frame.strokeWeight = 1;
  }

  var width = options.width || null;
  var height = options.height || null;
  if (width || height) {
    frame.resize(width || 1, height || 1);
  }

  if (direction === "VERTICAL") {
    frame.primaryAxisSizingMode = height ? "FIXED" : "AUTO";
    frame.counterAxisSizingMode = width ? "FIXED" : "AUTO";
  } else {
    frame.primaryAxisSizingMode = width ? "FIXED" : "AUTO";
    frame.counterAxisSizingMode = height ? "FIXED" : "AUTO";
  }

  return frame;
}

function normalizePadding(padding) {
  if (typeof padding === "number") {
    return { top: padding, right: padding, bottom: padding, left: padding };
  }
  if (!padding) {
    return { top: 0, right: 0, bottom: 0, left: 0 };
  }
  return {
    top: padding.top || 0,
    right: padding.right || 0,
    bottom: padding.bottom || 0,
    left: padding.left || 0
  };
}

function createText(value, size, isBold, hex, options) {
  options = options || {};
  var text = figma.createText();
  text.name = "text";
  text.fontName = textFontForWeight(options.weight || (isBold ? "bold" : "regular"));
  text.fontSize = size;
  text.fills = [paint(hex || "#000000")];
  text.lineHeight = { unit: "PIXELS", value: lineHeightFor(size) };
  if (options.align) {
    text.textAlignHorizontal = options.align;
  }

  if (options.width && options.height) {
    text.textAutoResize = "NONE";
    text.resize(options.width, options.height);
    text.textAlignVertical = options.verticalAlign || "TOP";
  } else if (options.width) {
    text.textAutoResize = "HEIGHT";
    text.resize(options.width, textEstimatedHeight(value, size, options.width, options.minLines));
  } else {
    text.textAutoResize = "WIDTH_AND_HEIGHT";
  }

  text.characters = sanitizeText(value);

  if (options.width && options.height) {
    text.resize(options.width, options.height);
  } else if (options.width) {
    text.resize(options.width, textEstimatedHeight(value, size, options.width, options.minLines));
  }

  return text;
}

function createTextWithFont(value, size, fontName, hex, options) {
  options = options || {};
  var text = figma.createText();
  text.name = "text";
  text.fontName = fontName || fontRegular;
  text.fontSize = size;
  text.fills = [paint(hex || "#000000")];
  text.lineHeight = { unit: "PIXELS", value: lineHeightFor(size) };
  if (options.align) {
    text.textAlignHorizontal = options.align;
  }

  if (options.width && options.height) {
    text.textAutoResize = "NONE";
    text.resize(options.width, options.height);
    text.textAlignVertical = options.verticalAlign || "TOP";
  } else if (options.width) {
    text.textAutoResize = "HEIGHT";
    text.resize(options.width, textEstimatedHeight(value, size, options.width, options.minLines));
  } else {
    text.textAutoResize = "WIDTH_AND_HEIGHT";
  }

  text.characters = sanitizeText(value);

  if (options.width && options.height) {
    text.resize(options.width, options.height);
  } else if (options.width) {
    text.resize(options.width, textEstimatedHeight(value, size, options.width, options.minLines));
  }

  return text;
}

function textFontForWeight(weight) {
  if (weight === "bold") return fontBold;
  if (weight === "medium") return fontMedium || fontRegular;
  return fontRegular;
}

function textEstimatedHeight(value, size, width, minLines) {
  var text = sanitizeText(value);
  var manualLines = text.split("\n");
  var maxUnitsPerLine = Math.max(1, Math.floor((width || 120) / Math.max(1, size * 0.64)));
  var lineCount = manualLines.reduce(function (sum, line) {
    var clean = String(line || "");
    var units = textWidthUnits(clean);
    return sum + Math.max(1, Math.ceil(units / maxUnitsPerLine));
  }, 0);
  return textBoxHeight(size, Math.max(minLines || 1, lineCount));
}

function textWidthUnits(value) {
  return String(value || "").split("").reduce(function (sum, char) {
    return sum + (/[가-힣ㄱ-ㅎㅏ-ㅣ]/.test(char) ? 1.08 : 0.62);
  }, 0);
}

function textEstimatedWidth(value, size) {
  var text = sanitizeText(value);
  var longest = text.split("\n").reduce(function (max, line) {
    return Math.max(max, String(line || "").length);
  }, 1);
  return Math.max(8, Math.ceil(longest * size * 0.62));
}

function lineHeightFor(size) {
  return Math.round(size * 1.38);
}

function textBoxHeight(size, lines) {
  return lineHeightFor(size) * (lines || 1) + Math.ceil(size * 0.22);
}

function sanitizeText(value) {
  if (value == null) return "";
  return String(value).replace(/\u2028|\u2029/g, "\n");
}

function paint(hex, opacity) {
  var rgb = hexToRgb(hex);
  var solid = {
    type: "SOLID",
    color: {
      r: rgb.r / 255,
      g: rgb.g / 255,
      b: rgb.b / 255
    }
  };
  if (opacity != null) solid.opacity = opacity;
  return solid;
}

function hexToRgb(hex) {
  var normalized = String(hex || "#000000").replace("#", "").trim();
  if (normalized.length === 3) {
    normalized = normalized.split("").map(function (char) {
      return char + char;
    }).join("");
  }
  var value = parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255
  };
}

function readableTextColor(hex) {
  var rgb = hexToRgb(hex);
  var brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
  return brightness > 150 ? "#393939" : "#FFFFFF";
}

function normalizeAnalysis(analysis) {
  var fallback = sampleAnalysis();
  var merged = Object.assign({}, fallback, analysis || {});
  merged.period = formatPromotionPeriod(merged.period) || merged.period || fallback.period;
  merged.schemes = normalizeChipList(merged.schemes, fallback.schemes, 6);
  merged.grounds = normalizeChipList(merged.grounds, fallback.grounds, 5);
  merged.accuracy = Object.assign({}, fallback.accuracy, merged.accuracy || {});
  var campaignCopy = buildCampaignCopy(merged);
  merged.concepts = safeArray(merged.concepts, fallback.concepts).slice(0, 3).map(function (concept, index) {
    var fallbackConcept = fallback.concepts[index] || fallback.concepts[0];
    return Object.assign({}, fallbackConcept, concept || {}, {
      campaignCopy: Object.assign({}, campaignCopy, concept && concept.campaignCopy),
      colors: safeArray(concept && concept.colors, fallbackConcept.colors).slice(0, 3),
      asset: Object.assign({}, fallbackConcept.asset, concept && concept.asset),
      typography: normalizeTypography(Object.assign({}, fallbackConcept.typography, concept && concept.typography)),
      referenceKeywords: normalizeReferenceKeywords((concept && concept.referenceKeywords) || fallbackConcept.referenceKeywords),
      references: normalizeReferences((concept && concept.references) || fallbackConcept.references),
      designNotes: safeArray(concept && concept.designNotes, fallbackConcept.designNotes)
    });
  });
  return merged;
}

function normalizeChipList(values, fallback, limit) {
  var seen = {};
  var source = [];
  safeArray(values, fallback).forEach(function (value) {
    String(value || "")
      .replace(/(?:^|\s)(?:\d+|[①②③④⑤⑥⑦⑧⑨⑩])[\).]\s*/g, "\n")
      .split(/\n|,|·|•|\||\/|ㆍ|;|，|、/g)
      .forEach(function (item) {
        source.push(item);
      });
  });
  return source.map(cleanOutputChip).filter(function (value) {
    if (!value || value.length < 2 || value.length > 28) return false;
    if (/^\d{1,2}[./-]\d{1,2}$/.test(value) || /^\d+$/.test(value)) return false;
    if (/https?|pin\.it|www\.|[A-Za-z0-9]{8,}/i.test(value)) return false;
    if (/^slide\s*\d+$/i.test(value) || /^메인\s*페이지$/i.test(value)) return false;
    if (/^[A-Za-z0-9_.-]{3,}$/.test(value)) return false;
    var key = value.replace(/\s/g, "").toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).slice(0, limit || 6);
}

function cleanOutputChip(value) {
  return String(value || "")
    .replace(/^(?:\d+|[①②③④⑤⑥⑦⑧⑨⑩])[\).]\s*/g, "")
    .replace(/^[ㄴ└┗↳]+\s*/g, "")
    .replace(/^(?:프로모션명|프로모션 명|행사명|캠페인명|이벤트명|진행기간|기간|일정|개요|목적|목표|주요 스킴|스킴|명분|배경|근거|인사이트|디자인|컨셉|혜택)\s*[:：\-–—]?\s*/i, "")
    .replace(/^[ㄴ└┗↳]+\s*/g, "")
    .replace(/[()[\]{}“”"']/g, " ")
    .replace(/(?:유입용|미기획|예정|참고|진행예정|진행)/gi, " ")
    .replace(/(?:^|\s)(?:월|화|수|목|금|토|일)(?:\s|$)/g, " ")
    .replace(/\d+\s*일\s*\d+\s*명/g, " ")
    .replace(/\s*(?:활용한|활용하는|활용)\s*/g, " ")
    .replace(/\s*이상\s*할인율(?:\s*예정)?/g, " 할인")
    .replace(/100\s*원\s*가격제/g, "100원 딜")
    .replace(/https?:\/\/\S+|pin\.it\S*|www\.\S+/gi, " ")
    .replace(/변동/g, " ")
    .replace(/[<>•·ㆍ]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/^[\s:：\-–—.,]+|[\s:：\-–—.,]+$/g, "")
    .trim();
}

function buildCampaignCopy(analysis) {
  var referenceCopy = buildReferenceSlideCopy(analysis.mainReferenceBrief, analysis.mainReferenceHeroText);
  if (referenceCopy) return referenceCopy;

  var headline = analysis.promotionName || "프로모션";
  return {
    headline: cleanHeroText(headline),
    benefit: "",
    action: "",
    support: "",
    period: analysis.period || "",
    heroOnly: true
  };
}

function buildReferenceSlideCopy(referenceBrief, heroText) {
  var lines = referenceBriefLines(referenceBrief);
  var headline = cleanHeroText(heroText) || extractHeroTextFromReferenceBrief(referenceBrief) || cleanHeroText(lines[0]);
  if (!headline) return null;
  return {
    headline: headline,
    benefit: "",
    action: "",
    support: "",
    period: "",
    heroOnly: true
  };
}

function extractHeroTextFromReferenceBrief(referenceBrief) {
  var lines = referenceBriefLines(referenceBrief);
  if (!lines.length) return "";

  var phraseCandidates = [];
  lines.forEach(function (line) {
    var phrase = extractLatinCampaignPhrase(line);
    if (phrase) phraseCandidates.push(phrase);
  });
  if (phraseCandidates.length) {
    phraseCandidates.sort(function (a, b) {
      return b.length - a.length;
    });
    return cleanHeroText(phraseCandidates[0]);
  }

  var scored = lines.map(function (line, index) {
    var value = cleanHeroText(line);
    var score = 0;
    if (isNonHeroText(value)) score -= 200;
    if (/^[A-Z0-9\s&.'-]+$/.test(value) && /[A-Z]{2,}/.test(value)) score += 80;
    if (/[A-Z]{2,}/.test(value)) score += 40;
    if (/[가-힣]{2,}/.test(value)) score += 20;
    if (value.length >= 4 && value.length <= 32) score += 20;
    if (/[\[\]{}]|쿠폰|강의|신규|시그니처|오늘의|PICK|EVENT|PROMO/i.test(value)) score -= 60;
    if (/^\d|%|https?:\/\//i.test(value)) score -= 40;
    score -= index;
    return { value: value, score: score };
  }).filter(function (item) {
    return item.value;
  }).sort(function (a, b) {
    return b.score - a.score;
  });

  return scored.length ? cleanHeroText(scored[0].value) : "";
}

function extractLatinCampaignPhrase(value) {
  var text = String(value || "").replace(/\s+/g, " ").trim();
  var matches = text.match(/[A-Z][A-Z0-9]*(?:\s+[A-Z][A-Z0-9]*){1,5}/g);
  if (!matches || !matches.length) return "";
  matches.sort(function (a, b) {
    return b.length - a.length;
  });
  return matches[0];
}

function cleanHeroText(value) {
  var text = normalizeFetchedText(value)
    .split(/\n+/)
    .map(function (line) {
      return line
        .replace(/^[\-•·ㆍ\*]+\s*/, "")
        .replace(/https?:\/\/\S+|pin\.it\S*|www\.\S+/gi, " ")
        .replace(/\s+/g, " ")
        .trim();
    })
        .filter(function (line) {
          if (!line || line.length < 2) return false;
          if (isNonHeroText(line)) return false;
          if (/^\d{1,2}[./-]\d{1,2}|^20\d{2}[./-]/.test(line)) return false;
      if (/^EVENT$|^PROMO$/i.test(line)) return false;
      return true;
    })
    .join("\n");

  var phrase = extractLatinCampaignPhrase(text);
  if (phrase && (!/[가-힣]/.test(text) || phrase.length >= 10)) {
    text = phrase;
  }

  text = text
    .replace(/\bCOLOSO\s+SUMMER\s+BLACK\s+FRIDAY\b/i, "COLOSO\nSUMMER\nBLACK FRIDAY")
    .replace(/\bSUMMER\s+BLACK\s+FRIDAY\b/i, "SUMMER\nBLACK FRIDAY");

  var seen = {};
  return text.split(/\n+/).map(function (line) {
    return line.trim();
  }).filter(function (line) {
    var key = line.replace(/\s/g, "").toLowerCase();
    if (!line || seen[key]) return false;
    seen[key] = true;
    return true;
  }).slice(0, 4).join("\n");
}

function referenceBriefLines(referenceBrief) {
  return String(referenceBrief || "")
    .split(/\n+/)
    .map(function (line) {
      return line.replace(/^[\-•·ㆍ\*]+\s*/, "").replace(/\s+/g, " ").trim();
    })
    .filter(function (line) {
      if (!line || line.length < 2) return false;
      if (isNonHeroText(line)) return false;
      if (/^slide\s+\d+$/i.test(line)) return false;
      if (/^(main\s*)?visual\s*reference|reference|ref\.?|메인\s*페이지|레퍼런스|참고|image|이미지$/i.test(line)) return false;
      if (/https?:\/\/|pin\.it|www\./i.test(line)) return false;
      return true;
    })
    .slice(0, 6);
}

function firstMatching(values, pattern) {
  for (var i = 0; i < values.length; i++) {
    if (pattern.test(String(values[i] || ""))) return values[i];
  }
  return "";
}

function formatPromotionPeriod(value) {
  var dates = parsePromotionDates(value);
  if (!dates.length) return "";
  var start = dates[0];
  var end = dates[1] || dates[0];
  var startTime = start.time || "10:30";
  var endTime = dates[1] && end.time ? end.time : "23:59";
  var totalDays = Math.max(1, Math.round((dateOnly(end) - dateOnly(start)) / 86400000) + 1);
  return formatDatePart(start) + " " + startTime + " - " + formatDatePart(end) + " " + endTime + " *총 " + totalDays + "일";
}

function parsePromotionDates(value) {
  var source = String(value || "");
  var fullPattern = /(20\d{2})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})(?:일)?(?:\s*\(([월화수목금토일])\))?(?:\s*(\d{1,2}):(\d{2}))?/g;
  var dates = [];
  var match;
  while ((match = fullPattern.exec(source))) {
    dates.push(toDatePart(match[1], match[2], match[3], match[4], match[5], match[6], match.index, match[0].length));
  }
  dedupeDateParts(dates);
  if (dates.length >= 2) return dates.slice(0, 2);
  if (!dates.length) return [];

  var tail = source.slice(dates[0].index + dates[0].length);
  var shortPattern = /(?:^|[^\d])(\d{1,2})[.\-/월]\s*(\d{1,2})(?:일)?(?:\s*\(([월화수목금토일])\))?(?:\s*(\d{1,2}):(\d{2}))?/g;
  while ((match = shortPattern.exec(tail))) {
    var shortDate = toDatePart(dates[0].year, match[1], match[2], match[3], match[4], match[5], match.index, match[0].length);
    if (shortDate.month >= 1 && shortDate.month <= 12 && shortDate.day >= 1 && shortDate.day <= 31) {
      dates.push(shortDate);
      break;
    }
  }
  if (dates.length === 1) {
    var partialMonth = tail.match(/20\d{2}[.\-/년]\s*(\d{1,2})(?:월)?(?!\s*[.\-/]\s*\d)/) || tail.match(/(?:^|[^\d])(\d{1,2})월(?!\s*\d)/);
    if (partialMonth) {
      var month = Number(partialMonth[1]);
      if (month >= 1 && month <= 12) {
        dates.push(toDatePart(dates[0].year, month, new Date(dates[0].year, month, 0).getDate(), "", "", "", 0, 0));
      }
    }
  }
  return dates.slice(0, 2);
}

function dedupeDateParts(dates) {
  var seen = {};
  for (var index = dates.length - 1; index >= 0; index--) {
    var date = dates[index];
    var key = date.year + "." + date.month + "." + date.day + "." + (date.time || "");
    if (seen[key]) {
      dates.splice(index, 1);
    } else {
      seen[key] = true;
    }
  }
}

function toDatePart(year, month, day, weekday, hour, minute, index, length) {
  return {
    year: Number(year),
    month: Number(month),
    day: Number(day),
    weekday: weekday || "",
    time: hour && minute ? pad2(hour) + ":" + pad2(minute) : "",
    index: index || 0,
    length: length || 0
  };
}

function dateOnly(part) {
  return new Date(part.year, part.month - 1, part.day);
}

function formatDatePart(part) {
  return part.year + "." + pad2(part.month) + "." + pad2(part.day) + " (" + (part.weekday || weekdayKo(part)) + ")";
}

function weekdayKo(part) {
  return ["일", "월", "화", "수", "목", "금", "토"][dateOnly(part).getDay()];
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function normalizeTypography(typography) {
  var source = typography || {};
  var legacyFonts = safeArray(source.fonts, []).map(function (name) {
    return {
      name: name,
      source: inferFontSource(name),
      usage: "Headline",
      reason: "컨셉 톤에 맞는 실사용 후보"
    };
  });
  return {
    style: source.style || "Commercial Sans System",
    fontCandidates: safeArray(source.fontCandidates, legacyFonts).filter(function (font) {
      return font && font.name && isAllowedFontName(font.name);
    }).slice(0, 4),
    features: safeArray(source.features, []),
    pairingGuide: source.pairingGuide || "헤드라인용 후보와 본문용 Spoqa Han Sans Neo 또는 Noto Sans KR을 조합해 정보 가독성을 확보합니다.",
    searchKeywords: safeArray(source.searchKeywords, [])
  };
}

function inferFontSource(name) {
  if (/Acumin|Futura PT|DIN Condensed|Source Han Sans|Adobe Clean|Cooper Black Std|Termina/i.test(name)) return "Adobe Fonts";
  if (/Sandoll|산돌|격동|네오고딕|삼립|공병각/i.test(name)) return "Sandoll";
  return "Free";
}

function isAllowedFontName(name) {
  return !/Druk|Compacta|Neue Haas Grotesk/i.test(String(name || ""));
}

function normalizeReferenceKeywords(keywords) {
  var source = keywords || {};
  return {
    mainVisual: safeArray(source.mainVisual, []),
    material: safeArray(source.material, []),
    composition: safeArray(source.composition, []),
    typography: safeArray(source.typography, []),
    landingLayout: safeArray(source.landingLayout, [])
  };
}

function normalizeReferences(references) {
  var source = references || {};
  return {
    mainVisual: normalizeReferenceList(source.mainVisual),
    texture: normalizeReferenceList(source.texture || source.material),
    angle: normalizeReferenceList(source.angle || source.composition),
    background: normalizeReferenceList(source.background || source.landingLayout || source.layout),
    material: normalizeReferenceList(source.material || source.texture),
    composition: normalizeReferenceList(source.composition || source.angle),
    landingLayout: normalizeReferenceList(source.landingLayout || source.background || source.layout),
    typography: normalizeReferenceList(source.typography),
    layout: normalizeReferenceList(source.layout)
  };
}

function normalizeReferenceList(values) {
  return safeArray(values, []).filter(function (item) {
    return item && (item.imageUrl || item.inlineImageDataUrl || item.imageDataUrl);
  }).map(function (item) {
    return {
      title: item.title || "",
      source: item.source || "",
      imageUrl: item.imageUrl || "",
      pinId: item.pinId || "",
      inlineImageDataUrl: item.inlineImageDataUrl || item.imageDataUrl || "",
      fallbackImageUrls: safeArray(item.fallbackImageUrls, []),
      pageUrl: item.pageUrl || "",
      keyword: item.keyword || "",
      whyMatched: item.whyMatched || ""
    };
  }).slice(0, 8);
}

function safeArray(value, fallback) {
  return Array.isArray(value) && value.length ? value : fallback;
}

function sampleAnalysis() {
  return {
    promotionName: "70만 회원 감사제",
    period: "2026.5.22(금) 10:30 - 2026.5.28(목) 23:59 (KST)",
    purpose: "신규 유입 극대화",
    schemes: ["기간한정 판매", "정액 쿠폰팩", "클래스컷", "웰컴 혜택"],
    grounds: ["유저 70만명 달성", "기념", "감사"],
    insight: "단순 할인을 넘어 사용자의 성장을 축하하고 보답하는 페스티벌 무드 형성",
    accuracy: {
      confirmedFromBrief: ["promotionName", "period", "purpose", "schemes", "grounds"],
      inferred: ["insight", "concepts", "tone", "visualDirection"],
      needsCheck: []
    },
    concepts: [
      {
        title: "Chrome Celebration",
        koreanTitle: "미래지향적 축제",
        hook: "'70만'이라는 수치를 빛나는 훈장처럼 시각화하여 프리미엄 가치 부여",
        description: "세련된 크롬 질감과 빛의 굴절을 활용한 미래적인 감사제 비주얼",
        visualNumber: "70",
        colors: [
          { hex: "#070A18", ratio: 68 },
          { hex: "#3247FF", ratio: 22 },
          { hex: "#C7FF00", ratio: 10 }
        ],
        mainAsset: "굴절률이 높은 유리 구체 안에 부유하는 디지털 숫자와 크롬 리본",
        strategicReason: "회원 수 달성이라는 명분을 고급스러운 상징물로 격상시켜 감사제의 규모감과 프리미엄 보상감을 동시에 전달합니다.",
        targetEmotion: "큰 이벤트에 초대받았다는 기대감과 특별 대우를 받는 감정",
        visualDirection: "다크 스튜디오 위에 크롬, 유리, 얇은 빛의 링을 배치하고 숫자를 중심 오브젝트로 세우는 프리미엄 3D 히어로",
        imagePrompt: "Premium futuristic celebration key visual for a 70만 회원 감사제 promotion, iridescent chrome ribbon, glass sphere, floating 70 milestone number, dark minimal studio background, cinematic rim light, high-end 3D render, luxury tech mood, central hero composition, sharp reflections, landing page hero image, No readable Korean text except simple number or symbol, no long copy inside image",
        negativePrompt: "low quality, cartoonish, messy confetti, cheap sale banner, readable Korean copy, long text, distorted number, noisy background, stock photo people",
        mobileCropGuide: "숫자와 유리 구체를 중앙 60% 안에 유지하고, 크롬 리본은 상하로 잘려도 형태가 이어지게 배치합니다.",
        landingUsageGuide: "첫 화면 히어로 배경으로 사용하고, 실제 혜택 문구는 이미지 위가 아닌 별도 텍스트 레이어와 CTA로 분리합니다.",
        risk: "프리미엄 무드가 강해 혜택의 즉시성이 약해질 수 있으므로 쿠폰/기간 배지를 UI로 보강해야 합니다.",
        designNotes: ["중앙 숫자 대비를 가장 높게 유지", "노란 포인트는 CTA와 주요 배지에만 반복", "반사광은 좌상단에서 우하단으로 흐르게 구성", "배경은 어둡지만 디테일은 과밀하지 않게 처리"],
        asset: {
          texture: "Iridescent Silver",
          angle: "중앙 집중형 로우 앵글",
          background: "미니멀한 다크 모드 스튜디오",
          searchKeywords: ["KR premium event page mobile dark chrome", "iridescent chrome texture", "Pinterest black friday app event page premium"]
        },
        referenceKeywords: {
          mainVisual: ["KR premium event page mobile dark chrome", "Pinterest black friday app event page premium", "luxury membership promotion page Korea"],
          material: ["iridescent chrome texture", "transparent glass sphere render", "metallic ribbon 3d"],
          composition: ["centered hero object composition", "premium product pedestal lighting"],
          typography: ["condensed sans headline Adobe Fonts", "premium event typography"],
          landingLayout: ["ecommerce promotion hero section", "campaign landing page key visual"]
        },
        references: { mainVisual: [], typography: [], layout: [] },
        typography: {
          style: "Condensed + Massive Scale",
          features: ["세로 압축형", "초대형 사용", "로우앵글 시 웅장함 극대화"],
          fontCandidates: [
            { name: "Acumin Pro Condensed", source: "Adobe Fonts", usage: "숫자/헤드라인", reason: "압축 비율과 대형 사용 안정성이 좋아 프리미엄 키비주얼에 적합" },
            { name: "DIN Condensed", source: "Adobe Fonts", usage: "기간/혜택 라벨", reason: "테크 무드와 정보형 숫자 표현에 강함" },
            { name: "Spoqa Han Sans Neo", source: "Free", usage: "본문/CTA", reason: "한국어 UI 가독성이 높고 프로모션 정보 구조에 안정적" }
          ],
          pairingGuide: "Acumin Pro Condensed로 숫자와 헤드라인을 크게 잡고, Spoqa Han Sans Neo로 혜택/기간 정보를 안정적으로 분리합니다.",
          searchKeywords: ["condensed sans headline Adobe Fonts", "premium event typography", "large number campaign typography"]
        }
      },
      {
        title: "Benefit Carnival",
        koreanTitle: "쿠폰 페스티벌",
        hook: "혜택 요소를 티켓과 리본의 리듬으로 묶어 즉각적인 참여감을 강화",
        description: "쿠폰팩, 기간한정, 웰컴 혜택을 한 번에 읽히게 만드는 밝은 이벤트형 비주얼",
        visualNumber: "SALE",
        colors: [
          { hex: "#FF4D16", ratio: 56 },
          { hex: "#FFE500", ratio: 26 },
          { hex: "#111111", ratio: 18 }
        ],
        mainAsset: "쿠폰 티켓이 원형으로 터져 나오는 혜택 오브젝트와 숫자 배지",
        strategicReason: "기간한정 판매와 쿠폰팩 같은 즉시성 높은 스킴을 티켓 오브젝트로 시각화해 참여 이유를 빠르게 이해시킵니다.",
        targetEmotion: "혜택을 놓치면 아깝다는 긴장감과 즐거운 참여감",
        visualDirection: "쿠폰, 티켓, 배지가 원형으로 터지는 밝은 이벤트 스테이지와 선명한 혜택 오브젝트 중심의 3D 구성",
        imagePrompt: "Vibrant benefit carnival promotional key visual, isometric glossy coupon tickets bursting in a circular rhythm, sale badges, colorful ribbons, clean event stage, blue orange mint palette, high-quality 3D render, playful but polished commerce landing hero, clear empty space for UI overlay, No readable Korean text except simple number or symbol, no long copy inside image",
        negativePrompt: "crowded supermarket flyer, unreadable coupon text, cheap clipart, excessive confetti, real brand logos, readable Korean copy, low resolution",
        mobileCropGuide: "중앙 쿠폰 폭발 오브젝트와 대표 배지 1개가 세로형 크롭에서도 남도록 좌우 장식은 보조로 둡니다.",
        landingUsageGuide: "혜택 리스트 상단 또는 쿠폰 모듈 배경으로 사용하고, 각 스킴 pill과 함께 반복 노출합니다.",
        risk: "세일감이 강해 브랜드의 감사/기념 명분이 약해질 수 있으므로 상단 카피에서 milestone을 명확히 보완해야 합니다.",
        designNotes: ["쿠폰 안에는 실제 문장을 넣지 않음", "배지는 2~3개 크기 위계로 제한", "민트 컬러는 보조 액션에 사용", "모바일에서는 티켓 밀도를 70% 수준으로 줄임"],
        asset: {
          texture: "Glossy Paper",
          angle: "아이소메트릭 3/4 뷰",
          background: "밝은 그리드형 이벤트 스테이지",
          searchKeywords: ["KR coupon event page mobile", "isometric coupon burst composition", "app coupon sale event page design"]
        },
        referenceKeywords: {
          mainVisual: ["KR coupon event page mobile", "Pinterest benefit promotion page Korea", "app coupon sale event page design"],
          material: ["glossy paper ticket texture", "colorful ribbon 3d render", "rounded sale badge material"],
          composition: ["isometric coupon burst composition", "radial ticket layout promotion"],
          typography: ["bold rounded Korean font campaign", "sale badge typography system"],
          landingLayout: ["promotion coupon landing page", "benefit module ecommerce layout"]
        },
        references: { mainVisual: [], typography: [], layout: [] },
        typography: {
          style: "Bold Rounded + Badge System",
          features: ["짧은 혜택 문구 강조", "라벨형 배지 조합", "모바일에서도 높은 가독성"],
          fontCandidates: [
            { name: "SUIT", source: "Free", usage: "혜택 헤드라인", reason: "굵은 웨이트와 한국어 가독성이 안정적이며 무료 상업용 사용이 명확함" },
            { name: "Gmarket Sans", source: "Free", usage: "쿠폰/배지 문구", reason: "프로모션 배지에 어울리는 둥근 인상과 높은 주목도를 제공" },
            { name: "Sandoll GothicNeo", source: "Sandoll", usage: "서브카피/본문", reason: "상용 Sandoll 라이선스 범위에서 브랜드형 UI 문구에 안정적" }
          ],
          pairingGuide: "SUIT 또는 Gmarket Sans를 혜택 숫자에 쓰고, Spoqa Han Sans Neo나 Sandoll GothicNeo로 설명 문구를 받쳐 모바일 가독성을 유지합니다.",
          searchKeywords: ["bold rounded Korean font campaign", "sale badge typography system", "high readability coupon typography"]
        }
      },
      {
        title: "Growth Festival",
        koreanTitle: "함께 성장한 기념일",
        hook: "회원 수 달성을 브랜드와 사용자가 함께 만든 성장의 순간으로 해석",
        description: "커뮤니티의 확장, 감사, 환영의 메시지를 따뜻한 축하 무드로 전달",
        visualNumber: "700K",
        colors: [
          { hex: "#F6F0E4", ratio: 58 },
          { hex: "#FF6A9A", ratio: 24 },
          { hex: "#0E7A62", ratio: 18 }
        ],
        mainAsset: "성장 곡선을 따라 떠오르는 카드, 별, 리본, 환영 아이콘 클러스터",
        strategicReason: "70만 회원 달성을 브랜드와 유저가 함께 만든 성장의 순간으로 해석해 감사와 환영의 명분을 가장 자연스럽게 전달합니다.",
        targetEmotion: "함께 성장했다는 소속감, 따뜻한 축하감, 브랜드에 대한 호감",
        visualDirection: "밝은 배경 위 성장 곡선, 리본, 별, 카드 오브젝트가 부드럽게 떠오르는 커뮤니티형 축하 히어로",
        imagePrompt: "Warm growth festival campaign key visual, soft paper ribbons, floating cards, stars, welcome symbols, upward growth curve, friendly premium 3D paper craft style, green gold pink palette, bright celebration stage, balanced symmetrical hero composition, brand community milestone mood, No readable Korean text except simple number or symbol, no long copy inside image",
        negativePrompt: "childish party graphic, overly cute mascot, cluttered icons, readable Korean copy, long typography, generic stock celebration, harsh neon",
        mobileCropGuide: "성장 곡선의 시작과 끝이 세로 화면 안에서 읽히도록 중앙 세로축에 오브젝트를 집중시킵니다.",
        landingUsageGuide: "감사 메시지와 웰컴 혜택을 함께 보여주는 스토리형 랜딩 첫 구간에 적합합니다.",
        risk: "따뜻한 무드가 강해 구매 전환 압박이 약할 수 있으므로 하단 CTA와 기간한정 배지를 선명하게 배치해야 합니다.",
        designNotes: ["성장 곡선은 좌하단에서 우상단으로 배치", "금색은 milestone 강조에만 사용", "사람 이미지는 쓰지 않고 추상 커뮤니티 오브젝트로 표현", "배경은 밝고 여백을 충분히 둠"],
        asset: {
          texture: "Soft Gradient Paper",
          angle: "정면 대칭형 히어로 구도",
          background: "밝은 축하 무대와 페이퍼 오브젝트",
          searchKeywords: ["KR warm community event page mobile", "soft paper craft promotion page", "upward growth curve composition"]
        },
        referenceKeywords: {
          mainVisual: ["KR warm community event page mobile", "membership milestone event page Korea", "soft paper craft promotion page"],
          material: ["soft paper ribbon texture", "warm gradient paper craft", "floating card 3d render"],
          composition: ["upward growth curve composition", "friendly symmetric hero visual"],
          typography: ["humanist sans Korean typography", "friendly campaign headline type"],
          landingLayout: ["community campaign landing page", "milestone celebration hero section"]
        },
        references: { mainVisual: [], typography: [], layout: [] },
        typography: {
          style: "Humanist Sans + Friendly Weight",
          features: ["감사 메시지 중심", "부드러운 굵기 대비", "축하 문구와 숫자 병렬"],
          fontCandidates: [
            { name: "Spoqa Han Sans Neo", source: "Free", usage: "전체 UI/본문", reason: "한국어 랜딩 전반에 안정적인 기본값" },
            { name: "Source Han Sans", source: "Adobe Fonts", usage: "감사 메시지/본문", reason: "Adobe Fonts 범위에서 다국어와 긴 문장 가독성이 좋음" },
            { name: "Sandoll GothicNeo", source: "Sandoll", usage: "헤드라인/서브카피", reason: "부드러운 브랜드 톤과 명확한 Sandoll 라이선스 사용에 적합" }
          ],
          pairingGuide: "Spoqa Han Sans Neo 또는 Sandoll GothicNeo로 메시지를 따뜻하게 잡고, Source Han Sans로 긴 안내 문구의 안정성을 확보합니다.",
          searchKeywords: ["humanist sans Korean typography", "friendly campaign headline type", "milestone celebration typography"]
        }
      }
    ]
  };
}
