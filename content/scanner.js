/* Scanner Module — extracts Mendeley Cite citations from .docx Content Controls (SDT) */

if (typeof MC2Zotero === "undefined") var MC2Zotero = {};

MC2Zotero.Scanner = {
  W_NS: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

  /**
   * Main scan entry point. Reads a .docx, finds Mendeley Cite SDT blocks,
   * and extracts citation metadata using whichever strategy works:
   *   Strategy A: CSL/citation JSON embedded in <w:tag> value
   *   Strategy B: Metadata in customXml/item*.xml
   *   Strategy C: Text-only parsing (fallback)
   */
  async scan(docxData) {
    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;
    const zip = await JSZipRef.loadAsync(docxData);

    const results = { citations: [], bibliography: null, errors: [], strategy: null };

    // Try to load customXml metadata (Strategy B)
    const customXmlMeta = await this._loadCustomXmlMetadata(zip);

    // Scan document.xml, footnotes.xml, endnotes.xml for SDT blocks
    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;
      const text = await file.async("string");
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(text, "application/xml");
      const context = {
        isFootnote: xmlPath === "word/footnotes.xml",
        isEndnote: xmlPath === "word/endnotes.xml"
      };
      if (typeof Zotero !== "undefined") {
        Zotero.debug("MC2Zotero Scanner: scanning " + xmlPath);
      }
      await this._extractSDTs(xmlDoc, results, context, customXmlMeta);
    }

    if (typeof Zotero !== "undefined") {
      Zotero.debug("MC2Zotero Scanner: found " + results.citations.length + " citations" +
        ", strategy=" + results.strategy + ", errors=" + results.errors.length);
    }
    return results;
  },

  /**
   * Load metadata from customXml/item*.xml files (Strategy B).
   * Mendeley Cite may store citation data in the customXml part of the DOCX.
   */
  async _loadCustomXmlMetadata(zip) {
    const meta = { items: new Map(), found: false };

    // Look for customXml files
    const customXmlFiles = [];
    zip.forEach((path, entry) => {
      if (path.startsWith("customXml/") && path.endsWith(".xml") && !path.endsWith(".rels")) {
        customXmlFiles.push(path);
      }
    });

    for (const xmlPath of customXmlFiles) {
      try {
        const text = await zip.file(xmlPath).async("string");
        // Look for Mendeley-related content
        if (text.indexOf("Mendeley") !== -1 || text.indexOf("mendeley") !== -1 ||
            text.indexOf("citation") !== -1 || text.indexOf("CSL") !== -1) {
          meta.found = true;
          if (typeof Zotero !== "undefined") {
            Zotero.debug("MC2Zotero Scanner: found Mendeley metadata in " + xmlPath);
          }

          // Try parsing as JSON array of citations
          try {
            const jsonMatch = text.match(/\[[\s\S]*\]/);
            if (jsonMatch) {
              const items = JSON.parse(jsonMatch[0]);
              if (Array.isArray(items)) {
                for (const item of items) {
                  const id = item.id || item.mendeleyId || item.citationId;
                  if (id) meta.items.set(String(id), item);
                }
              }
            }
          } catch (e) { /* not JSON, try XML parsing */ }

          // Try parsing as XML with citation elements
          try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(text, "application/xml");
            const citationEls = xmlDoc.querySelectorAll("citation, Citation, mendeleyCitation");
            for (const el of citationEls) {
              const id = el.getAttribute("id") || el.getAttribute("citationId");
              if (id) {
                meta.items.set(id, this._xmlCitationToCSL(el));
              }
            }
          } catch (e) { /* ignore */ }
        }
      } catch (e) {
        if (typeof Zotero !== "undefined") {
          Zotero.debug("MC2Zotero Scanner: error reading " + xmlPath + ": " + e.message);
        }
      }
    }

    return meta;
  },

  /**
   * Extract SDT (Structured Document Tag / Content Control) blocks
   * that belong to Mendeley Cite.
   */
  async _extractSDTs(xmlDoc, results, context, customXmlMeta) {
    const sdtElements = xmlDoc.getElementsByTagNameNS(this.W_NS, "sdt");

    for (let i = 0; i < sdtElements.length; i++) {
      const sdt = sdtElements[i];
      const sdtPr = this._getChildNS(sdt, "sdtPr");
      if (!sdtPr) continue;

      // Check if this SDT belongs to Mendeley Cite
      const alias = this._getChildNS(sdtPr, "alias");
      const tag = this._getChildNS(sdtPr, "tag");
      const aliasVal = alias ? alias.getAttribute("w:val") || "" : "";
      const tagVal = tag ? tag.getAttribute("w:val") || "" : "";

      const isMendeley = this._isMendeleySDT(aliasVal, tagVal);
      if (!isMendeley) continue;

      // Determine if citation or bibliography
      const isBibliography = this._isBibliographySDT(aliasVal, tagVal);

      // Extract display text from <w:sdtContent>
      const sdtContent = this._getChildNS(sdt, "sdtContent");
      const displayText = sdtContent ? this._extractTextFromContent(sdtContent) : "";

      // Get footnote/endnote context
      let footnoteId = null;
      let fnParent = sdt;
      while (fnParent && fnParent.parentNode) {
        fnParent = fnParent.parentNode;
        if (fnParent.localName === "footnote") {
          footnoteId = fnParent.getAttribute("w:id");
          break;
        }
        if (fnParent.localName === "endnote") {
          footnoteId = fnParent.getAttribute("w:id");
          break;
        }
      }

      if (isBibliography) {
        results.bibliography = {
          rawTagValue: tagVal,
          displayText: displayText,
          sdtIndex: i
        };
        continue;
      }

      // Try to extract citation metadata using available strategies
      let citationItems = null;

      // Strategy A: Parse metadata from <w:tag> value
      citationItems = this._parseTagMetadata(tagVal);
      if (citationItems && citationItems.length > 0) {
        if (!results.strategy) results.strategy = "tag_metadata";
      }

      // Strategy B: Look up in customXml metadata
      if ((!citationItems || citationItems.length === 0) && customXmlMeta.found) {
        citationItems = this._lookupCustomXmlMeta(tagVal, aliasVal, customXmlMeta);
        if (citationItems && citationItems.length > 0) {
          if (!results.strategy) results.strategy = "custom_xml";
        }
      }

      // Strategy C: Text parsing fallback
      if (!citationItems || citationItems.length === 0) {
        citationItems = this._parseDisplayText(displayText);
        if (citationItems && citationItems.length > 0) {
          if (!results.strategy) results.strategy = "text_parsing";
        }
      }

      if (!citationItems || citationItems.length === 0) {
        results.errors.push({
          type: "citation_no_data",
          message: "Could not extract metadata from Mendeley Cite content control",
          displayText: displayText,
          tagValue: tagVal
        });
        continue;
      }

      const citation = {
        fieldIndex: results.citations.length,
        rawTagValue: tagVal,
        mendeleyId: this._extractMendeleyId(tagVal, aliasVal) || "mc_" + results.citations.length,
        citationItems: citationItems,
        formattedText: displayText,
        sdtIndex: i,
        isInFootnote: context.isFootnote,
        isInEndnote: context.isEndnote,
        footnoteId: footnoteId
      };

      results.citations.push(citation);
    }
  },

  /* ---- SDT identification helpers ---- */

  _isMendeleySDT(aliasVal, tagVal) {
    const lower = (aliasVal + " " + tagVal).toLowerCase();
    return lower.indexOf("mendeley") !== -1 ||
           lower.indexOf("csl_citation") !== -1 ||
           lower.indexOf("mendeleycitationnote") !== -1 ||
           lower.indexOf("mendeley_citation") !== -1;
  },

  _isBibliographySDT(aliasVal, tagVal) {
    return tagVal.toUpperCase().startsWith("MENDELEY_BIBLIOGRAPHY") ||
           aliasVal.toLowerCase() === "mendeley bibliography" ||
           (aliasVal + " " + tagVal).toLowerCase().indexOf("mendeley_bibliography") !== -1;
  },

  _extractMendeleyId(tagVal, aliasVal) {
    // Mendeley Cite format: MENDELEY_CITATION_v3_<base64>
    // Extract citationID from decoded base64
    const mcMatch = tagVal.match(/MENDELEY_CITATION[_\s]v\d+[_\s](.*)/i);
    if (mcMatch) {
      try {
        const data = JSON.parse(atob(mcMatch[1]));
        if (data.citationID) return data.citationID;
      } catch (e) { /* ignore */ }
    }

    // Try to extract a UUID or citation ID from the tag value
    const uuidMatch = tagVal.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
    if (uuidMatch) return uuidMatch[1];

    // Try extracting citationID from JSON in tag
    try {
      const json = JSON.parse(tagVal);
      if (json.citationID) return json.citationID;
      if (json.mendpiId) return json.mendpiId;
    } catch (e) { /* not JSON */ }

    return null;
  },

  /* ---- Strategy A: Parse metadata from <w:tag> value ---- */

  _parseTagMetadata(tagVal) {
    if (!tagVal || tagVal.length < 5) return null;

    // Mendeley Cite may store CSL_CITATION JSON in the tag value
    // or a URL-encoded/escaped version of it
    let jsonStr = tagVal;

    // Try URL decode
    try {
      if (tagVal.indexOf("%7B") !== -1 || tagVal.indexOf("%22") !== -1) {
        jsonStr = decodeURIComponent(tagVal);
      }
    } catch (e) { /* use original */ }

    // Try HTML entity decode
    if (jsonStr.indexOf("&quot;") !== -1) {
      jsonStr = jsonStr.replace(/&quot;/g, '"').replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    }

    // Try parsing as CSL_CITATION JSON
    try {
      let data;

      // Handle "ADDIN CSL_CITATION {...}" prefix
      const cslMatch = jsonStr.match(/CSL_CITATION\s*(\{[\s\S]*\})/);
      if (cslMatch) {
        data = JSON.parse(cslMatch[1]);
      } else if (jsonStr.trim().startsWith("{")) {
        data = JSON.parse(jsonStr);
      }

      if (data) {
        return this._cslCitationToItems(data);
      }
    } catch (e) { /* not valid JSON */ }

    // Mendeley Cite format: MENDELEY_CITATION_v3_<base64blob>
    // The base64 decodes to a CSL_CITATION JSON with citationItems[].itemData
    const mendeleyMatch = tagVal.match(/MENDELEY_CITATION[_\s]v\d+[_\s](.*)/i);
    if (mendeleyMatch) {
      const b64 = mendeleyMatch[1];
      try {
        const decoded = atob(b64);
        const data = JSON.parse(decoded);
        // Mendeley Cite uses same structure as CSL_CITATION
        if (data.citationItems) {
          return this._cslCitationToItems(data);
        }
        // Fallback: array of items
        if (Array.isArray(data)) {
          return data.map(item => this._mendeleyItemToCSL(item));
        }
      } catch (e) { /* not valid base64/JSON */ }
    }

    return null;
  },

  /**
   * Convert a CSL_CITATION object (as used by Zotero/Mendeley) to our citation items format.
   */
  _cslCitationToItems(cslCitation) {
    if (!cslCitation.citationItems && !cslCitation.properties) return null;

    const items = cslCitation.citationItems || [];
    return items.map(ci => ({
      paperpileItemId: ci.id || "",
      mendeleyUUID: ci.mendeleyId || ci.id || "",
      cslData: ci.itemData || ci,
      locator: ci.locator || null,
      locatorType: ci.label || null,
      prefix: ci.prefix || null,
      suffix: ci.suffix || null,
      suppressAuthor: ci["suppress-author"] === true || ci.suppressAuthor === true
    }));
  },

  /* ---- Strategy B: customXml lookup ---- */

  _lookupCustomXmlMeta(tagVal, aliasVal, customXmlMeta) {
    // Try to find citation data by ID in customXml
    const id = this._extractMendeleyId(tagVal, aliasVal);
    if (id && customXmlMeta.items.has(id)) {
      const item = customXmlMeta.items.get(id);
      if (item.citationItems) {
        return item.citationItems.map(ci => ({
          paperpileItemId: ci.id || "",
          mendeleyUUID: ci.mendeleyId || ci.id || "",
          cslData: ci.itemData || ci,
          locator: ci.locator || null,
          locatorType: ci.label || null,
          prefix: ci.prefix || null,
          suffix: ci.suffix || null,
          suppressAuthor: ci["suppress-author"] === true
        }));
      }
      // Single item reference
      return [{
        paperpileItemId: id,
        mendeleyUUID: id,
        cslData: item,
        locator: null,
        locatorType: null,
        prefix: null,
        suffix: null,
        suppressAuthor: false
      }];
    }
    return null;
  },

  /* ---- Strategy C: Text parsing fallback ---- */

  _parseDisplayText(displayText) {
    if (!displayText || displayText.length < 3) return null;

    const text = displayText.trim();

    // Detect citation style and parse accordingly
    // Numeric: [1], [1,2,5], [1-3]
    const numericMatch = text.match(/^\[(\d[\d,\s\-–]+)\]$/);
    if (numericMatch) {
      return this._parseNumericCitation(numericMatch[1]);
    }

    // Author-year parenthetical: (Smith et al., 2020), (Smith & Jones, 2020; Brown, 2019)
    const parenMatch = text.match(/^\((.+)\)$/);
    if (parenMatch) {
      return this._parseAuthorYearCitation(parenMatch[1]);
    }

    // Narrative: Smith et al. (2020), Smith and Jones (2020)
    const narrativeMatch = text.match(/^(.+?)\s*\((\d{4}[a-z]?)\)$/);
    if (narrativeMatch) {
      return this._parseNarrativeCitation(narrativeMatch[1], narrativeMatch[2]);
    }

    // Superscript numbers: just a number (common in Vancouver style)
    const superscriptMatch = text.match(/^(\d+)$/);
    if (superscriptMatch) {
      return [{
        paperpileItemId: "",
        mendeleyUUID: "",
        cslData: { title: "", type: "article" },
        locator: null,
        locatorType: null,
        prefix: null,
        suffix: null,
        suppressAuthor: false,
        _textParsingNumericIndex: parseInt(superscriptMatch[1])
      }];
    }

    return null;
  },

  _parseNumericCitation(numStr) {
    const items = [];
    const parts = numStr.split(/[,;]/);
    for (const part of parts) {
      const trimmed = part.trim();
      const rangeMatch = trimmed.match(/^(\d+)\s*[-–]\s*(\d+)$/);
      if (rangeMatch) {
        const start = parseInt(rangeMatch[1]);
        const end = parseInt(rangeMatch[2]);
        for (let n = start; n <= end; n++) {
          items.push({
            paperpileItemId: "",
            mendeleyUUID: "",
            cslData: { title: "", type: "article" },
            locator: null, locatorType: null, prefix: null, suffix: null,
            suppressAuthor: false,
            _textParsingNumericIndex: n
          });
        }
      } else {
        const num = parseInt(trimmed);
        if (!isNaN(num)) {
          items.push({
            paperpileItemId: "",
            mendeleyUUID: "",
            cslData: { title: "", type: "article" },
            locator: null, locatorType: null, prefix: null, suffix: null,
            suppressAuthor: false,
            _textParsingNumericIndex: num
          });
        }
      }
    }
    return items.length > 0 ? items : null;
  },

  _parseAuthorYearCitation(innerText) {
    // Split multiple citations: "Smith, 2020; Brown, 2019"
    const parts = innerText.split(/;\s*/);
    const items = [];

    for (const part of parts) {
      const match = part.trim().match(/^(.+?),?\s*(\d{4}[a-z]?)(?:\s*,\s*(.+))?$/);
      if (!match) continue;

      const authorPart = match[1].trim();
      const year = match[2];
      const locator = match[3] || null;

      // Parse author: "Smith et al." or "Smith & Jones" or "Smith, Jones, & Brown"
      const isEtAl = /et\s+al\.?/i.test(authorPart);
      const authorClean = authorPart.replace(/\s*et\s+al\.?\s*/i, "").replace(/\s*&\s*/, ", ");
      const authorNames = authorClean.split(/,\s*/).filter(Boolean);

      const authors = authorNames.map(name => ({ family: name.trim(), given: "" }));
      if (isEtAl && authors.length > 0) {
        // Mark that there are more authors
        authors[0]._isEtAl = true;
      }

      items.push({
        paperpileItemId: "",
        mendeleyUUID: "",
        cslData: {
          title: "",
          type: "article",
          author: authors,
          issued: { "date-parts": [[parseInt(year)]] }
        },
        locator: locator,
        locatorType: locator ? "page" : null,
        prefix: null,
        suffix: null,
        suppressAuthor: false
      });
    }

    return items.length > 0 ? items : null;
  },

  _parseNarrativeCitation(authorPart, year) {
    const isEtAl = /et\s+al\.?/i.test(authorPart);
    const authorClean = authorPart.replace(/\s*et\s+al\.?\s*/i, "")
      .replace(/\s+and\s+/i, ", ").replace(/\s*&\s*/, ", ");
    const authorNames = authorClean.split(/,\s*/).filter(Boolean);

    const authors = authorNames.map(name => ({ family: name.trim(), given: "" }));

    return [{
      paperpileItemId: "",
      mendeleyUUID: "",
      cslData: {
        title: "",
        type: "article",
        author: authors,
        issued: { "date-parts": [[parseInt(year)]] }
      },
      locator: null,
      locatorType: null,
      prefix: null,
      suffix: null,
      suppressAuthor: true // narrative citation = author outside parens
    }];
  },

  /* ---- Conversion helpers ---- */

  _mendeleyItemToCSL(item) {
    // Convert Mendeley-specific item format to CSL-JSON
    const csl = {
      title: item.title || "",
      DOI: item.doi || item.DOI || "",
      PMID: item.pmid || "",
      ISBN: item.isbn || ""
    };

    if (item.authors || item.author) {
      const authorList = item.authors || item.author;
      csl.author = (Array.isArray(authorList) ? authorList : []).map(a => {
        if (typeof a === "string") {
          const parts = a.split(/,\s*/);
          return { family: parts[0] || "", given: parts[1] || "" };
        }
        return { family: a.last || a.family || a.lastName || "", given: a.first || a.given || a.firstName || "" };
      });
    }

    if (item.year) {
      csl.issued = { "date-parts": [[parseInt(item.year)]] };
    } else if (item.issued) {
      csl.issued = item.issued;
    }

    if (item.source || item.journal || item["container-title"]) {
      csl["container-title"] = item.source || item.journal || item["container-title"];
    }
    if (item.volume) csl.volume = item.volume;
    if (item.issue) csl.issue = item.issue;
    if (item.pages || item.page) csl.page = item.pages || item.page;
    if (item.publisher) csl.publisher = item.publisher;
    if (item.url || item.URL) csl.URL = item.url || item.URL;

    // Map type
    const typeMap = {
      "journal": "article-journal",
      "book": "book",
      "book_section": "chapter",
      "conference_paper": "paper-conference",
      "thesis": "thesis",
      "report": "report",
      "web_page": "webpage",
      "generic": "article"
    };
    csl.type = typeMap[item.type] || item.type || "article";

    return {
      paperpileItemId: item.id || "",
      mendeleyUUID: item.mendeleyId || item.id || "",
      cslData: csl,
      locator: null,
      locatorType: null,
      prefix: null,
      suffix: null,
      suppressAuthor: false
    };
  },

  _xmlCitationToCSL(el) {
    const csl = {};
    const getEl = (name) => {
      const child = el.querySelector(name);
      return child ? child.textContent : "";
    };
    csl.title = getEl("title");
    csl.DOI = getEl("doi") || getEl("DOI");
    csl.type = getEl("type") || "article";
    // Add more fields as needed
    return csl;
  },

  /* ---- DOM helpers ---- */

  _getChildNS(parent, localName) {
    for (let i = 0; i < parent.childNodes.length; i++) {
      const n = parent.childNodes[i];
      if (n.nodeType === 1 && n.localName === localName && n.namespaceURI === this.W_NS) return n;
    }
    return null;
  },

  _extractTextFromContent(sdtContent) {
    // Recursively extract all text from <w:t> elements within the SDT content
    let text = "";
    const walk = (node) => {
      if (node.nodeType === 1) {
        if (node.localName === "t" && node.namespaceURI === this.W_NS) {
          text += node.textContent;
        } else if (node.localName === "br") {
          text += "\n";
        } else if (node.localName === "tab") {
          text += "\t";
        }
        for (let i = 0; i < node.childNodes.length; i++) {
          walk(node.childNodes[i]);
        }
      }
    };
    walk(sdtContent);
    return text;
  }
};
