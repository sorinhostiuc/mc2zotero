/* Converter Module — rewrites Mendeley Cite Content Controls to Zotero format in .docx */

if (typeof MC2Zotero === "undefined") var MC2Zotero = {};

MC2Zotero.Converter = {
  ZOTERO_SCHEMA: "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",

  async convert(docxData, matchResults, options = {}) {
    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;
    const zip = await JSZipRef.loadAsync(docxData);

    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    const allComments = [];
    let nextCommentId = 1;

    // Reference marks mode delegates to ODF writer (outputs .odt)
    if (options.fieldMode === "referencemarks") {
      return await MC2Zotero.ODFWriter.convertToODT(docxData, matchResults, options);
    }

    const useBookmarks = options.fieldMode === "bookmarks";
    const allBookmarkData = {};

    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;

      let xmlText = await file.async("string");

      if (useBookmarks) {
        const result = this._rewriteSDTsAsBookmarks(xmlText, matchResults, options);
        xmlText = result.xmlText;
        Object.assign(allBookmarkData, result.bookmarkData);
      } else {
        xmlText = this._rewriteSDTs(xmlText, matchResults, options);
      }

      if (options.addComments) {
        const result = this._addSkippedComments(xmlText, matchResults, nextCommentId);
        xmlText = result.xmlText;
        allComments.push(...result.comments);
        nextCommentId += result.comments.length;
      }

      zip.file(xmlPath, xmlText);
    }

    if (allComments.length > 0) {
      zip.file("word/comments.xml", this._buildCommentsXml(allComments));
      await this._ensureCommentsRelationship(zip);
    }

    if (useBookmarks && Object.keys(allBookmarkData).length > 0) {
      await this._updateSettingsWithDocVars(zip, allBookmarkData, options);
    } else {
      // Field codes mode: still need to write ZOTERO_PREF so Zotero knows
      // the citation style (without this, Refresh defaults to numeric style)
      await this._updateSettingsWithDocVars(zip, {}, options);
    }

    return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  },

  /**
   * Find all Mendeley Cite SDT blocks in raw XML text and replace with Zotero field codes.
   */
  _rewriteSDTs(xmlText, matchResults, options) {
    const blocks = this._findMendeleyCiteSDTBlocks(xmlText);

    // Process in reverse to preserve string positions
    for (let i = blocks.length - 1; i >= 0; i--) {
      const block = blocks[i];

      if (block.isBibliography) {
        const zoteroBibl = { uncited: [], omitted: [], custom: [] };
        const escaped = this._escapeXml(JSON.stringify(zoteroBibl));
        const beginRuns = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
          + '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL ' + escaped + ' CSL_BIBLIOGRAPHY </w:instrText></w:r>'
          + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
        const endRun = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
        const replacement = this._wrapFieldCodeAroundContent(beginRuns, block.displayContent, endRun);
        xmlText = xmlText.substring(0, block.startPos) + replacement + xmlText.substring(block.endPos);
        continue;
      }

      if (!block.isCitation) continue;

      // Find matching result
      const matchResult = this._findMatchResult(block, matchResults);
      if (!matchResult) continue;

      const allSkipped = matchResult.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (allSkipped) continue;

      const zoteroCitation = this._buildZoteroCitation(matchResult);
      if (!zoteroCitation) continue;

      const zoteroJSON = JSON.stringify(zoteroCitation);
      const escaped = this._escapeXml(zoteroJSON);

      const beginRuns = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        + '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION ' + escaped + ' </w:instrText></w:r>'
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
      const endRun = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
      const replacement = this._wrapFieldCodeAroundContent(beginRuns, block.displayContent, endRun);
      xmlText = xmlText.substring(0, block.startPos) + replacement + xmlText.substring(block.endPos);
    }

    return xmlText;
  },

  /**
   * Find the match result for a given SDT block by trying multiple strategies.
   */
  _findMatchResult(block, matchResults) {
    for (const mr of matchResults) {
      const cit = mr.citation;
      // Match by mendeleyId appearing in tagValue
      if (cit.mendeleyId && block.tagValue.indexOf(cit.mendeleyId) !== -1) return mr;
      // Match by SDT index (both count all SDTs in DOM order)
      if (cit.sdtIndex === block.blockIndex) return mr;
      // Match by mendeleyId pattern "mc_N" (both count only Mendeley citation SDTs)
      if (block.mendeleyIndex !== undefined && cit.mendeleyId === "mc_" + block.mendeleyIndex) return mr;
    }
    return null;
  },

  _buildZoteroCitation(matchResult) {
    const items = matchResult.itemMatches
      .filter(m => m.match !== null)
      .map(m => {
        const zItem = m.match;
        const cslJSON = MC2Zotero.Matcher.getCSLJSON(zItem);
        const uri = MC2Zotero.Matcher.buildURI(zItem);

        const citItem = {
          id: zItem.id,
          uris: [uri],
          uri: [uri],
          itemData: cslJSON
        };

        if (m.locator) citItem.locator = m.locator;
        if (m.locatorType) citItem.label = m.locatorType;
        if (m.prefix) citItem.prefix = m.prefix;
        if (m.suffix) citItem.suffix = m.suffix;
        if (m.suppressAuthor) citItem["suppress-author"] = true;

        return citItem;
      });

    if (items.length === 0) return null;

    const noteIndex = matchResult.citation.isInFootnote || matchResult.citation.isInEndnote
      ? (matchResult.citation.fieldIndex + 1) : 0;

    return {
      citationID: this._generateCitationID(),
      properties: {
        formattedCitation: "",
        plainCitation: "",
        noteIndex: noteIndex
      },
      citationItems: items,
      schema: this.ZOTERO_SCHEMA
    };
  },

  _generateCitationID() {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "";
    for (let i = 0; i < 8; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  },

  _escapeXml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  },

  /**
   * Wrap field code begin/end runs around displayContent.
   * For block-level content (contains <w:p>), injects begin runs into the
   * first <w:p> and end run into the last <w:p> to keep OOXML valid.
   * For inline content (no <w:p>), places runs directly around content.
   */
  _wrapFieldCodeAroundContent(beginRuns, displayContent, endRun) {
    const firstPMatch = displayContent.match(/<w:p[\s>]/);
    if (!firstPMatch) {
      // Inline content — runs are already inside a parent <w:p>
      return beginRuns + displayContent + endRun;
    }

    // Block-level: find insertion point after the first <w:p ...> opening tag
    const firstPPos = firstPMatch.index;
    // Find the closing ">" of this <w:p> or <w:p ...> tag
    const firstPClose = displayContent.indexOf('>', firstPPos);

    // Find the last </w:p> for the end run
    const lastPClose = displayContent.lastIndexOf('</w:p>');

    // Insert begin runs right after first <w:p...>, end run right before last </w:p>
    return displayContent.substring(0, firstPClose + 1)
      + beginRuns
      + displayContent.substring(firstPClose + 1, lastPClose)
      + endRun
      + displayContent.substring(lastPClose);
  },

  _unescapeXml(str) {
    return str
      .replace(/&quot;/g, '"')
      .replace(/&gt;/g, '>')
      .replace(/&lt;/g, '<')
      .replace(/&amp;/g, '&');
  },

  /**
   * Find all Mendeley Cite SDT blocks in raw XML text.
   * Uses iterative approach to handle nested SDTs correctly.
   */
  _findMendeleyCiteSDTBlocks(xmlText) {
    const blocks = [];
    let searchFrom = 0;
    let blockIndex = 0;       // counts ALL SDTs (matches Scanner sdtIndex)
    let mendeleyIndex = 0;    // counts only Mendeley citation SDTs (matches Scanner "mc_N" fallback)

    while (searchFrom < xmlText.length) {
      const sdtStart = xmlText.indexOf("<w:sdt>", searchFrom);
      const sdtStartAlt = xmlText.indexOf("<w:sdt ", searchFrom);

      let startPos = -1;
      if (sdtStart !== -1 && sdtStartAlt !== -1) {
        startPos = Math.min(sdtStart, sdtStartAlt);
      } else if (sdtStart !== -1) {
        startPos = sdtStart;
      } else if (sdtStartAlt !== -1) {
        startPos = sdtStartAlt;
      } else {
        break;
      }

      let endPos = this._findMatchingSDTEnd(xmlText, startPos);
      if (endPos === -1) {
        searchFrom = startPos + 7;
        continue;
      }

      const fullBlock = xmlText.substring(startPos, endPos);

      const sdtPrMatch = fullBlock.match(/<w:sdtPr[^>]*>([\s\S]*?)<\/w:sdtPr>/);
      if (!sdtPrMatch) {
        searchFrom = endPos;
        blockIndex++;
        continue;
      }

      const sdtPr = sdtPrMatch[1];
      const aliasMatch = sdtPr.match(/<w:alias\s+w:val="([^"]*)"/);
      const tagMatch = sdtPr.match(/<w:tag\s+w:val="([^"]*)"/);
      const aliasVal = aliasMatch ? aliasMatch[1] : "";
      const tagVal = tagMatch ? this._unescapeXml(tagMatch[1]) : "";

      const combined = (aliasVal + " " + tagVal).toLowerCase();
      const isMendeley = combined.indexOf("mendeley") !== -1 ||
                         combined.indexOf("csl_citation") !== -1 ||
                         combined.indexOf("mendeleycitationnote") !== -1 ||
                         combined.indexOf("mendeley_citation") !== -1;

      if (!isMendeley) {
        searchFrom = endPos;
        blockIndex++;
        continue;
      }

      const isBibliography = combined.indexOf("mendeley_bibliography") !== -1 ||
                             aliasVal.toLowerCase() === "mendeley bibliography" ||
                             (tagVal.toUpperCase().startsWith("MENDELEY_BIBLIOGRAPHY"));

      const contentMatch = fullBlock.match(/<w:sdtContent[^>]*>([\s\S]*?)<\/w:sdtContent>/);
      const displayContent = contentMatch ? contentMatch[1] : "";

      blocks.push({
        startPos: startPos,
        endPos: endPos,
        tagValue: tagVal,
        aliasValue: aliasVal,
        displayContent: displayContent,
        isCitation: !isBibliography,
        isBibliography: isBibliography,
        blockIndex: blockIndex,
        mendeleyIndex: mendeleyIndex
      });

      if (!isBibliography) mendeleyIndex++;
      searchFrom = endPos;
      blockIndex++;
    }

    return blocks;
  },

  _findMatchingSDTEnd(xmlText, startPos) {
    let depth = 0;
    let pos = startPos;

    while (pos < xmlText.length) {
      const openIdx = xmlText.indexOf("<w:sdt", pos + 1);
      const closeIdx = xmlText.indexOf("</w:sdt>", pos + 1);

      if (closeIdx === -1) return -1;

      if (openIdx !== -1 && openIdx < closeIdx) {
        const afterTag = xmlText[openIdx + 6];
        if (afterTag === ">" || afterTag === " ") {
          depth++;
          pos = openIdx;
          continue;
        }
      }

      if (depth === 0) {
        return closeIdx + 8;
      }

      depth--;
      pos = closeIdx;
    }

    return -1;
  },

  /* ---- Bookmark mode ---- */

  _rewriteSDTsAsBookmarks(xmlText, matchResults, options) {
    const bookmarkData = {};
    let nextBmId = this._getMaxBookmarkId(xmlText) + 1;
    const blocks = this._findMendeleyCiteSDTBlocks(xmlText);

    for (let i = blocks.length - 1; i >= 0; i--) {
      const block = blocks[i];

      if (block.isBibliography) {
        const bmName = this._generateBookmarkName();
        const bmId = nextBmId++;
        const zoteroBibl = { uncited: [], omitted: [], custom: [] };
        const val = " ADDIN ZOTERO_BIBL " + JSON.stringify(zoteroBibl) + " CSL_BIBLIOGRAPHY ";
        bookmarkData[bmName] = val;

        const replacement = '<w:bookmarkStart w:id="' + bmId + '" w:name="' + bmName + '"/>'
          + block.displayContent
          + '<w:bookmarkEnd w:id="' + bmId + '"/>';
        xmlText = xmlText.substring(0, block.startPos) + replacement + xmlText.substring(block.endPos);
        continue;
      }

      if (!block.isCitation) continue;

      const matchResult = this._findMatchResult(block, matchResults);
      if (!matchResult) continue;

      const allSkipped = matchResult.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (allSkipped) continue;

      const zoteroCitation = this._buildZoteroCitation(matchResult);
      if (!zoteroCitation) continue;

      const bmName = this._generateBookmarkName();
      const bmId = nextBmId++;
      const val = " ADDIN ZOTERO_ITEM CSL_CITATION " + JSON.stringify(zoteroCitation) + " ";
      bookmarkData[bmName] = val;

      const replacement = '<w:bookmarkStart w:id="' + bmId + '" w:name="' + bmName + '"/>'
        + block.displayContent
        + '<w:bookmarkEnd w:id="' + bmId + '"/>';
      xmlText = xmlText.substring(0, block.startPos) + replacement + xmlText.substring(block.endPos);
    }

    return { xmlText, bookmarkData };
  },

  _addSkippedComments(xmlText, matchResults, startId) {
    const comments = [];
    const skippedCitations = [];

    for (const mr of matchResults) {
      const allSkipped = mr.itemMatches.every(m =>
        m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
      );
      if (!allSkipped) continue;

      const mendeleyId = mr.citation.mendeleyId;
      if (!mendeleyId) continue;

      const title = mr.itemMatches[0]?.cslData?.title || "Unknown reference";
      const authors = (mr.itemMatches[0]?.cslData?.author || [])
        .map(a => a.family || a.given || "").filter(Boolean).join(", ");

      skippedCitations.push({ mendeleyId, title, authors, formattedText: mr.citation.formattedText });
    }

    const positions = [];
    for (const sc of skippedCitations) {
      const searchText = sc.formattedText;
      if (!searchText) continue;
      const escaped = this._escapeXml(searchText);
      let pos = xmlText.indexOf(escaped);
      if (pos === -1) pos = xmlText.indexOf(searchText);
      if (pos !== -1) {
        positions.push({ pos, ...sc });
      }
    }
    positions.sort((a, b) => b.pos - a.pos);

    for (const item of positions) {
      const commentId = startId + comments.length;
      const commentText = "MC2Zotero: This citation was not converted to Zotero format.\n"
        + (item.title ? "Title: " + item.title + "\n" : "")
        + (item.authors ? "Authors: " + item.authors : "");

      let pStart = xmlText.lastIndexOf("<w:p ", item.pos);
      const pStart2 = xmlText.lastIndexOf("<w:p>", item.pos);
      if (pStart2 > pStart) pStart = pStart2;
      const pEnd = xmlText.indexOf("</w:p>", item.pos);

      if (pStart === -1 || pEnd === -1) continue;

      const pOpenEnd = xmlText.indexOf(">", pStart) + 1;
      const startMarker = '<w:commentRangeStart w:id="' + commentId + '"/>';
      const endMarker = '<w:commentRangeEnd w:id="' + commentId + '"/>'
        + '<w:r><w:commentReference w:id="' + commentId + '"/></w:r>';

      xmlText = xmlText.substring(0, pEnd) + endMarker + xmlText.substring(pEnd);
      xmlText = xmlText.substring(0, pOpenEnd) + startMarker + xmlText.substring(pOpenEnd);

      comments.push({ id: commentId, text: commentText });
    }

    return { xmlText, comments };
  },

  _buildCommentsXml(comments) {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<w:comments xmlns:w="' + W_NS + '"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

    const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    for (const c of comments) {
      const lines = c.text.split("\n");
      xml += '<w:comment w:id="' + c.id + '" w:author="MC2Zotero" w:date="' + now + '">';
      for (const line of lines) {
        xml += '<w:p><w:r><w:t>' + this._escapeXml(line) + '</w:t></w:r></w:p>';
      }
      xml += '</w:comment>';
    }

    xml += '</w:comments>';
    return xml;
  },

  async _ensureCommentsRelationship(zip) {
    const ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      let ct = await ctFile.async("string");
      if (ct.indexOf("comments.xml") === -1) {
        ct = ct.replace("</Types>",
          '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
        zip.file("[Content_Types].xml", ct);
      }
    }

    const relsFile = zip.file("word/_rels/document.xml.rels");
    if (relsFile) {
      let rels = await relsFile.async("string");
      if (rels.indexOf("comments.xml") === -1) {
        const idMatches = [...rels.matchAll(/Id="rId(\d+)"/g)];
        const maxId = idMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
        const newId = "rId" + (maxId + 1);
        rels = rels.replace("</Relationships>",
          '<Relationship Id="' + newId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>');
        zip.file("word/_rels/document.xml.rels", rels);
      }
    }
  },

  _generateBookmarkName() {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "ZOTERO_BREF_";
    for (let i = 0; i < 10; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  },

  _getMaxBookmarkId(xmlText) {
    const idRegex = /w:id="(\d+)"/g;
    let max = 0;
    let m;
    while ((m = idRegex.exec(xmlText)) !== null) {
      const val = parseInt(m[1], 10);
      if (val > max) max = val;
    }
    return max;
  },

  async _updateSettingsWithDocVars(zip, bookmarkData, options = {}) {
    const settingsFile = zip.file("word/settings.xml");
    if (!settingsFile) return;

    let settings = await settingsFile.async("string");

    let docVarsContent = "";

    const sessionId = this._generateCitationID();
    const useBookmarks = options.fieldMode === "bookmarks";
    const fieldType = useBookmarks ? "Bookmark" : "Field";

    // Build ZOTERO_PREF XML — includes citation style so Zotero applies
    // the correct format (e.g., Harvard author-date) on Refresh.
    // If citationStyle is empty, omit <style> and let Zotero use its default.
    const citationStyle = options.citationStyle || "";

    const styleTag = citationStyle
      ? '<style id="' + citationStyle + '" hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
      : '';

    const prefXml = '<data data-version="3" zotero-version="7.0.0">'
      + '<session id="' + sessionId + '"/>'
      + styleTag
      + '<prefs>'
      + '<pref name="fieldType" value="' + fieldType + '"/>'
      + '</prefs>'
      + '</data>';
    docVarsContent += '<w:docVar w:name="ZOTERO_PREF_1" w:val="' + this._escapeXml(prefXml) + '"/>';

    for (const [name, val] of Object.entries(bookmarkData)) {
      docVarsContent += '<w:docVar w:name="' + this._escapeXml(name) + '" w:val="' + this._escapeXml(val) + '"/>';
    }

    if (settings.indexOf("<w:docVars>") !== -1) {
      settings = settings.replace("</w:docVars>", docVarsContent + "</w:docVars>");
    } else {
      settings = settings.replace("</w:settings>", "<w:docVars>" + docVarsContent + "</w:docVars></w:settings>");
    }

    zip.file("word/settings.xml", settings);
  },

  async createBackup(filePath) {
    const backupPath = filePath + ".bak";
    await IOUtils.copy(filePath, backupPath);
    Zotero.debug("MC2Zotero: Backup created at " + backupPath);
    return backupPath;
  },

  async saveFile(filePath, data) {
    await IOUtils.write(filePath, data);
    Zotero.debug("MC2Zotero: Saved converted file to " + filePath);
  }
};
