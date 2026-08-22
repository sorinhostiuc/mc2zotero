var MC2Zotero;
var chromeHandle;

// Polyfill setImmediate for JSZip (not available in Gecko)
if (typeof setImmediate === "undefined") {
  var setImmediate = fn => setTimeout(fn, 0);
}

function install(data, reason) {}

function startup({ id, version, resourceURI, rootURI }, reason) {
  // Register chrome://mc2zotero/content/ mapping
  var aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"]
    .getService(Ci.amIAddonManagerStartup);
  var manifestURI = Services.io.newURI(rootURI + "manifest.json");
  chromeHandle = aomStartup.registerChrome(manifestURI, [
    ["content", "mc2zotero", rootURI + "content/"]
  ]);

  var chromeBase = "chrome://mc2zotero/content/";
  Services.scriptloader.loadSubScript(chromeBase + "lib/jszip.min.js");
  Services.scriptloader.loadSubScript(chromeBase + "mc2zotero.js");
  Services.scriptloader.loadSubScript(chromeBase + "scanner.js");
  Services.scriptloader.loadSubScript(chromeBase + "matcher.js");
  Services.scriptloader.loadSubScript(chromeBase + "converter.js");
  Services.scriptloader.loadSubScript(chromeBase + "odf-writer.js");

  MC2Zotero.init({ id, version, rootURI });
}

function shutdown({ id, version, resourceURI, rootURI }, reason) {
  if (reason === APP_SHUTDOWN) return;
  MC2Zotero.shutdown();
  MC2Zotero = undefined;
  if (chromeHandle) {
    chromeHandle.destruct();
    chromeHandle = null;
  }
}

function uninstall(data, reason) {}
