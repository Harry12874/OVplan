import assert from "node:assert/strict";
import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");

assert.match(html, /id="errorScreen"/);
assert.match(html, /id="errorReloadBtn"/);
assert.match(html, /id="errorCopyBtn"/);
assert.match(html, /id="errorDetails"/);

console.log("App shell markup tests passed.");
