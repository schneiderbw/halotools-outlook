// Public surface of @iusehalo/halo-api.
//
// Pure logic — no React, no Fluent, no Office.js. Consumers install a storage
// adapter (via setStorage) and, for sign-in, supply a DialogOpener.

export * from "./api.js";
export * from "./auth.js";
export * from "./config.js";
export * from "./pkce.js";
export * from "./storage.js";
export * from "./types.js";
