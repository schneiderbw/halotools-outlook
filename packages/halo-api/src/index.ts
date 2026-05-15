// Public surface of @iusehalo/halo-api.
//
// Pure logic — no React, no Fluent, no Office.js. Consumers install a storage
// adapter (via setStorage) and, for sign-in, supply a DialogOpener.

export * from "./api";
export * from "./auth";
export * from "./config";
export * from "./pkce";
export * from "./storage";
export * from "./types";
