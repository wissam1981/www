/* eslint-disable no-console */
const https = require("https");
const express = require("express");
const path = require("path");
const devCerts = require("@microsoft/office-addin-dev-certs");

const PORT = process.env.PORT || 3000;
const app = express();
app.use(express.static(path.join(__dirname, "public")));
app.use("/assets", express.static(path.join(__dirname, "assets")));

async function start() {
  const options = await devCerts.getHttpsServerOptions();
  if (!options || !options.cert || !options.key) {
    await devCerts.install();
  }
  const httpsOptions = await devCerts.getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`HTTPS dev server running at https://localhost:${PORT}`);
  });
}

start().catch((error) => {
  console.error("Failed to start dev server", error);
  process.exit(1);
});
