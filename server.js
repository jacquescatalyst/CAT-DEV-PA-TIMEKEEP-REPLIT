const express = require("express");
const path = require("path");

const app = express();
const PORT = 5000;

app.use((_req, res, next) => {
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  next();
});

app.use(express.static(path.join(__dirname, "public")));
app.use("/dist", express.static(path.join(__dirname, "dist")));

app.get("/{*path}", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Time Keep Phase 2 server running on http://0.0.0.0:${PORT}`);
});
