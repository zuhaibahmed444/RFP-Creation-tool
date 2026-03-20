import dotenv from "dotenv";
dotenv.config();

import express from "express";
import path from "path";
import { generatePocHandler } from "./controller";

const app = express();
const PORT = process.env.PORT || 4500;

app.use(express.json());
app.use(express.static(path.join(__dirname, "..", "public")));
app.use("/assets", express.static(path.join(__dirname, "..", "assets")));

app.post("/generate-poc", generatePocHandler);

app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
