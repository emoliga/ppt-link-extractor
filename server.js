import express from "express";
import multer from "multer";
import JSZip from "jszip";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

function unique(arr) {
  return [...new Set(arr.filter(Boolean))];
}

function decodeXml(str) {
  return String(str || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

function extractUrlsFromRelsXml(xml) {
  const urls = [];

  const regex = /<Relationship\b[^>]*\bType="[^"]*\/hyperlink"[^>]*\bTarget="([^"]+)"[^>]*\bTargetMode="External"[^>]*\/?>/gi;

  let match;
  while ((match = regex.exec(xml)) !== null) {
    const url = decodeXml(match[1]);
    if (/^https?:\/\//i.test(url)) {
      urls.push(url);
    }
  }

  return unique(urls);
}

app.get("/", (_req, res) => {
  res.json({
    ok: true,
    service: "ppt-link-extractor",
    endpoints: {
      health: "/health",
      extract: "/extract-urls"
    }
  });
});

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.post("/extract-urls", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        ok: false,
        error: "No file uploaded. Use field name 'file'."
      });
    }

    const fileName = req.file.originalname || "unknown";
    const buffer = req.file.buffer;

    const zip = await JSZip.loadAsync(buffer);
    const zipNames = Object.keys(zip.files);

    const relPaths = zipNames.filter((name) =>
      /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(name)
    );

    const allUrls = [];
    const urlsBySlide = {};

    for (const relPath of relPaths) {
      const xml = await zip.files[relPath].async("string");

      const slideMatch = relPath.match(/slide(\d+)\.xml\.rels$/i);
      const slideNumber = slideMatch ? Number(slideMatch[1]) : null;

      const urls = extractUrlsFromRelsXml(xml);

      if (slideNumber !== null) {
        urlsBySlide[String(slideNumber)] = urls;
      }

      allUrls.push(...urls);
    }

    const uniqueUrls = unique(allUrls);

    return res.json({
      ok: true,
      fileName,
      count: uniqueUrls.length,
      urls: uniqueUrls,
      urlsBySlide,
      relFilesFound: relPaths.length
    });
  } catch (error) {
    console.error("Extraction error:", error);
    return res.status(500).json({
      ok: false,
      error: error.message || "Unexpected error"
    });
  }
});

const port = process.env.PORT || 10000;
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
