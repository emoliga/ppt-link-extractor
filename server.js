import express from "express";
import multer from "multer";
import JSZip from "jszip";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

function extractUrlsFromRelsXml(xml) {
  const urls = [];
  const regex = /<Relationship\b[^>]*Type="[^"]*\/hyperlink"[^>]*Target="([^"]+)"[^>]*TargetMode="External"[^>]*\/?>/gi;

  let match;
  while ((match = regex.exec(xml)) !== null) {
    const url = match[1].replace(/&amp;/g, '&').trim();
    if (/^https?:\/\//i.test(url)) {
      urls.push(url);
    }
  }

  return [...new Set(urls)];
}

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

    const relPaths = Object.keys(zip.files).filter((name) =>
      /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(name)
    );

    const allUrls = [];
    const urlsBySlide = {};

    for (const relPath of relPaths) {
      const xml = await zip.files[relPath].async("string");

      const slideMatch = relPath.match(/slide(\d+)\.xml\.rels$/i);
      const slideNumber = slideMatch ? Number(slideMatch[1]) : null;

      const urls = extractUrlsFromRelsXml(xml);

      if (urls.length > 0) {
        urlsBySlide[String(slideNumber)] = urls;
        allUrls.push(...urls);
      }
    }

    const uniqueUrls = [...new Set(allUrls)];

    res.json({
      ok: true,
      fileName,
      count: uniqueUrls.length,
      urls: uniqueUrls,
      urlsBySlide
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message || "Unexpected error"
    });
  }
});

const port = process.env.PORT || 10000;
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
