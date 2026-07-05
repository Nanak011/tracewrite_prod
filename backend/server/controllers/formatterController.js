const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");
const { query } = require("../../db");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const mammoth = require("mammoth");

// Windows Python installs are almost always exposed as `python`, not
// `python3` - macOS/Linux usually have both, but `python3` is the safe
// default there. Override with the PYTHON_BIN env var if your setup
// differs (e.g. a venv with a specific interpreter path).
const PYTHON_BIN = process.env.PYTHON_BIN || (process.platform === "win32" ? "python" : "python3");

// Files older than this get swept up regardless of download count. This
// replaces the old "delete 5s after every download" behavior, which is
// what caused "I can only download once" - deletion is now decoupled
// from downloads entirely.
const TEMP_MAX_AGE_MS = 30 * 60 * 1000; // 30 minutes
let cleanupIntervalStarted = false;

function cleanupOldTempFiles(tempDir) {
  try {
    const now = Date.now();
    for (const name of fs.readdirSync(tempDir)) {
      const filePath = path.join(tempDir, name);
      try {
        const stat = fs.statSync(filePath);
        if (now - stat.mtimeMs > TEMP_MAX_AGE_MS) {
          fs.unlinkSync(filePath);
        }
      } catch (e) { /* file may have been removed already - fine */ }
    }
    // Prune formattedDocs entries pointing at files that no longer exist,
    // so downloadFormatted gives a clean 404 instead of a stale mapping.
    global.formattedDocs = global.formattedDocs || {};
    for (const [docId, filePath] of Object.entries(global.formattedDocs)) {
      if (!fs.existsSync(filePath)) {
        delete global.formattedDocs[docId];
      }
    }
  } catch (e) {
    console.error("Temp cleanup sweep failed:", e);
  }
}

function ensureCleanupInterval(tempDir) {
  if (cleanupIntervalStarted) return;
  cleanupIntervalStarted = true;
  setInterval(() => cleanupOldTempFiles(tempDir), 10 * 60 * 1000);
}

/**
 * Convert HTML content to a basic DOCX file.
 * NOTE: this is a naive <p> extractor. If Quill's list module is in play,
 * ordered/unordered lists come through as <ol><li>/<ul><li>, not literal
 * "1." text - those are stripped here along with all other tags, which
 * means a Quill-native numbered list and a hand-typed "1. Heading" line
 * currently look identical by the time the formatter sees them. If you
 * want Quill's own list format to bypass heading detection entirely,
 * this is the place to special-case <li> before the tag-strip below.
 */
async function htmlToDocx(html) {
  const paragraphs = [];
  const paragraphRegex = /<p[^>]*>(.*?)<\/p>/gi;
  let match;

  while ((match = paragraphRegex.exec(html)) !== null) {
    const text = match[1]
      .replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#039;/g, "'")
      .trim();

    if (text) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun(text)],
        })
      );
    }
  }

  const doc = new Document({
    sections: [{
      children: paragraphs.length > 0 ? paragraphs : [
        new Paragraph({ children: [new TextRun("Empty document")] })
      ],
    }],
  });

  return await Packer.toBuffer(doc);
}

/**
 * Pulls the machine-readable "WARNINGS_JSON:[...]" line out of the
 * formatter's stdout. Returns [] if the line is missing or malformed -
 * never throws, since a parsing hiccup here shouldn't fail formatting.
 */
function extractWarnings(stdout) {
  const match = stdout.match(/WARNINGS_JSON:(\[.*\])/);
  if (!match) return [];
  try {
    return JSON.parse(match[1]);
  } catch (e) {
    console.error("Failed to parse WARNINGS_JSON:", e);
    return [];
  }
}

/**
 * Live formatting with socket updates.
 */
async function formatLive(req, res) {
  try {
    const projectId = Number(req.params.projectId);
    const userId = req.session.user.id;
    const options = req.body.options || {};

    console.log("Live formatting for project:", projectId);

    const projectRows = await query("SELECT title FROM projects WHERE id = ?", [projectId]);
    if (!projectRows || projectRows.length === 0) {
      return res.status(404).json({ error: "Project not found" });
    }

    const docRows = await query("SELECT content FROM project_documents WHERE project_id = ?", [projectId]);
    if (!docRows || docRows.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const io = req.app.get("io");
    const userRoom = `user_${userId}`;

    if (io) {
      io.to(userRoom).emit("format_progress", {
        stage: "extraction",
        message: "Converting document to DOCX format...",
        progress: 20,
      });
    }

    const pythonScript = path.join(__dirname, "..", "..", "..", "nlp", "new_formatter.py");
    const tempDir = path.join(__dirname, "..", "..", "temp");
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    ensureCleanupInterval(tempDir);

    const timestamp = Date.now();
    const inputFile = path.join(tempDir, `input_${projectId}_${timestamp}.docx`);
    const outputFile = path.join(tempDir, `output_${projectId}_${timestamp}.docx`);

    try {
      const docxBuffer = await htmlToDocx(docRows[0].content);
      fs.writeFileSync(inputFile, docxBuffer);

      if (io) {
        io.to(userRoom).emit("format_progress", {
          stage: "conversion",
          message: "Document converted successfully",
          progress: 30,
        });
      }
    } catch (convError) {
      console.error("Conversion error:", convError);
      if (io) {
        io.to(userRoom).emit("format_error", {
          error: "Failed to convert document",
          details: convError.message,
        });
      }
      return res.json({ success: false, error: "Conversion failed" });
    }

    // Build config with all styling options. Only fields StyleConfig
    // actually knows about are worth sending - anything else is ignored
    // by StyleConfig.from_dict() on the Python side but there's no
    // reason to send stale keys (e.g. enable_nlp_backup no longer exists).
    const config = {
      // Document features
      include_toc: options.include_toc !== false,
      enable_nlp_fallback: options.enable_nlp_fallback !== false,
      nlp_min_heading_score: options.nlp_min_heading_score || 0.6,

      // Font settings
      font_name: options.font_name || "Times New Roman",
      body_font_size_pt: options.body_font_size_pt || 12,
      heading_font_name: options.heading_font_name || null,

      // Heading sizes
      heading_1_size_pt: options.heading_1_size_pt || 16,
      heading_2_size_pt: options.heading_2_size_pt || 14,
      heading_3_size_pt: options.heading_3_size_pt || 13,

      // Heading colors (hex without #)
      heading_1_color: options.heading_1_color || null,
      heading_2_color: options.heading_2_color || null,

      // Spacing
      line_spacing: options.line_spacing || 1.5,
      paragraph_space_after_pt: options.paragraph_space_after_pt || 6,

      // Alignment
      body_alignment: options.body_alignment || "justify",
      enable_first_line_indent: options.enable_first_line_indent !== false,
    };

    const configFile = path.join(tempDir, `config_${projectId}_${timestamp}.json`);
    fs.writeFileSync(configFile, JSON.stringify(config));

    res.json({ success: true, message: "Formatting started" });

    if (io) {
      io.to(userRoom).emit("format_progress", {
        stage: "processing",
        message: "Running formatter...",
        progress: 40,
      });
    }

    console.log("Spawning Python process...");
    console.log("Script:", pythonScript);
    console.log("Args:", ["--input", inputFile, "--output", outputFile, "--config", configFile]);

    const python = spawn(PYTHON_BIN, [
      "-u",
      pythonScript,
      "--input", inputFile,
      "--output", outputFile,
      "--config", configFile,
    ], {
      stdio: ["pipe", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";
    let processTimedOut = false;

    python.on("error", (error) => {
      console.error("Failed to start Python process:", error);
      clearTimeout(timeout);

      if (io) {
        io.to(userRoom).emit("format_error", {
          error: "Failed to start formatter",
          details: `Could not start Python: ${error.message}`,
        });
      }

      try {
        fs.unlinkSync(inputFile);
        fs.unlinkSync(configFile);
      } catch (e) {}
    });

    // NOTE: 30s was originally sized around spaCy/NLTK model load time on
    // every request. The fallback tier is now heuristic-gated (most
    // paragraphs never touch NLTK at all) and NLTK data is expected to be
    // pre-downloaded at deploy time, not fetched per-request - so this
    // should rarely be hit. Left in as a safety net, not a normal path.
    const timeout = setTimeout(() => {
      processTimedOut = true;
      console.log("Python process timeout - killing process");
      python.kill();

      if (io) {
        io.to(userRoom).emit("format_error", {
          error: "Formatting timeout",
          details: "Process took too long (>30 seconds). The document may be too complex.",
        });
      }

      try {
        fs.unlinkSync(inputFile);
        fs.unlinkSync(configFile);
      } catch (e) {}
    }, 30000);

    python.stdout.on("data", (data) => {
      const output = data.toString();
      stdout += output;

      const isMarkerLine = output.includes("PREVIEW_HTML_b64:") || output.includes("WARNINGS_JSON:");
      if (!isMarkerLine) {
        console.log("Python stdout:", output);
      }
      if (io && !isMarkerLine) {
        io.to(userRoom).emit("format_progress", {
          stage: "processing",
          message: output.trim(),
          progress: 60,
        });
      }
    });

    python.stderr.on("data", (data) => {
      stderr += data.toString();
      const errMsg = data.toString();
      console.error("Python stderr:", errMsg);
    });

    python.on("close", async (code) => {
      clearTimeout(timeout);

      if (processTimedOut) {
        console.log("Process already timed out");
        return;
      }

      console.log("Python process closed with code:", code);

      if (!fs.existsSync(outputFile)) {
        console.error("Python formatter exit code:", code);
        console.error("Python stderr:", stderr);

        if (io) {
          io.to(userRoom).emit("format_error", {
            error: "Output file not created",
            details: stderr || "Python script completed but no output file was generated",
          });
        }
        try {
          fs.unlinkSync(inputFile);
          fs.unlinkSync(configFile);
        } catch (e) {}
        return;
      }

      const warnings = extractWarnings(stdout);

      const docId = `formatted_${projectId}_${timestamp}`;
      global.formattedDocs = global.formattedDocs || {};
      global.formattedDocs[docId] = outputFile;

      let previewHtml = "";
      try {
        const docxBuffer = fs.readFileSync(outputFile);
        const result = await mammoth.convertToHtml({ buffer: docxBuffer }, {
          styleMap: [
            "p[style-name='Heading 1'] => h1:fresh",
            "p[style-name='Heading 2'] => h2:fresh",
            "p[style-name='Heading 3'] => h3:fresh",
          ],
        });
        previewHtml = result.value;

        const warningsHtml = warnings.length > 0
          ? `<div class="alert alert-warning mb-3">
               <strong>⚠️ ${warnings.length} heading${warnings.length > 1 ? "s" : ""} need confirmation:</strong>
               <ul class="mb-0">${warnings.map(w => `<li>${escapeHtml(w)}</li>`).join("")}</ul>
             </div>`
          : "";

        // Mirror the ACTUAL document's fonts/sizes/colors instead of a
        // hardcoded blue Georgia theme unrelated to what the user chose -
        // this is what was making the preview always look blue regardless
        // of the "Default (Black)" heading color setting.
        const bodyFont = config.font_name || "Times New Roman";
        const headingFont = config.heading_font_name || bodyFont;
        const h1Color = config.heading_1_color ? `#${config.heading_1_color}` : "#000000";
        const h2Color = config.heading_2_color ? `#${config.heading_2_color}` : "#000000";
        const h1Size = config.heading_1_size_pt || 16;
        const h2Size = config.heading_2_size_pt || 14;
        const h3Size = config.heading_3_size_pt || 13;
        const bodySize = config.body_font_size_pt || 12;

        previewHtml = `
          <style>
            h1, h2, h3 {
              font-family: '${headingFont}', serif;
              font-weight: bold;
              margin-top: 1.5em;
              margin-bottom: 0.5em;
            }
            h1 { font-size: ${h1Size}pt; color: ${h1Color}; }
            h2 { font-size: ${h2Size}pt; color: ${h2Color}; }
            h3 { font-size: ${h3Size}pt; color: #000000; }
            p {
              font-size: ${bodySize}pt;
              line-height: ${config.line_spacing || 1.5};
              margin-bottom: 1em;
              font-family: '${bodyFont}', serif;
              text-align: ${config.body_alignment || "justify"};
            }
            ul, ol { margin-left: 2em; margin-bottom: 1em; }
            li { margin-bottom: 0.5em; }
          </style>
          <div class="alert alert-success mb-3">
            <strong>✅ Formatting Complete!</strong><br>
            Review the preview below${warnings.length > 0 ? " and the warnings above" : ""}, then click "Download DOCX" to save.
          </div>
          ${warningsHtml}
          ${previewHtml}
        `;
      } catch (previewError) {
        console.error("Preview generation error:", previewError);
        previewHtml = `
          <div class="alert alert-success">
            <h3>✅ Formatting Complete!</h3>
            <p>Click "Download DOCX" to save the formatted document.</p>
          </div>
        `;
      }

      if (io) {
        io.to(userRoom).emit("format_complete", {
          documentId: docId,
          preview: previewHtml,
          warnings: warnings,
          message: warnings.length > 0
            ? `Formatting complete with ${warnings.length} item(s) to review.`
            : "Formatting complete!",
        });
      }

      try {
        fs.unlinkSync(inputFile);
        fs.unlinkSync(configFile);
      } catch (e) {}
    });

  } catch (error) {
    console.error("Live formatting error:", error);
    return res.status(500).json({
      error: "Formatting failed",
      details: error.message,
    });
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Download formatted document.
 */
async function downloadFormatted(req, res) {
  try {
    const docId = req.params.docId;

    global.formattedDocs = global.formattedDocs || {};
    const filePath = global.formattedDocs[docId];

    if (!filePath || !fs.existsSync(filePath)) {
      return res.status(404).json({ error: "Formatted document not found" });
    }

    // NOTE: this used to delete the file 5 seconds after every single
    // download, which is what made "click Download again" fail with a
    // confusing 404 - it wasn't a deliberate one-download limit, just
    // overly aggressive cleanup. Deletion is now handled by a periodic
    // age-based sweep (see ensureCleanupInterval) instead, so the same
    // file can be downloaded as many times as needed within its
    // lifetime window.
    res.download(filePath, "formatted_document.docx", (err) => {
      if (err) {
        console.error("Download error:", err);
      }
    });

  } catch (error) {
    console.error("Download error:", error);
    return res.status(500).json({
      error: "Download failed",
      details: error.message,
    });
  }
}

module.exports = {
  formatLive,
  downloadFormatted,
};