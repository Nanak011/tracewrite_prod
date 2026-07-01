const { query } = require("../../db");
const { colorFromUserId, stripHtml } = require("../serverHelpers");

function countWords(text = "") {
  const clean = String(text || "").replace(/&nbsp;/gi, " ").replace(/\s+/g, " ").trim();
  return clean ? clean.split(/\s+/).length : 0;
}

function normalizeColor(value = "") {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, "");
}

function hexToRgb(hex) {
  const clean = String(hex || "").replace("#", "").trim();
  if (clean.length !== 6) return "";
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  if ([r, g, b].some((n) => Number.isNaN(n))) return "";
  return `rgb(${r},${g},${b})`;
}

function hslToRgb(hsl) {
  const match = String(hsl || "")
    .trim()
    .toLowerCase()
    .match(/hsla?\(\s*([\d.]+)\s*,\s*([\d.]+)%\s*,\s*([\d.]+)%/i);
  if (!match) return "";

  const hue = Number(match[1]) % 360;
  const saturation = Number(match[2]) / 100;
  const lightness = Number(match[3]) / 100;

  if ([hue, saturation, lightness].some((n) => Number.isNaN(n))) return "";

  const c = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = c * (1 - Math.abs(((hue / 60) % 2) - 1));
  const m = lightness - c / 2;

  let r1 = 0;
  let g1 = 0;
  let b1 = 0;

  if (hue < 60) {
    r1 = c; g1 = x; b1 = 0;
  } else if (hue < 120) {
    r1 = x; g1 = c; b1 = 0;
  } else if (hue < 180) {
    r1 = 0; g1 = c; b1 = x;
  } else if (hue < 240) {
    r1 = 0; g1 = x; b1 = c;
  } else if (hue < 300) {
    r1 = x; g1 = 0; b1 = c;
  } else {
    r1 = c; g1 = 0; b1 = x;
  }

  const r = Math.round((r1 + m) * 255);
  const g = Math.round((g1 + m) * 255);
  const b = Math.round((b1 + m) * 255);
  return `rgb(${r},${g},${b})`;
}

function escapeRegex(text = "") {
  return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractNetWordsByAuthorColor(html, members) {
  const perUser = new Map();
  const colorToUserId = new Map();
  let workingHtml = String(html || "");
  const validUserIds = new Set(members.map((m) => Number(m.user_id)));

  members.forEach((m) => {
    const id = Number(m.user_id);
    perUser.set(id, 0);
    const color = normalizeColor(colorFromUserId(id));
    const rgb = normalizeColor(hexToRgb(color) || hslToRgb(color));
    if (color) colorToUserId.set(color, id);
    if (rgb) colorToUserId.set(rgb, id);
  });

  // Attribute caption words directly to the user who inserted them.
  const captionRegex = /<p\b[^>]*\bdata-author-id=("|')(\d+)\1[^>]*>([\s\S]*?)<\/p>/gi;
  workingHtml = workingHtml.replace(captionRegex, (full, _q, authorIdText, innerHtml) => {
    const authorId = Number(authorIdText);
    if (validUserIds.has(authorId)) {
      const words = countWords(stripHtml(innerHtml));
      if (words > 0) {
        perUser.set(authorId, Number(perUser.get(authorId) || 0) + words);
      }
    }
    return "";
  });

  const spanRegex = /<span\b[^>]*style=("|')([^"']*)\1[^>]*>([\s\S]*?)<\/span>/gi;
  let match;
  while ((match = spanRegex.exec(workingHtml)) !== null) {
    const style = String(match[2] || "");
    const inner = String(match[3] || "");
    const colorMatch = style.match(/(?:^|;)\s*color\s*:\s*([^;]+)/i);
    if (!colorMatch) continue;

    const styleColor = normalizeColor(colorMatch[1]);
    const userId = colorToUserId.get(styleColor);
    if (!userId) continue;

    const words = countWords(stripHtml(inner));
    if (!words) continue;
    perUser.set(userId, Number(perUser.get(userId) || 0) + words);
  }

  return perUser;
}

async function projectSummary(req, res) {
  try {
    const projectId = Number(req.params.projectId);

    const members = await query(
      `SELECT u.id AS user_id, u.name, u.email, pm.role
       FROM project_members pm
       JOIN users u ON u.id = pm.user_id
       WHERE pm.project_id = ?
       ORDER BY u.name ASC`,
      [projectId]
    );

    const perUser = await query(
      `SELECT u.id AS user_id, u.name,
          SUM(CASE WHEN cl.action_type = 'write' THEN cl.word_count ELSE 0 END) AS words_written,
          COUNT(*) AS edits_count,
          COUNT(DISTINCT cl.page_number) AS pages_edited,
          SUM(cl.time_spent) AS total_time_spent
       FROM contribution_logs cl
       JOIN users u ON u.id = cl.user_id
       WHERE cl.project_id = ?
       GROUP BY u.id, u.name
       ORDER BY words_written DESC`,
      [projectId]
    );

    const totals = await query(
      `SELECT
         SUM(CASE WHEN action_type = 'write' THEN word_count ELSE 0 END) AS total_words,
         COUNT(*) AS total_edits,
         SUM(time_spent) AS total_time
       FROM contribution_logs
       WHERE project_id = ?`,
      [projectId]
    );

    const totalWords = Number(totals[0]?.total_words || 0);
    const summary = perUser.map((row) => ({
      ...row,
      contribution_percent: totalWords ? Number(((Number(row.words_written || 0) / totalWords) * 100).toFixed(2)) : 0,
    }));

    const actionBreakdownRows = await query(
      `SELECT cl.user_id, u.name, cl.action_type,
          COUNT(*) AS action_count,
          SUM(cl.word_count) AS total_words,
          SUM(cl.time_spent) AS total_time
       FROM contribution_logs cl
       JOIN users u ON u.id = cl.user_id
       WHERE cl.project_id = ?
       GROUP BY cl.user_id, u.name, cl.action_type
       ORDER BY u.name ASC, cl.action_type ASC`,
      [projectId]
    );

    const actionTotals = await query(
      `SELECT action_type, COUNT(*) AS total_count, SUM(word_count) AS total_words
       FROM contribution_logs
       WHERE project_id = ?
       GROUP BY action_type
       ORDER BY total_count DESC`,
      [projectId]
    );

    const logs = await query(
      `SELECT cl.*, u.name
       FROM contribution_logs cl
       JOIN users u ON u.id = cl.user_id
       WHERE cl.project_id = ?
       ORDER BY cl.timestamp ASC`,
      [projectId]
    );

    const pageProgress = await query(
      `SELECT p.page_number,
          LENGTH(TRIM(REPLACE(REPLACE(REPLACE(p.content, '<p>', ' '), '</p>', ' '), '&nbsp;', ' '))) AS char_count,
          p.updated_at
       FROM pages p
       WHERE p.project_id = ?
       ORDER BY p.page_number ASC`,
      [projectId]
    );

    const documentRows = await query(
      "SELECT content, updated_by FROM project_documents WHERE project_id = ?",
      [projectId]
    );
    const currentContent = String(documentRows[0]?.content || "");
    const updatedByUserId = Number(documentRows[0]?.updated_by || 0);
    const currentTotalWords = countWords(stripHtml(currentContent));
    const netWordsMap = extractNetWordsByAuthorColor(currentContent, members);

    // Force full current-state accounting: assign any residual words to the latest updater.
    const initialAttributedWords = members.reduce(
      (sum, m) => sum + Number(netWordsMap.get(Number(m.user_id)) || 0),
      0
    );
    const residualWords = Math.max(0, currentTotalWords - initialAttributedWords);
    if (residualWords > 0 && members.length) {
      const validMemberIds = new Set(members.map((m) => Number(m.user_id)));
      const nonViewerMembers = members.filter((m) => String(m.role || "") !== "Viewer");
      const fallbackUserId = validMemberIds.has(updatedByUserId)
        ? updatedByUserId
        : Number((nonViewerMembers[0] || members[0]).user_id);
      netWordsMap.set(
        fallbackUserId,
        Number(netWordsMap.get(fallbackUserId) || 0) + residualWords
      );
    }

    const summaryByUserId = new Map(summary.map((row) => [Number(row.user_id), row]));
    const attributedWords = members.reduce(
      (sum, m) => sum + Number(netWordsMap.get(Number(m.user_id)) || 0),
      0
    );
    const percentBase = attributedWords || currentTotalWords || 1;
    const memberSummary = members.map((m) => {
      const s = summaryByUserId.get(Number(m.user_id));
      const netWords = Number(netWordsMap.get(Number(m.user_id)) || 0);
      return {
        user_id: m.user_id,
        name: m.name,
        email: m.email,
        role: m.role,
        words_written: Number(s?.words_written || 0),
        edits_count: Number(s?.edits_count || 0),
        pages_edited: Number(s?.pages_edited || 0),
        total_time_spent: Number(s?.total_time_spent || 0),
        contribution_percent: Number(s?.contribution_percent || 0),
        net_words: netWords,
        net_percent: Number(((netWords / percentBase) * 100).toFixed(2)),
        author_color: colorFromUserId(m.user_id),
      };
    });

    const unattributedWords = Math.max(0, currentTotalWords - attributedWords);
    const currentStateRows = memberSummary.filter((m) => String(m.role || "") !== "Viewer");

    res.json({
      members,
      summary,
      memberSummary,
      actionBreakdown: actionBreakdownRows,
      actionTotals,
      totals: totals[0] || {},
      logs,
      pageProgress,
      currentDocument: {
        total_words: currentTotalWords,
        unattributed_words: unattributedWords,
        per_user: currentStateRows.map((m) => ({
          user_id: m.user_id,
          name: m.name,
          words: Number(m.net_words || 0),
          percent: Number(m.net_percent || 0),
          color: m.author_color,
        })),
      },
      currentStateSummary: currentStateRows
        .map((m) => ({
          user_id: m.user_id,
          name: m.name,
          role: m.role,
          net_words: Number(m.net_words || 0),
          net_percent: Number(m.net_percent || 0),
        }))
        .sort((a, b) => Number(b.net_words || 0) - Number(a.net_words || 0)),
      fullHistory: {
        totals: totals[0] || {},
        per_user: summary,
        action_breakdown: actionBreakdownRows,
        logs,
      },
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}

// ─── ADDED: ACTION TRACKING FOR INDIVIDUAL LIVE CHANGES ─────────────────────
async function trackActivity(req, res) {
  try {
    const projectId = Number(req.params.projectId);
    const userId = req.user?.id || req.session?.user?.id;
    const { actions, wordsAdded, charsAdded, time } = req.body;

    // Guard checking if user id is missing
    if (!userId) {
      return res.status(401).json({ error: "Unauthorized user identity context missing" });
    }

    // Process array of string action tags (e.g., ['write', 'delete'])
    // We break them up and log them individually if multiple types occurred in a single batch update
    for (const actionType of actions) {
      let wordCount = 0;
      if (actionType === "write") {
        wordCount = Number(wordsAdded || 0);
      }

      await query(
        `INSERT INTO contribution_logs 
          (project_id, user_id, action_type, word_count, time_spent, page_number, timestamp) 
         VALUES (?, ?, ?, ?, ?, 1, NOW())`,
        [projectId, userId, actionType, wordCount, Number(time || 1)]
      );
    }

    return res.status(200).json({ status: "success" });
  } catch (err) {
    console.error("Error logging background analytics stream:", err);
    return res.status(500).json({ error: err.message });
  }
}

// Export both handling endpoints
module.exports = { projectSummary, trackActivity };